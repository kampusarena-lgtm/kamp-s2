from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file, g, jsonify, has_request_context, abort
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from pathlib import Path
from datetime import datetime, date, timedelta, timezone
from decimal import Decimal, ROUND_HALF_UP
from functools import wraps
from io import BytesIO
import sqlite3, os, json, time, threading, tempfile, subprocess, shutil, hashlib, hmac, secrets, base64, socket, sys, zipfile, re
WINDOWS_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)

app = Flask(__name__)
def installation_session_key():
    configured = os.environ.get("KGA_SECRET_KEY", "").strip()
    if configured:
        if len(configured) < 32 or configured == "kga-muhasebe-v10-change-me":
            raise RuntimeError("KGA_SECRET_KEY en az 32 karakterlik rastgele bir anahtar olmalıdır.")
        return configured
    key_path = Path(__file__).resolve().parent / "runtime" / "session.key"
    key_path.parent.mkdir(exist_ok=True)
    try:
        fd = os.open(key_path, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
    except FileExistsError:
        key = key_path.read_text(encoding="ascii").strip()
    else:
        key = secrets.token_hex(32)
        with os.fdopen(fd, "w", encoding="ascii") as target:
            target.write(key)
    if len(key) != 64:
        raise RuntimeError("Oturum anahtarı geçersiz. runtime/session.key dosyasını güvenli yedekten geri yükleyin.")
    return key

app.secret_key = installation_session_key()
app.config.update(SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE="Lax")
ADMIN_IDLE_SECONDS = 600
ADMIN_SESSIONS = {}
ADMIN_SESSION_LOCK = threading.Lock()
SCHEDULED_BACKUP_LOCK = threading.Lock()

BASE_DIR = Path(__file__).resolve().parent
DB_FILE = BASE_DIR / "KampusArena_Muhasebe.db"
EXCEL_FILE = BASE_DIR / "KampusArena_Muhasebe.xlsx"
BACKUP_DIR = BASE_DIR / "backups"
RUNTIME_DIR = BASE_DIR / "runtime"
CRYPTO_SCRIPT = BASE_DIR / "excel_crypto.ps1"
ADMIN_SECRET_SCRIPT = BASE_DIR / "admin_secret.ps1"
ADMIN_SECRET_FILE = BASE_DIR / "admin_excel_secret.dat"
BACKUP_DIR.mkdir(exist_ok=True)
RUNTIME_DIR.mkdir(exist_ok=True)

CATEGORIES = [
    "İÇECEK", "YİYECEK", "ELEKTRİK", "İNT.", "SU", "MARKET", "KİRA", "EXTRA",
    "ELEMAN", "MUHASEBE", "SİGORTA", "MUTFAK", "KAHVE", "DMCNA", "BLOKE ÇÖZÜM",
    "TEMİZLİK", "PRONET", "VERGİ", "YEMEK"
]
SYSTEM_EXPENSE_CATEGORY = "ORTAK ÖDEMESİ"
PERSONAL_CARD_CLAIM_SOURCE = "PERSONAL_CARD_CLAIM"
PERSONAL_CARD_REIMBURSEMENT_TYPE = "PARTNER_CARD_REIMBURSEMENT"
PAYMENT_SOURCES = ["İŞLETME", "ŞAHSİ KART"]
PAYMENT_METHODS = ["-", "NAKİT", "POS", "IBAN"]
CREDIT_TYPES = ["BORÇ", "ÖDEME"]
PARTNER_MOVEMENT_TYPES = ["PARA ÇIKIŞI", "ALACAK EKLE", "ALACAK ÖDEME"]
CARD_OWNERS = ["Murat", "Ahmet", "Emre"]
ADMIN_PBKDF2_ITERATIONS = 240000
LOG_LIMIT = 3000
EXCEL_SYNC_DELAY = 8.0

# V10.6.3 Personel / Puantaj / Maaş
PERSONNEL_HOURS_DIVISOR = 260.0
PERSONNEL_MEAL_BLOCK_HOURS = 10.0
PERSONNEL_MEAL_AMOUNT = 200.0
STAFF_ADVANCE_TYPE = "STAFF_ADVANCE"
STAFF_SALARY_TYPE = "STAFF_SALARY"
STAFF_MEAL_TYPE = "STAFF_MEAL"
SHIFT_DEFINITIONS = {
    "09-19": {"label": "09:00 - 19:00", "start": "09:00", "end": "19:00"},
    "13-21": {"label": "13:00 - 21:00", "start": "13:00", "end": "21:00"},
    "19-05": {"label": "19:00 - 05:00", "start": "19:00", "end": "05:00"},
    "21-09": {"label": "21:00 - 09:00", "start": "21:00", "end": "09:00"},
}

EXCEL_SYNC_LOCK = threading.Lock()
EXCEL_SYNC_GENERATION = 0
EXCEL_SYNC_TIMER = None
EXCEL_SCHEDULE_LOCK = threading.Lock()
EXCEL_PASSWORD_CACHE = None
LAST_BACKUP_TS = 0.0

ACCOUNTING_DAY_START_HOUR = 9
DAY_LOCK_HOUR = 12
PANCAFE_LIVE_DB = Path(r"C:\Program Files (x86)\Pan Group\PanCafe Pro Server\Data\USDB.DAT")
PANCAFE_EMBED_DIR = Path(r"C:\FirebirdEmbed")
PANCAFE_HELPER = BASE_DIR / "pancafe_total.ps1"
PANCAFE_RETRY_SECONDS = 300
PANCAFE_AUTO_LOCK = threading.Lock()
DAY_AUTO_LOCK = threading.RLock()
LAST_PANCAFE_AUTO_ATTEMPT = {}


def now_iso():
    return datetime.now().isoformat(timespec="seconds")


def accounting_day(now=None):
    """PanCafe muhasebe günü 09:00'da başlar ve ertesi gün 09:00'da biter."""
    now = now or datetime.now()
    d = now.date()
    if now.hour < ACCOUNTING_DAY_START_HOUR:
        d -= timedelta(days=1)
    return d


def accounting_day_iso(now=None):
    return accounting_day(now).isoformat()


def accounting_bounds(day_value):
    d = datetime.strptime(parse_day(day_value), "%Y-%m-%d")
    start = d.replace(hour=ACCOUNTING_DAY_START_HOUR, minute=0, second=0, microsecond=0)
    end = start + timedelta(days=1)
    return start, end


def entry_grace_open(tx_date, now=None):
    """Only the just-completed 09:00 accounting day may be edited before noon."""
    now = now or datetime.now()
    return (
        parse_day(tx_date) == (now.date()-timedelta(days=1)).isoformat()
        and ACCOUNTING_DAY_START_HOUR <= now.hour < DAY_LOCK_HOUR
    )


def cash_entry_default_date(now=None):
    now = now or datetime.now()
    previous = (now.date()-timedelta(days=1)).isoformat()
    return previous if entry_grace_open(previous, now) else accounting_day_iso(now)


def money(v):
    """Accept dot decimals or Turkish grouped amounts; reject invalid input."""
    raw = str(v or "").strip()
    if not raw:
        return 0.0
    import re
    if "," in raw:
        if not re.fullmatch(r"[+]?(?:[0-9]+|[0-9]{1,3}(?:\.[0-9]{3})+),[0-9]{1,2}", raw):
            raise RuntimeError("Tutar geçersiz. Örnek: 1234,56 veya 1234.56.")
        raw = raw.replace(".", "").replace(",", ".")
    elif not re.fullmatch(r"[+]?[0-9]+(?:\.[0-9]{1,2})?", raw):
        raise RuntimeError("Tutar pozitif veya sıfır olmalı ve en fazla iki ondalık içermeli.")
    value = Decimal(raw)
    if not value.is_finite() or value < 0 or value > Decimal("999999999.99"):
        raise RuntimeError("Tutar 0 ile 999.999.999,99 TL arasında olmalı.")
    return float(value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def parse_day(v):
    if isinstance(v, date):
        return v.isoformat()
    raw = str(v or "").strip()
    datetime.strptime(raw, "%Y-%m-%d")
    return raw


def month_bounds(month):
    if not month or len(month) != 7:
        month = accounting_day().strftime("%Y-%m")
    y, m = map(int, month.split("-"))
    first = date(y, m, 1)
    nxt = date(y + (m == 12), 1 if m == 12 else m + 1, 1)
    return first.isoformat(), (nxt - timedelta(days=1)).isoformat()


def default_range():
    start, end = month_bounds(accounting_day().strftime("%Y-%m"))
    return start, end


def cash_calc(nakit, pos, iban, hatali, indirim, new_debt, old_debt_paid):
    return float(sum((Decimal(str(v)) for v in (nakit, pos, iban, hatali, indirim, new_debt)), Decimal("0")) - Decimal(str(old_debt_paid)))


def cash_status(pancafe, calc):
    if pancafe is None:
        return None, "PANCAFE BEKLİYOR"
    diff = calc - pancafe
    return diff, "EŞİT" if abs(diff) < 0.01 else ("FAZLA" if diff > 0 else "EKSİK")


DB_WRITE_LOCK = threading.RLock()


def _is_write_sql(sql):
    s = (sql or "").lstrip().upper()
    return s.startswith((
        "INSERT", "UPDATE", "DELETE", "REPLACE",
        "CREATE", "ALTER", "DROP", "VACUUM",
        "BEGIN", "COMMIT", "ROLLBACK"
    ))


class SerializedConnection(sqlite3.Connection):
    """Serialize SQLite writers inside this Flask process.

    Reads remain concurrent under WAL. The global RLock is acquired on the
    first write in a transaction and released by commit/rollback/close.
    """

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._kga_write_lock_held = False

    def _acquire_write(self):
        if not self._kga_write_lock_held:
            DB_WRITE_LOCK.acquire()
            self._kga_write_lock_held = True

    def _release_write(self):
        if self._kga_write_lock_held:
            self._kga_write_lock_held = False
            try:
                DB_WRITE_LOCK.release()
            except RuntimeError:
                pass

    def execute(self, sql, parameters=()):
        write = _is_write_sql(sql)
        acquired_now = False
        if write and not self._kga_write_lock_held:
            self._acquire_write()
            acquired_now = True

        last = None
        for attempt in range(8):
            try:
                return super().execute(sql, parameters)
            except sqlite3.OperationalError as e:
                last = e
                if "locked" not in str(e).lower() and "busy" not in str(e).lower():
                    if acquired_now and not self.in_transaction:
                        self._release_write()
                    raise
                time.sleep(0.08 * (attempt + 1))

        if acquired_now and not self.in_transaction:
            self._release_write()
        raise last

    def executemany(self, sql, seq_of_parameters):
        if _is_write_sql(sql):
            self._acquire_write()
        last = None
        for attempt in range(8):
            try:
                return super().executemany(sql, seq_of_parameters)
            except sqlite3.OperationalError as e:
                last = e
                if "locked" not in str(e).lower() and "busy" not in str(e).lower():
                    raise
                time.sleep(0.08 * (attempt + 1))
        raise last

    def executescript(self, sql_script):
        # Schema/migrations are writes.
        self._acquire_write()
        last = None
        for attempt in range(8):
            try:
                return super().executescript(sql_script)
            except sqlite3.OperationalError as e:
                last = e
                if "locked" not in str(e).lower() and "busy" not in str(e).lower():
                    raise
                time.sleep(0.08 * (attempt + 1))
        raise last

    def commit(self):
        try:
            return super().commit()
        finally:
            self._release_write()

    def rollback(self):
        try:
            return super().rollback()
        finally:
            self._release_write()

    def close(self):
        try:
            return super().close()
        finally:
            self._release_write()


def db_conn():
    con = sqlite3.connect(
        DB_FILE,
        timeout=30,
        factory=SerializedConnection,
        check_same_thread=False
    )
    con.row_factory = sqlite3.Row
    if has_request_context():
        g.setdefault("db_connections", []).append(con)
    con.execute("PRAGMA foreign_keys=ON")
    con.execute("PRAGMA synchronous=NORMAL")
    con.execute("PRAGMA busy_timeout=30000")
    con.execute("PRAGMA temp_store=MEMORY")
    return con


def init_db():
    con = db_conn()
    con.execute("PRAGMA journal_mode=WAL")
    con.execute("PRAGMA cache_size=-16384")
    con.executescript("""
    CREATE TABLE IF NOT EXISTS cash_entries (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tx_date TEXT NOT NULL,
        created_at TEXT NOT NULL,
        pancafe REAL,
        nakit REAL NOT NULL DEFAULT 0,
        pos REAL NOT NULL DEFAULT 0,
        iban REAL NOT NULL DEFAULT 0,
        hatali REAL NOT NULL DEFAULT 0,
        indirim REAL NOT NULL DEFAULT 0,
        new_debt REAL NOT NULL DEFAULT 0,
        old_debt_paid REAL NOT NULL DEFAULT 0,
        note TEXT NOT NULL DEFAULT '',
        actor_name TEXT NOT NULL DEFAULT '',
        actor_type TEXT NOT NULL DEFAULT 'PERSONEL',
        active INTEGER NOT NULL DEFAULT 1,
        created_by TEXT NOT NULL DEFAULT '',
        updated_at TEXT,
        updated_by TEXT,
        canceled_at TEXT,
        canceled_by TEXT,
        cancel_reason TEXT NOT NULL DEFAULT '',
        version INTEGER NOT NULL DEFAULT 1
    );

    CREATE TABLE IF NOT EXISTS expenses (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tx_date TEXT NOT NULL,
        created_at TEXT NOT NULL,
        category TEXT NOT NULL,
        amount REAL NOT NULL,
        note TEXT NOT NULL DEFAULT '',
        payment_source TEXT NOT NULL DEFAULT 'İŞLETME',
        payment_method TEXT NOT NULL DEFAULT '-',
        card_owner TEXT NOT NULL DEFAULT '',
        system_type TEXT NOT NULL DEFAULT '',
        source_type TEXT NOT NULL DEFAULT '',
        source_id INTEGER,
        claim_category TEXT NOT NULL DEFAULT '',
        claim_origin TEXT NOT NULL DEFAULT '',
        actor_name TEXT NOT NULL DEFAULT '',
        actor_type TEXT NOT NULL DEFAULT 'PERSONEL',
        active INTEGER NOT NULL DEFAULT 1,
        created_by TEXT NOT NULL DEFAULT '',
        updated_at TEXT,
        updated_by TEXT,
        canceled_at TEXT,
        canceled_by TEXT,
        cancel_reason TEXT NOT NULL DEFAULT '',
        version INTEGER NOT NULL DEFAULT 1
    );

    CREATE TABLE IF NOT EXISTS credit_entries (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tx_date TEXT NOT NULL,
        created_at TEXT NOT NULL,
        person TEXT NOT NULL,
        movement_type TEXT NOT NULL,
        amount REAL NOT NULL,
        payment_method TEXT NOT NULL DEFAULT '-',
        note TEXT NOT NULL DEFAULT '',
        actor_name TEXT NOT NULL DEFAULT '',
        actor_type TEXT NOT NULL DEFAULT 'PERSONEL',
        active INTEGER NOT NULL DEFAULT 1,
        created_by TEXT NOT NULL DEFAULT '',
        updated_at TEXT,
        updated_by TEXT,
        canceled_at TEXT,
        canceled_by TEXT,
        cancel_reason TEXT NOT NULL DEFAULT '',
        version INTEGER NOT NULL DEFAULT 1
    );

    CREATE TABLE IF NOT EXISTS partners (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE COLLATE NOCASE,
        share_percent REAL NOT NULL DEFAULT 0,
        note TEXT NOT NULL DEFAULT '',
        active INTEGER NOT NULL DEFAULT 1,
        created_at TEXT NOT NULL,
        updated_at TEXT,
        updated_by TEXT,
        canceled_at TEXT,
        canceled_by TEXT,
        cancel_reason TEXT NOT NULL DEFAULT '',
        version INTEGER NOT NULL DEFAULT 1
    );

    CREATE TABLE IF NOT EXISTS partner_movements (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tx_date TEXT NOT NULL,
        created_at TEXT NOT NULL,
        partner_id INTEGER NOT NULL,
        movement_type TEXT NOT NULL,
        amount REAL NOT NULL,
        payment_method TEXT NOT NULL DEFAULT '-',
        payment_target TEXT NOT NULL DEFAULT 'AUTO',
        note TEXT NOT NULL DEFAULT '',
        source_type TEXT NOT NULL DEFAULT '',
        source_id INTEGER,
        actor_name TEXT NOT NULL DEFAULT '',
        actor_type TEXT NOT NULL DEFAULT 'ADMIN',
        active INTEGER NOT NULL DEFAULT 1,
        updated_at TEXT,
        updated_by TEXT,
        canceled_at TEXT,
        canceled_by TEXT,
        cancel_reason TEXT NOT NULL DEFAULT '',
        version INTEGER NOT NULL DEFAULT 1,
        FOREIGN KEY(partner_id) REFERENCES partners(id)
    );

    CREATE TABLE IF NOT EXISTS profit_distributions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        period_month TEXT NOT NULL UNIQUE,
        period_start TEXT NOT NULL,
        period_end TEXT NOT NULL,
        operating_profit REAL NOT NULL DEFAULT 0,
        carry_loss_in REAL NOT NULL DEFAULT 0,
        card_protection_in REAL NOT NULL DEFAULT 0,
        card_protection_out REAL NOT NULL DEFAULT 0,
        reserve_amount REAL NOT NULL DEFAULT 0,
        distributable_profit REAL NOT NULL DEFAULT 0,
        carry_loss_out REAL NOT NULL DEFAULT 0,
        total_share REAL NOT NULL DEFAULT 0,
        note TEXT NOT NULL DEFAULT '',
        active INTEGER NOT NULL DEFAULT 1,
        created_at TEXT NOT NULL,
        created_by TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS profit_distribution_lines (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        distribution_id INTEGER NOT NULL,
        partner_id INTEGER NOT NULL,
        share_percent REAL NOT NULL,
        amount REAL NOT NULL,
        partner_movement_id INTEGER,
        created_at TEXT NOT NULL,
        UNIQUE(distribution_id, partner_id),
        FOREIGN KEY(distribution_id) REFERENCES profit_distributions(id),
        FOREIGN KEY(partner_id) REFERENCES partners(id),
        FOREIGN KEY(partner_movement_id) REFERENCES partner_movements(id)
    );

    CREATE TABLE IF NOT EXISTS partner_payment_allocations (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        payment_movement_id INTEGER NOT NULL,
        claim_movement_id INTEGER NOT NULL,
        amount REAL NOT NULL DEFAULT 0,
        expense_id INTEGER,
        active INTEGER NOT NULL DEFAULT 1,
        created_at TEXT NOT NULL,
        updated_at TEXT,
        UNIQUE(payment_movement_id, claim_movement_id),
        FOREIGN KEY(payment_movement_id) REFERENCES partner_movements(id),
        FOREIGN KEY(claim_movement_id) REFERENCES partner_movements(id),
        FOREIGN KEY(expense_id) REFERENCES expenses(id)
    );

    CREATE TABLE IF NOT EXISTS personnel (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE COLLATE NOCASE,
        monthly_salary REAL NOT NULL DEFAULT 0,
        default_shift TEXT NOT NULL DEFAULT '',
        active INTEGER NOT NULL DEFAULT 1,
        created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS personnel_attendance (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        personnel_id INTEGER NOT NULL,
        work_date TEXT NOT NULL,
        shift_code TEXT NOT NULL,
        scheduled_start TEXT NOT NULL,
        scheduled_end TEXT NOT NULL,
        actual_start TEXT NOT NULL,
        actual_end TEXT NOT NULL,
        worked_minutes INTEGER NOT NULL DEFAULT 0,
        note TEXT NOT NULL DEFAULT '',
        active INTEGER NOT NULL DEFAULT 1,
        created_at TEXT NOT NULL,
        created_by TEXT NOT NULL DEFAULT '',
        updated_at TEXT,
        updated_by TEXT,
        canceled_at TEXT,
        canceled_by TEXT,
        cancel_reason TEXT NOT NULL DEFAULT '',
        FOREIGN KEY(personnel_id) REFERENCES personnel(id)
    );

    CREATE TABLE IF NOT EXISTS personnel_advances (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        personnel_id INTEGER NOT NULL,
        tx_date TEXT NOT NULL,
        amount REAL NOT NULL,
        payment_method TEXT NOT NULL DEFAULT 'NAKİT',
        expense_id INTEGER,
        note TEXT NOT NULL DEFAULT '',
        active INTEGER NOT NULL DEFAULT 1,
        created_at TEXT NOT NULL,
        created_by TEXT NOT NULL DEFAULT '',
        canceled_at TEXT,
        canceled_by TEXT,
        cancel_reason TEXT NOT NULL DEFAULT '',
        FOREIGN KEY(personnel_id) REFERENCES personnel(id),
        FOREIGN KEY(expense_id) REFERENCES expenses(id)
    );

    CREATE TABLE IF NOT EXISTS personnel_payments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        personnel_id INTEGER NOT NULL,
        period_month TEXT NOT NULL,
        tx_date TEXT NOT NULL,
        payment_kind TEXT NOT NULL,
        amount REAL NOT NULL,
        payment_method TEXT NOT NULL,
        expense_id INTEGER,
        note TEXT NOT NULL DEFAULT '',
        active INTEGER NOT NULL DEFAULT 1,
        created_at TEXT NOT NULL,
        created_by TEXT NOT NULL DEFAULT '',
        canceled_at TEXT,
        canceled_by TEXT,
        cancel_reason TEXT NOT NULL DEFAULT '',
        FOREIGN KEY(personnel_id) REFERENCES personnel(id),
        FOREIGN KEY(expense_id) REFERENCES expenses(id)
    );

    CREATE TABLE IF NOT EXISTS admin_accounts (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT NOT NULL UNIQUE COLLATE NOCASE,
        password_salt TEXT NOT NULL,
        password_hash TEXT NOT NULL,
        is_owner INTEGER NOT NULL DEFAULT 0,
        active INTEGER NOT NULL DEFAULT 1,
        created_at TEXT NOT NULL,
        created_by TEXT NOT NULL DEFAULT ''
    );

    CREATE TABLE IF NOT EXISTS day_closures (
        tx_date TEXT PRIMARY KEY,
        is_closed INTEGER NOT NULL DEFAULT 1,
        closed_at TEXT NOT NULL,
        closed_by TEXT NOT NULL,
        close_note TEXT NOT NULL DEFAULT '',
        reopened_at TEXT,
        reopened_by TEXT
    );

    CREATE TABLE IF NOT EXISTS audit_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        created_at TEXT NOT NULL,
        actor_name TEXT NOT NULL DEFAULT '',
        actor_type TEXT NOT NULL DEFAULT '',
        entity_type TEXT NOT NULL,
        entity_id INTEGER,
        action TEXT NOT NULL,
        summary TEXT NOT NULL DEFAULT '',
        before_json TEXT NOT NULL DEFAULT '',
        after_json TEXT NOT NULL DEFAULT ''
    );

    CREATE TABLE IF NOT EXISTS app_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        created_at TEXT NOT NULL,
        level TEXT NOT NULL,
        action TEXT NOT NULL,
        detail TEXT NOT NULL DEFAULT '',
        actor TEXT NOT NULL DEFAULT '',
        source TEXT NOT NULL DEFAULT 'SYSTEM'
    );

    CREATE TABLE IF NOT EXISTS system_meta (
        meta_key TEXT PRIMARY KEY,
        meta_value TEXT NOT NULL DEFAULT ''
    );

    CREATE INDEX IF NOT EXISTS idx_cash_date ON cash_entries(tx_date, active);
    CREATE INDEX IF NOT EXISTS idx_exp_date ON expenses(tx_date, active);
    CREATE INDEX IF NOT EXISTS idx_credit_date ON credit_entries(tx_date, active);
    CREATE INDEX IF NOT EXISTS idx_partner_mov_date ON partner_movements(tx_date, active);
    CREATE INDEX IF NOT EXISTS idx_profit_dist_month ON profit_distributions(period_month, active);
    CREATE INDEX IF NOT EXISTS idx_profit_line_dist ON profit_distribution_lines(distribution_id);
    CREATE INDEX IF NOT EXISTS idx_partner_alloc_payment ON partner_payment_allocations(payment_movement_id, active);
    CREATE INDEX IF NOT EXISTS idx_partner_alloc_claim ON partner_payment_allocations(claim_movement_id, active);
    CREATE INDEX IF NOT EXISTS idx_audit_created ON audit_log(created_at);
    CREATE INDEX IF NOT EXISTS idx_cash_pending_pancafe ON cash_entries(tx_date, pancafe, active);
    CREATE INDEX IF NOT EXISTS idx_day_closure_state ON day_closures(tx_date, is_closed);
    CREATE INDEX IF NOT EXISTS idx_app_log_created ON app_log(created_at);
    CREATE INDEX IF NOT EXISTS idx_attendance_date_person ON personnel_attendance(work_date, personnel_id, active);
    CREATE UNIQUE INDEX IF NOT EXISTS idx_attendance_person_day_active ON personnel_attendance(personnel_id, work_date) WHERE active=1;
    CREATE INDEX IF NOT EXISTS idx_advance_date_person ON personnel_advances(tx_date, personnel_id, active);
    CREATE INDEX IF NOT EXISTS idx_personnel_payment_month ON personnel_payments(period_month, personnel_id, payment_kind, active);
    """)
    # V10.5 migration: older databases are upgraded in-place.
    pm_cols = {r[1] for r in con.execute("PRAGMA table_info(partner_movements)").fetchall()}
    if "payment_target" not in pm_cols:
        con.execute("ALTER TABLE partner_movements ADD COLUMN payment_target TEXT NOT NULL DEFAULT 'AUTO'")
    if "claim_category" not in pm_cols:
        con.execute("ALTER TABLE partner_movements ADD COLUMN claim_category TEXT NOT NULL DEFAULT ''")
    if "claim_origin" not in pm_cols:
        con.execute("ALTER TABLE partner_movements ADD COLUMN claim_origin TEXT NOT NULL DEFAULT ''")

    pd_cols = {r[1] for r in con.execute("PRAGMA table_info(profit_distributions)").fetchall()}
    if "card_protection_in" not in pd_cols:
        con.execute("ALTER TABLE profit_distributions ADD COLUMN card_protection_in REAL NOT NULL DEFAULT 0")
    if "card_protection_out" not in pd_cols:
        con.execute("ALTER TABLE profit_distributions ADD COLUMN card_protection_out REAL NOT NULL DEFAULT 0")

    # V10.6.3 migration: mevcut personel listesi korunur; yalnız maaş/vardiya alanları eklenir.
    personnel_cols = {r[1] for r in con.execute("PRAGMA table_info(personnel)").fetchall()}
    if "monthly_salary" not in personnel_cols:
        con.execute("ALTER TABLE personnel ADD COLUMN monthly_salary REAL NOT NULL DEFAULT 0")
    if "default_shift" not in personnel_cols:
        con.execute("ALTER TABLE personnel ADD COLUMN default_shift TEXT NOT NULL DEFAULT ''")
    con.execute("""CREATE TABLE IF NOT EXISTS personnel_salary_history (
        personnel_id INTEGER NOT NULL REFERENCES personnel(id),
        effective_month TEXT NOT NULL,
        monthly_salary REAL NOT NULL,
        PRIMARY KEY(personnel_id, effective_month)
    )""")
    con.execute("""INSERT INTO personnel_salary_history(personnel_id,effective_month,monthly_salary)
        SELECT id,'0001-01',monthly_salary FROM personnel p
        WHERE NOT EXISTS (SELECT 1 FROM personnel_salary_history h WHERE h.personnel_id=p.id)""")
    duplicate = con.execute("SELECT tx_date FROM cash_entries WHERE active=1 GROUP BY tx_date HAVING COUNT(*)>1 LIMIT 1").fetchone()
    if duplicate:
        con.rollback(); con.close()
        raise RuntimeError(f"{duplicate[0]} için birden fazla aktif kasa kaydı var. Yedek alıp yinelenen kayıtları inceleyin; otomatik silme yapılmadı.")
    con.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_cash_one_active_day ON cash_entries(tx_date) WHERE active=1")
    con.commit(); con.close()


def rowdict(row):
    return dict(row) if row else None


def meta_get(key, default=""):
    con = db_conn()
    row = con.execute("SELECT meta_value FROM system_meta WHERE meta_key=?", (key,)).fetchone()
    con.close()
    return row[0] if row else default


def meta_set(key, value):
    con = db_conn()
    con.execute("INSERT INTO system_meta(meta_key,meta_value) VALUES(?,?) ON CONFLICT(meta_key) DO UPDATE SET meta_value=excluded.meta_value", (key, str(value)))
    con.commit(); con.close()


def add_log(level, action, detail="", actor="", source="SYSTEM"):
    try:
        con = db_conn()
        con.execute("INSERT INTO app_log(created_at,level,action,detail,actor,source) VALUES(?,?,?,?,?,?)", (now_iso(), level, action, str(detail), actor or "", source))
        con.execute("DELETE FROM app_log WHERE id NOT IN (SELECT id FROM app_log ORDER BY id DESC LIMIT ?)", (LOG_LIMIT,))
        con.commit(); con.close()
    except sqlite3.OperationalError as e:
        # Logs are diagnostic; they must never break an accounting request.
        if "locked" not in str(e).lower() and "busy" not in str(e).lower():
            pass
    except Exception:
        pass


def add_audit(entity_type, entity_id, action, summary, before=None, after=None, actor_name=None, actor_type=None):
    actor_name = actor_name if actor_name is not None else current_actor_name()
    actor_type = actor_type if actor_type is not None else current_actor_type()
    con = db_conn()
    con.execute(
        "INSERT INTO audit_log(created_at,actor_name,actor_type,entity_type,entity_id,action,summary,before_json,after_json) VALUES(?,?,?,?,?,?,?,?,?)",
        (now_iso(), actor_name or "", actor_type or "", entity_type, entity_id, action, summary, json.dumps(before or {}, ensure_ascii=False, default=str), json.dumps(after or {}, ensure_ascii=False, default=str))
    )
    con.commit(); con.close()


def current_actor_name():
    if session.get("admin"):
        return session.get("admin_username", "Yönetici")
    return ""


def current_actor_type():
    return "ADMIN" if session.get("admin") else "PERSONEL"


def is_day_closed(tx_date):
    tx_date = parse_day(tx_date)
    con = db_conn()
    row = con.execute("SELECT is_closed FROM day_closures WHERE tx_date=?", (tx_date,)).fetchone()
    con.close()
    return bool(row and row[0])


def _reopened_today_status(tx_date):
    tx_date = parse_day(tx_date)
    con = db_conn()
    row = con.execute(
        "SELECT is_closed,reopened_at FROM day_closures WHERE tx_date=?",
        (tx_date,)
    ).fetchone()
    con.close()
    if not row or row["is_closed"] or not row["reopened_at"]:
        return False
    try:
        reopened_dt = datetime.fromisoformat(str(row["reopened_at"]))
        return accounting_day(reopened_dt) == accounting_day()
    except Exception:
        return False


def day_waiting_pancafe(tx_date):
    tx_date = parse_day(tx_date)
    con = db_conn()
    n = con.execute(
        "SELECT COUNT(*) FROM cash_entries WHERE active=1 AND tx_date=? AND pancafe IS NULL",
        (tx_date,)
    ).fetchone()[0]
    con.close()
    return int(n or 0)


def ensure_day_open(tx_date):
    """09:00 accounting boundary, noon automatic lock, explicit old locks retained."""
    tx_date = parse_day(tx_date)
    today = accounting_day_iso()
    if tx_date > today:
        raise RuntimeError("Gelecek muhasebe gününe işlem girilemez.")
    if is_day_closed(tx_date):
        raise RuntimeError(
            f"{tx_date} gün kilidinde. Düzeltme gerekiyorsa Ana Yönetici geçici olarak yeniden açmalıdır."
        )
    if tx_date == today or entry_grace_open(tx_date) or _reopened_today_status(tx_date):
        return
    raise RuntimeError(
        f"{tx_date} gününün kayıt süresi, kasanın kapandığı gün saat 12:00'de sona erdi. "
        "Düzeltme gerekiyorsa Ana Yönetici günü geçici olarak yeniden açmalıdır."
    )


def ensure_pancafe_entry_allowed(tx_date):
    """Manager may complete a missing PanCafe on a past D-1 pending day."""
    tx_date = parse_day(tx_date)
    if is_day_closed(tx_date):
        raise RuntimeError(
            f"{tx_date} günü kapanmış. PanCafe değişikliği için Ana Yönetici günü yeniden açmalıdır."
        )


def day_status(tx_date):
    tx_date = parse_day(tx_date)
    con = db_conn()
    row = con.execute("SELECT * FROM day_closures WHERE tx_date=?", (tx_date,)).fetchone()
    con.close()

    d = dict(row) if row else {
        "tx_date": tx_date, "is_closed": False, "closed_at": None, "closed_by": None,
        "close_note": "", "reopened_at": None, "reopened_by": None
    }
    d["is_closed"] = bool(d.get("is_closed"))
    d["is_today"] = tx_date == accounting_day_iso()
    d["is_past"] = tx_date < accounting_day_iso()
    d["waiting"] = day_waiting_pancafe(tx_date)
    d["reopened_today"] = _reopened_today_status(tx_date)
    d["entry_grace"] = entry_grace_open(tx_date) and not d["is_closed"]
    if d["is_today"]:
        d["auto_state"] = "BUGÜN AÇIK"
        d["effective_locked"] = False
    elif d["reopened_today"]:
        d["auto_state"] = "GEÇİCİ AÇIK"
        d["effective_locked"] = False
    elif d["entry_grace"]:
        d["auto_state"] = "12:00’YE KADAR KAYIT AÇIK"
        d["effective_locked"] = False
    elif d["is_closed"]:
        d["auto_state"] = "OTOMATİK KAPALI"
        d["effective_locked"] = True
    elif d["waiting"]:
        d["auto_state"] = "PANCAFE BEKLİYOR / DİĞER İŞLEMLER KİLİTLİ"
        d["effective_locked"] = True
    elif d["is_past"]:
        d["auto_state"] = "OTOMATİK KİLİTLİ / İŞLEMSİZ"
        d["effective_locked"] = True
    else:
        d["auto_state"] = "AÇIK"
        d["effective_locked"] = False
    return d


def _transaction_dates_through(cutoff):
    """Only return past dates that still need physical closure work.

    Effective date-locking itself is date based, so already-closed historical
    days do not need to be rescanned on every background pass.
    """
    con = db_conn()
    rows = con.execute(
        """
        WITH activity(tx_date) AS (
            SELECT tx_date FROM cash_entries WHERE active=1 AND tx_date<=?
            UNION
            SELECT tx_date FROM expenses WHERE active=1 AND tx_date<=?
            UNION
            SELECT tx_date FROM credit_entries WHERE active=1 AND tx_date<=?
            UNION
            SELECT tx_date FROM partner_movements WHERE active=1 AND tx_date<=?
        )
        SELECT a.tx_date
        FROM activity a
        LEFT JOIN day_closures d ON d.tx_date=a.tx_date
        WHERE COALESCE(d.is_closed,0)=0
        ORDER BY a.tx_date
        """,
        (cutoff, cutoff, cutoff, cutoff)
    ).fetchall()
    con.close()
    return [r[0] for r in rows if r[0]]



def pancafe_auto_status():
    return {
        "live_db": str(PANCAFE_LIVE_DB),
        "embed_dir": str(PANCAFE_EMBED_DIR),
        "ready": PANCAFE_LIVE_DB.exists() and (PANCAFE_EMBED_DIR / "isql.exe").exists() and PANCAFE_HELPER.exists(),
        "last_ok": meta_get("pancafe_last_ok", "Henüz yok"),
        "last_error": meta_get("pancafe_last_error", ""),
        "last_total": meta_get("pancafe_last_total", ""),
        "last_day": meta_get("pancafe_last_day", ""),
    }


def fetch_pancafe_general_total(tx_date, force=False):
    """PanCafe raporundaki GENEL TOPLAM değerini 09:00-09:00 için salt-okuma snapshot üzerinden çeker."""
    tx_date = parse_day(tx_date)
    now_ts = time.time()
    last = LAST_PANCAFE_AUTO_ATTEMPT.get(tx_date, 0)
    if not force and now_ts - last < PANCAFE_RETRY_SECONDS:
        raise RuntimeError("PanCafe otomatik çekim kısa süre önce denendi; tekrar için birkaç dakika bekleniyor.")
    LAST_PANCAFE_AUTO_ATTEMPT[tx_date] = now_ts

    if not PANCAFE_HELPER.exists():
        raise RuntimeError("pancafe_total.ps1 bulunamadı.")
    if not PANCAFE_LIVE_DB.exists():
        raise RuntimeError(f"PanCafe veritabanı bulunamadı: {PANCAFE_LIVE_DB}")
    if not (PANCAFE_EMBED_DIR / "isql.exe").exists():
        raise RuntimeError("C:\\FirebirdEmbed\\isql.exe bulunamadı.")

    cmd = [
        "powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
        "-File", str(PANCAFE_HELPER),
        "-Day", tx_date,
        "-LiveDb", str(PANCAFE_LIVE_DB),
        "-EmbedDir", str(PANCAFE_EMBED_DIR),
        "-RuntimeDir", str(RUNTIME_DIR / "pancafe")
    ]
    p = subprocess.run(cmd, capture_output=True, text=True, timeout=120, creationflags=WINDOWS_NO_WINDOW)
    raw = (p.stdout or "").strip()
    err = (p.stderr or "").strip()
    if p.returncode != 0:
        raise RuntimeError(err or raw or f"PanCafe helper çıkış kodu: {p.returncode}")
    try:
        data = json.loads(raw.splitlines()[-1])
    except Exception:
        raise RuntimeError(f"PanCafe çıktısı çözülemedi: {raw[-800:]}")
    if not data.get("ok"):
        raise RuntimeError(data.get("error") or "PanCafe GENEL TOPLAM okunamadı.")
    total = float(data.get("total", 0))
    meta_set("pancafe_last_ok", now_iso())
    meta_set("pancafe_last_error", "")
    meta_set("pancafe_last_total", f"{total:.2f}")
    meta_set("pancafe_last_day", tx_date)
    return total, data


def auto_fill_pancafe_for_day(tx_date, force=False):
    """Tek günlük kasa kaydındaki boş PanCafe alanını otomatik GENEL TOPLAM ile doldurur."""
    tx_date = parse_day(tx_date)
    con = db_conn()
    rows = con.execute(
        "SELECT * FROM cash_entries WHERE active=1 AND tx_date=? ORDER BY id",
        (tx_date,)
    ).fetchall()
    con.close()
    pending = [r for r in rows if r["pancafe"] is None]
    if not pending:
        return {"ok": True, "changed": False, "message": "PanCafe bekleyen kayıt yok."}
    if not force:
        last = LAST_PANCAFE_AUTO_ATTEMPT.get(tx_date, 0)
        if last and time.time() - last < PANCAFE_RETRY_SECONDS:
            return {"ok": True, "changed": False, "message": "PanCafe otomatik tekrar denemesi beklemede."}
    if len(rows) != 1 or len(pending) != 1:
        msg = f"{tx_date} için {len(rows)} aktif kasa kaydı var. Otomatik PanCafe güvenlik nedeniyle uygulanmadı; günlük kasa kaydı tek olmalı."
        meta_set("pancafe_last_error", msg)
        add_log("ERROR","PANCAFE_OTOMATIK",msg,source="PANCAFE")
        return {"ok": False, "changed": False, "message": msg}

    with PANCAFE_AUTO_LOCK:
        try:
            total, info = fetch_pancafe_general_total(tx_date, force=force)
            rid = pending[0]["id"]
            con = db_conn()
            before = rowdict(con.execute("SELECT * FROM cash_entries WHERE id=?", (rid,)).fetchone())
            con.execute(
                "UPDATE cash_entries SET pancafe=?,updated_at=?,updated_by=?,version=version+1 WHERE id=? AND active=1 AND pancafe IS NULL",
                (total, now_iso(), "PANCAFE OTOMATIK", rid)
            )
            con.commit()
            after = rowdict(con.execute("SELECT * FROM cash_entries WHERE id=?", (rid,)).fetchone())
            con.close()
            add_audit(
                "CASH", rid, "PANCAFE_AUTO_SET",
                f"PanCafe GENEL TOPLAM otomatik alındı: {total:,.2f} TL | 09:00-09:00",
                before, after, actor_name="PANCAFE", actor_type="SYSTEM"
            )
            add_log("INFO","PANCAFE_OTOMATIK",f"{tx_date} GENEL TOPLAM = {total:,.2f} TL",source="PANCAFE")
            schedule_excel_sync()
            return {"ok": True, "changed": True, "total": total, "info": info}
        except Exception as e:
            msg = str(e)
            meta_set("pancafe_last_error", msg)
            add_log("ERROR","PANCAFE_OTOMATIK",f"{tx_date}: {msg}",source="PANCAFE")
            return {"ok": False, "changed": False, "message": msg}


def synchronized(lock):
    def deco(fn):
        @wraps(fn)
        def wrapped(*args, **kwargs):
            with lock:
                return fn(*args, **kwargs)
        return wrapped
    return deco


@synchronized(DAY_AUTO_LOCK)
def auto_close_due_days():
    """Background closure: 09:00 PanCafe boundaries, 12:00 entry deadline."""
    today = accounting_day()
    today_s = today.isoformat()
    cutoff = (today - timedelta(days=1)).isoformat()
    changed = []
    blocked = []

    con = db_conn()

    # V10.4 rule: current day must stay open. This also converts any old V10.2
    # manual same-day closure into today's open state on upgrade/startup.
    today_row = con.execute(
        "SELECT * FROM day_closures WHERE tx_date=?",
        (today_s,)
    ).fetchone()
    if today_row and today_row["is_closed"]:
        before = dict(today_row)
        con.execute(
            """UPDATE day_closures
               SET is_closed=0,reopened_at=?,reopened_by=?,
                   close_note=COALESCE(close_note,'') || ?
               WHERE tx_date=?""",
            (
                now_iso(), "OTOMATIK V10.6.3",
                " | V10.4: Bugün otomatik olarak açık tutulur.",
                today_s
            )
        )
        con.commit()
        after = rowdict(con.execute(
            "SELECT * FROM day_closures WHERE tx_date=?", (today_s,)
        ).fetchone())
        add_audit(
            "DAY", None, "AUTO_REOPEN_TODAY",
            f"{today_s} bugünün tarihi olduğu için V10.4 tarafından açık tutuldu.",
            before, after, actor_name="SYSTEM", actor_type="SYSTEM"
        )
        changed.append(today_s)

    con.close()

    for tx_date in _transaction_dates_through(cutoff):
        # Owner reopened this past date today: give them the whole day for corrections.
        if _reopened_today_status(tx_date):
            continue

        waiting = day_waiting_pancafe(tx_date)

        # Only the immediately completed accounting day (D-1) gets an automatic
        # PanCafe snapshot. Older incomplete days stay visible for manual force-fetch.
        if waiting and tx_date == cutoff:
            auto_fill_pancafe_for_day(tx_date, force=False)
            waiting = day_waiting_pancafe(tx_date)

        # Fetch the unchanged 09:00 PanCafe total, but do not lock during data entry.
        if entry_grace_open(tx_date):
            continue

        if waiting:
            blocked.append({"date": tx_date, "waiting": waiting})
            # Keep closure row explicitly open for status/history, but ordinary
            # operations are still frozen by ensure_day_open().
            con = db_conn()
            row = con.execute(
                "SELECT * FROM day_closures WHERE tx_date=?", (tx_date,)
            ).fetchone()
            if not row:
                con.execute(
                    """INSERT INTO day_closures(
                        tx_date,is_closed,closed_at,closed_by,close_note,reopened_at,reopened_by
                    ) VALUES(?,0,'','',?,NULL,NULL)""",
                    (
                        tx_date,
                        f"V10.4 otomatik kapanış bekliyor: {waiting} PanCafe kaydı eksik."
                    )
                )
                con.commit()
            elif not row["is_closed"]:
                con.execute(
                    "UPDATE day_closures SET close_note=? WHERE tx_date=?",
                    (
                        f"V10.4 otomatik kapanış bekliyor: {waiting} PanCafe kaydı eksik.",
                        tx_date
                    )
                )
                con.commit()
            con.close()
            continue

        if is_day_closed(tx_date):
            continue

        con = db_conn()
        before = rowdict(con.execute(
            "SELECT * FROM day_closures WHERE tx_date=?", (tx_date,)
        ).fetchone()) or {}
        con.execute(
            """INSERT INTO day_closures(
                tx_date,is_closed,closed_at,closed_by,close_note,reopened_at,reopened_by
            ) VALUES(?,1,?,?,?,NULL,NULL)
            ON CONFLICT(tx_date) DO UPDATE SET
                is_closed=1,
                closed_at=excluded.closed_at,
                closed_by=excluded.closed_by,
                close_note=excluded.close_note""",
            (
                tx_date, now_iso(), "OTOMATIK",
                "V10.6.3 otomatik gün kilidi — kasa 09:00, kayıt sonu 12:00."
            )
        )
        con.commit()
        after = rowdict(con.execute(
            "SELECT * FROM day_closures WHERE tx_date=?", (tx_date,)
        ).fetchone())
        con.close()

        add_audit(
            "DAY", None, "AUTO_CLOSE",
            f"{tx_date} otomatik olarak kilitlendi (kasa 09:00, kayıt sonu 12:00).",
            before, after, actor_name="SYSTEM", actor_type="SYSTEM"
        )
        changed.append(tx_date)

    if changed:
        try:
            backup_if_due(True)
        except Exception:
            pass
        try:
            schedule_excel_sync()
        except Exception:
            pass

    return {"closed_or_changed": changed, "blocked": blocked, "cutoff": cutoff}


# ----------------------- Personnel / Admin -----------------------

def record_salary_rate(con, personnel_id, salary):
    con.execute("""INSERT INTO personnel_salary_history(personnel_id,effective_month,monthly_salary)
        VALUES(?,?,?) ON CONFLICT(personnel_id,effective_month) DO UPDATE SET monthly_salary=excluded.monthly_salary""",
        (personnel_id, accounting_day().strftime("%Y-%m"), salary))


def personnel_rows(active_only=True):
    con = db_conn()
    sql = "SELECT * FROM personnel" + (" WHERE active=1" if active_only else "") + " ORDER BY name COLLATE NOCASE"
    rows = [dict(r) for r in con.execute(sql).fetchall()]
    con.close(); return rows


def personnel_names():
    return [r["name"] for r in personnel_rows(True)]


def validate_personnel(name):
    name = (name or "").strip()
    if not name:
        raise RuntimeError("Personel seçmeden işlem yapılamaz.")
    valid = {x.casefold(): x for x in personnel_names()}
    if name.casefold() not in valid:
        raise RuntimeError("Seçilen personel aktif değil.")
    return valid[name.casefold()]


def get_personnel_by_name(name):
    name = validate_personnel(name)
    con = db_conn()
    row = con.execute("SELECT * FROM personnel WHERE active=1 AND name=? COLLATE NOCASE", (name,)).fetchone()
    con.close()
    return dict(row) if row else None


def _clock_minutes(start_text, end_text):
    """Return elapsed minutes. End <= start means next calendar day."""
    try:
        sh, sm = [int(x) for x in str(start_text).split(":", 1)]
        eh, em = [int(x) for x in str(end_text).split(":", 1)]
        if not (0 <= sh <= 23 and 0 <= eh <= 23 and 0 <= sm <= 59 and 0 <= em <= 59):
            raise ValueError
    except Exception:
        raise RuntimeError("Giriş/çıkış saati geçersiz.")
    start = sh * 60 + sm
    end = eh * 60 + em
    if end <= start:
        end += 24 * 60
    elapsed = end - start
    if elapsed <= 0 or elapsed > 18 * 60:
        raise RuntimeError("Çalışma süresi 0 ile 18 saat arasında olmalıdır. Saatleri kontrol edin.")
    return elapsed


def shift_info(code):
    code = (code or "").strip()
    if code not in SHIFT_DEFINITIONS:
        raise RuntimeError("Geçerli bir vardiya seçin.")
    return SHIFT_DEFINITIONS[code]


def personnel_month_summary(month, connection=None):
    """Fast monthly payroll summary: 4 indexed grouped queries, no N+1 scans."""
    start, end = month_bounds(month)
    con = connection or db_conn()
    people = [dict(r) for r in con.execute(
        """SELECT p.id,p.name,p.default_shift,p.active,
          COALESCE((SELECT h.monthly_salary FROM personnel_salary_history h
            WHERE h.personnel_id=p.id AND h.effective_month<=?
            ORDER BY h.effective_month DESC LIMIT 1),0) monthly_salary
          FROM personnel p WHERE p.active=1
          OR EXISTS(SELECT 1 FROM personnel_attendance a WHERE a.personnel_id=p.id AND a.active=1 AND a.work_date BETWEEN ? AND ?)
          OR EXISTS(SELECT 1 FROM personnel_advances v WHERE v.personnel_id=p.id AND v.active=1 AND v.tx_date BETWEEN ? AND ?)
          OR EXISTS(SELECT 1 FROM personnel_payments x WHERE x.personnel_id=p.id AND x.active=1 AND x.period_month=?)
          ORDER BY p.name COLLATE NOCASE""", (month,start,end,start,end,month)
    ).fetchall()]
    attendance = {int(r["personnel_id"]): int(r["mins"] or 0) for r in con.execute(
        "SELECT personnel_id,COALESCE(SUM(worked_minutes),0) mins FROM personnel_attendance WHERE active=1 AND work_date BETWEEN ? AND ? GROUP BY personnel_id",
        (start, end)
    ).fetchall()}
    advances = {int(r["personnel_id"]): float(r["total"] or 0) for r in con.execute(
        "SELECT personnel_id,COALESCE(SUM(amount),0) total FROM personnel_advances WHERE active=1 AND tx_date BETWEEN ? AND ? GROUP BY personnel_id",
        (start, end)
    ).fetchall()}
    paid = {}
    for r in con.execute(
        "SELECT personnel_id,payment_kind,COALESCE(SUM(amount),0) total FROM personnel_payments WHERE active=1 AND period_month=? GROUP BY personnel_id,payment_kind",
        (month,)
    ).fetchall():
        paid[(int(r["personnel_id"]), str(r["payment_kind"]))] = float(r["total"] or 0)
    if connection is None:
        con.close()

    rows = []
    for p in people:
        pid = int(p["id"])
        mins = attendance.get(pid, 0)
        hours = mins / 60.0
        salary_base = float(p.get("monthly_salary") or 0)
        salary_earned = float((Decimal(str(salary_base)) * Decimal(mins) / Decimal("60") / Decimal(str(PERSONNEL_HOURS_DIVISOR))).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))
        meal_units = mins // int(PERSONNEL_MEAL_BLOCK_HOURS * 60)
        meal_earned = float(meal_units) * PERSONNEL_MEAL_AMOUNT
        advance_total = advances.get(pid, 0.0)
        salary_paid = paid.get((pid, "MAAŞ"), 0.0)
        meal_paid = paid.get((pid, "YEMEK"), 0.0)
        salary_balance = salary_earned - advance_total - salary_paid
        meal_balance = meal_earned - meal_paid
        rows.append({
            **p,
            "worked_minutes": mins,
            "worked_hours": round(hours, 2),
            "salary_earned": salary_earned,
            "meal_units": int(meal_units),
            "meal_earned": round(meal_earned, 2),
            "advance_total": round(advance_total, 2),
            "salary_paid": round(salary_paid, 2),
            "meal_paid": round(meal_paid, 2),
            "salary_due": round(max(salary_balance, 0.0), 2),
            "meal_due": round(max(meal_balance, 0.0), 2),
            "advance_excess": round(max(-salary_balance, 0.0), 2),
        })
    return rows


def personnel_month_totals(rows):
    keys = ("worked_hours", "salary_earned", "meal_earned", "advance_total", "salary_paid", "meal_paid", "salary_due", "meal_due")
    out = {k: round(sum(float(r.get(k, 0) or 0) for r in rows), 2) for k in keys}
    out["people"] = len(rows)
    return out


def personnel_attendance_rows(month, personnel_id=None):
    start, end = month_bounds(month)
    con = db_conn()
    args = [start, end]
    sql = """SELECT a.*,p.name personnel_name
             FROM personnel_attendance a JOIN personnel p ON p.id=a.personnel_id
             WHERE a.active=1 AND a.work_date BETWEEN ? AND ?"""
    if personnel_id:
        sql += " AND a.personnel_id=?"
        args.append(int(personnel_id))
    sql += " ORDER BY a.work_date DESC,a.id DESC"
    rows = [dict(r) for r in con.execute(sql, args).fetchall()]
    con.close()
    for r in rows:
        r["worked_hours"] = round(float(r["worked_minutes"] or 0) / 60.0, 2)
        scheduled = _clock_minutes(r["scheduled_start"], r["scheduled_end"])
        r["difference_hours"] = round((int(r["worked_minutes"] or 0) - scheduled) / 60.0, 2)
        r["shift_label"] = SHIFT_DEFINITIONS.get(r["shift_code"], {}).get("label", r["shift_code"])
    return rows


def personnel_financial_rows(month):
    start, end = month_bounds(month)
    con = db_conn()
    advances = [dict(r) for r in con.execute(
        """SELECT a.*,p.name personnel_name FROM personnel_advances a
           JOIN personnel p ON p.id=a.personnel_id
           WHERE a.active=1 AND a.tx_date BETWEEN ? AND ? ORDER BY a.tx_date DESC,a.id DESC""",
        (start, end)
    ).fetchall()]
    payments = [dict(r) for r in con.execute(
        """SELECT x.*,p.name personnel_name FROM personnel_payments x
           JOIN personnel p ON p.id=x.personnel_id
           WHERE x.active=1 AND x.period_month=? ORDER BY x.tx_date DESC,x.id DESC""",
        (month,)
    ).fetchall()]
    con.close()
    return advances, payments


def current_personnel_due(month, personnel_id, connection=None):
    for row in personnel_month_summary(month, connection):
        if int(row["id"]) == int(personnel_id):
            return row
    raise RuntimeError("Personelin bu dönem için kaydı bulunamadı.")


def password_digest(password, salt_b64=None):
    if salt_b64:
        salt = base64.b64decode(salt_b64.encode("ascii"))
    else:
        salt = secrets.token_bytes(16); salt_b64 = base64.b64encode(salt).decode("ascii")
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, ADMIN_PBKDF2_ITERATIONS)
    return salt_b64, base64.b64encode(digest).decode("ascii")


def admin_accounts_exist():
    con = db_conn(); n = con.execute("SELECT COUNT(*) FROM admin_accounts WHERE active=1").fetchone()[0]; con.close(); return n > 0


def get_admin_accounts():
    con = db_conn(); rows = [dict(r) for r in con.execute("SELECT id,username,is_owner,active,created_at,created_by FROM admin_accounts ORDER BY is_owner DESC, username COLLATE NOCASE").fetchall()]; con.close(); return rows


def create_admin(username, password, owner=False, created_by=""):
    username = (username or "").strip()
    if len(username) < 2: raise RuntimeError("Kullanıcı adı en az 2 karakter olmalı.")
    if len(password or "") < 8: raise RuntimeError("Şifre en az 8 karakter olmalı.")
    salt, digest = password_digest(password)
    con = db_conn()
    row = con.execute("SELECT id,active FROM admin_accounts WHERE username=? COLLATE NOCASE", (username,)).fetchone()
    if row and row[1]: con.close(); raise RuntimeError("Bu yönetici hesabı zaten aktif.")
    if row:
        con.execute("UPDATE admin_accounts SET password_salt=?,password_hash=?,active=1,created_at=?,created_by=? WHERE id=?", (salt,digest,now_iso(),created_by,row[0])); aid=row[0]
    else:
        cur = con.execute("INSERT INTO admin_accounts(username,password_salt,password_hash,is_owner,active,created_at,created_by) VALUES(?,?,?,?,1,?,?)", (username,salt,digest,1 if owner else 0,now_iso(),created_by)); aid=cur.lastrowid
    con.commit(); con.close(); return aid


def authenticate_admin(username, password):
    con = db_conn(); row = con.execute("SELECT * FROM admin_accounts WHERE username=? COLLATE NOCASE", ((username or "").strip(),)).fetchone(); con.close()
    if not row or not row["active"]: return None
    _, digest = password_digest(password or "", row["password_salt"])
    if not hmac.compare_digest(digest, row["password_hash"]): return None
    return dict(row)


def admin_required(fn):
    @wraps(fn)
    def wrapped(*args, **kwargs):
        if not session.get("admin") or not session.get("admin_id"):
            session.clear(); return redirect(url_for("admin_login"))
        con = db_conn(); row = con.execute("SELECT active,is_owner,username FROM admin_accounts WHERE id=?", (session["admin_id"],)).fetchone(); con.close()
        if not row or not row[0]:
            session.clear(); flash("Yönetici hesabı aktif değil.", "danger"); return redirect(url_for("admin_login"))
        session["admin_owner"] = bool(row["is_owner"])
        session["admin_username"] = row["username"]
        return fn(*args, **kwargs)
    return wrapped


def owner_required(fn):
    @wraps(fn)
    @admin_required
    def wrapped(*args, **kwargs):
        if not session.get("admin_owner"):
            flash("Bu işlem yalnız ana yöneticiye açık.", "danger")
            return redirect(request.referrer or url_for("admin"))
        return fn(*args, **kwargs)
    return wrapped


# ----------------------- Partner accounting -----------------------

def ensure_partner(name):
    name = (name or "").strip()
    if not name: raise RuntimeError("Ortak adı boş olamaz.")
    con = db_conn(); row = con.execute("SELECT * FROM partners WHERE name=? COLLATE NOCASE", (name,)).fetchone()
    if row:
        if not row["active"]:
            con.execute("UPDATE partners SET active=1,canceled_at=NULL,canceled_by=NULL,cancel_reason='',updated_at=?,updated_by=? WHERE id=?", (now_iso(), current_actor_name(), row["id"])); con.commit()
        pid = row["id"]
    else:
        cur = con.execute("INSERT INTO partners(name,share_percent,note,active,created_at) VALUES(?,0,?,1,?)", (name,"Şahsi kart giderinden otomatik oluşturuldu.",now_iso())); pid=cur.lastrowid; con.commit()
    con.close(); return pid


def partner_balance(partner_id, until_date=None, exclude_movement_id=None):
    clauses = ["active=1", "partner_id=?"]
    args = [partner_id]
    if until_date:
        clauses.append("tx_date<=?"); args.append(parse_day(until_date))
    if exclude_movement_id:
        clauses.append("id<>?"); args.append(exclude_movement_id)
    con = db_conn()
    row = con.execute(f"SELECT COALESCE(SUM(CASE WHEN movement_type='ALACAK EKLE' THEN amount WHEN movement_type='ALACAK ÖDEME' THEN -amount ELSE 0 END),0) FROM partner_movements WHERE {' AND '.join(clauses)}", args).fetchone()
    con.close(); return float(row[0] or 0)



def _claim_bucket(source_type):
    source_type = (source_type or "").upper()
    if source_type in ("EXPENSE", PERSONAL_CARD_CLAIM_SOURCE):
        return "MASRAF"
    if source_type == "PROFIT_DISTRIBUTION":
        return "KAR_PAYI"
    return "DIGER"


def partner_claim_breakdown(partner_id, until_date=None, exclude_movement_id=None):
    """Outstanding partner receivables split into expense / profit-share / other buckets.

    Payments are applied chronologically. A payment can target MASRAF, KAR_PAYI,
    or AUTO (oldest outstanding claim first).
    """
    clauses = ["active=1", "partner_id=?", "movement_type IN ('ALACAK EKLE','ALACAK ÖDEME')"]
    args = [partner_id]
    if until_date:
        clauses.append("tx_date<=?")
        args.append(parse_day(until_date))
    if exclude_movement_id:
        clauses.append("id<>?")
        args.append(exclude_movement_id)

    con = db_conn()
    rows = [dict(r) for r in con.execute(
        f"""SELECT id,tx_date,movement_type,amount,source_type,payment_target
            FROM partner_movements
            WHERE {' AND '.join(clauses)}
            ORDER BY tx_date,id""", args
    ).fetchall()]
    con.close()

    claims = []
    paid_total = 0.0
    unallocated_payment = 0.0

    for r in rows:
        if r["movement_type"] == "ALACAK EKLE":
            claims.append({
                "id": r["id"],
                "bucket": _claim_bucket(r.get("source_type")),
                "remaining": float(r["amount"] or 0),
            })
            continue

        payment = float(r["amount"] or 0)
        paid_total += payment
        target = (r.get("payment_target") or "AUTO").upper()

        if target == "MASRAF":
            eligible = [c for c in claims if c["bucket"] == "MASRAF" and c["remaining"] > 0.0001]
        elif target == "KAR_PAYI":
            eligible = [c for c in claims if c["bucket"] == "KAR_PAYI" and c["remaining"] > 0.0001]
        else:
            eligible = [c for c in claims if c["remaining"] > 0.0001]

        left = payment
        for c in eligible:
            if left <= 0.0001:
                break
            used = min(c["remaining"], left)
            c["remaining"] -= used
            left -= used
        if left > 0.0001:
            unallocated_payment += left

    expense_left = sum(c["remaining"] for c in claims if c["bucket"] == "MASRAF")
    profit_left = sum(c["remaining"] for c in claims if c["bucket"] == "KAR_PAYI")
    other_left = sum(c["remaining"] for c in claims if c["bucket"] == "DIGER")
    claim_add = sum(
        float(r["amount"] or 0) for r in rows if r["movement_type"] == "ALACAK EKLE"
    )

    return {
        "expense_left": round(expense_left, 2),
        "profit_left": round(profit_left, 2),
        "other_left": round(other_left, 2),
        "total_left": round(expense_left + profit_left + other_left, 2),
        "claim_add": round(claim_add, 2),
        "claim_paid": round(paid_total, 2),
        "unallocated_payment": round(unallocated_payment, 2),
    }


def profit_distribution_history(limit=24):
    con = db_conn()
    rows = [dict(r) for r in con.execute(
        """SELECT * FROM profit_distributions
           WHERE active=1 ORDER BY period_month DESC LIMIT ?""",
        (limit,)
    ).fetchall()]
    con.close()
    return rows


def previous_profit_loss_carry(period_month):
    con = db_conn()
    row = con.execute(
        """SELECT carry_loss_out FROM profit_distributions
           WHERE active=1 AND period_month<? ORDER BY period_month DESC LIMIT 1""",
        (period_month,)
    ).fetchone()
    con.close()
    return float(row[0] or 0) if row else 0.0


def previous_card_protection(period_month):
    con = db_conn()
    row = con.execute(
        """SELECT card_protection_out FROM profit_distributions
           WHERE active=1 AND period_month<? ORDER BY period_month DESC LIMIT 1""",
        (period_month,)
    ).fetchone()
    con.close()
    return float(row[0] or 0) if row else 0.0


def pending_personal_card_total(until_date=None):
    """Unpaid personal-card purchases that still belong to partners."""
    con = db_conn()
    partners = [r[0] for r in con.execute("SELECT id FROM partners WHERE active=1").fetchall()]
    con.close()
    return round(sum(partner_claim_breakdown(pid, until_date=until_date)["expense_left"] for pid in partners), 2)


def period_open_days(start, end):
    con = db_conn()
    days = set()
    for table in ("cash_entries", "expenses", "credit_entries", "partner_movements"):
        for r in con.execute(
            f"SELECT DISTINCT tx_date FROM {table} WHERE active=1 AND tx_date BETWEEN ? AND ?",
            (start, end)
        ).fetchall():
            days.add(r[0])
    closures = {
        r["tx_date"]: bool(r["is_closed"])
        for r in con.execute(
            "SELECT tx_date,is_closed FROM day_closures WHERE tx_date BETWEEN ? AND ?",
            (start, end)
        ).fetchall()
    }
    con.close()
    return sorted(d for d in days if not closures.get(d, False))


def _money_decimal(v):
    return Decimal(str(round(float(v or 0), 2))).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def profit_distribution_preview(period_month, reserve_amount=0):
    start, end = month_bounds(period_month)
    p = profit_loss(start, end)
    partners = partner_data()
    total_share = round(sum(float(x["share_percent"] or 0) for x in partners), 4)
    carry_in = previous_profit_loss_carry(period_month)
    card_protection_in = previous_card_protection(period_month)
    card_protection_out = pending_personal_card_total(end)
    card_protection_change = card_protection_out - card_protection_in
    # Outstanding personal-card debt is protected from distribution. Only the
    # change versus the previous finalized month is applied, so the same debt
    # is not deducted again every month. When a protected debt is reimbursed,
    # the released protection offsets that reimbursement's cash-basis expense.
    adjusted = float(p["operating_profit"]) - carry_in - card_protection_change
    max_reserve = max(0.0, adjusted)
    reserve = max(0.0, money(reserve_amount))
    reserve_valid = reserve <= max_reserve + 0.01
    if not reserve_valid:
        reserve = max_reserve

    distributable = max(0.0, adjusted - reserve)
    carry_out = max(0.0, -adjusted)

    dist_q = _money_decimal(distributable)
    lines = []
    allocated = Decimal("0.00")
    active_partners = [x for x in partners if float(x["share_percent"] or 0) > 0]
    for i, ptn in enumerate(active_partners):
        if i == len(active_partners) - 1:
            amount = dist_q - allocated
        else:
            amount = (dist_q * Decimal(str(ptn["share_percent"])) / Decimal("100")).quantize(
                Decimal("0.01"), rounding=ROUND_HALF_UP
            )
            allocated += amount
        lines.append({
            "partner_id": ptn["id"],
            "name": ptn["name"],
            "share_percent": float(ptn["share_percent"]),
            "amount": float(amount),
        })

    con = db_conn()
    existing = rowdict(con.execute(
        "SELECT * FROM profit_distributions WHERE period_month=? AND active=1",
        (period_month,)
    ).fetchone())
    con.close()

    open_days = period_open_days(start, end)
    month_finished = end < accounting_day_iso()
    ready = (
        abs(total_share - 100.0) < 0.01
        and int(p["waiting"]) == 0
        and not open_days
        and month_finished
        and reserve_valid
        and not existing
        and len(active_partners) > 0
    )
    return {
        "month": period_month,
        "start": start,
        "end": end,
        "operating_profit": round(float(p["operating_profit"]), 2),
        "carry_loss_in": round(carry_in, 2),
        "card_protection_in": round(card_protection_in, 2),
        "card_protection_out": round(card_protection_out, 2),
        "card_protection_change": round(card_protection_change, 2),
        "adjusted_profit": round(adjusted, 2),
        "reserve_amount": round(reserve, 2),
        "max_reserve": round(max_reserve, 2),
        "distributable_profit": round(distributable, 2),
        "carry_loss_out": round(carry_out, 2),
        "total_share": total_share,
        "waiting": int(p["waiting"]),
        "open_days": open_days,
        "month_finished": month_finished,
        "reserve_valid": reserve_valid,
        "existing": existing,
        "lines": lines,
        "ready": ready,
    }


def finalize_profit_distribution(period_month, reserve_amount, note=""):
    preview = profit_distribution_preview(period_month, reserve_amount)
    if preview["existing"]:
        raise RuntimeError(f"{period_month} dönemi daha önce kesinleştirilmiş.")
    if abs(preview["total_share"] - 100.0) >= 0.01:
        raise RuntimeError(f"Ortak payları toplamı %100 olmalı. Şu an %{preview['total_share']:.2f}.")
    if preview["waiting"]:
        raise RuntimeError(f"{preview['waiting']} PanCafe bekleyen kasa kaydı var.")
    if preview["open_days"]:
        sample = ", ".join(preview["open_days"][:5])
        more = "..." if len(preview["open_days"]) > 5 else ""
        raise RuntimeError(f"Dönemde açık gün sonları var: {sample}{more}")
    if not preview["month_finished"]:
        raise RuntimeError("Ay bitmeden kâr dağıtımı kesinleştirilemez.")
    if not preview["reserve_valid"]:
        raise RuntimeError("İşletmede bırakılacak tutar düzeltilmiş kârdan fazla olamaz.")
    if not preview["lines"]:
        raise RuntimeError("Kâr dağıtımı için aktif ortak bulunamadı.")

    con = db_conn()
    try:
        con.execute("BEGIN IMMEDIATE")
        cur = con.execute(
            """INSERT INTO profit_distributions(
                period_month,period_start,period_end,operating_profit,carry_loss_in,
                card_protection_in,card_protection_out,reserve_amount,distributable_profit,carry_loss_out,total_share,note,
                active,created_at,created_by
            ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,1,?,?)""",
            (
                preview["month"], preview["start"], preview["end"],
                preview["operating_profit"], preview["carry_loss_in"],
                preview["card_protection_in"], preview["card_protection_out"],
                preview["reserve_amount"], preview["distributable_profit"],
                preview["carry_loss_out"], preview["total_share"],
                (note or "").strip(), now_iso(), current_actor_name()
            )
        )
        did = cur.lastrowid

        created_movements = []
        for line in preview["lines"]:
            mid = None
            if line["amount"] > 0.0001:
                movement_note = (
                    f"{preview['month']} kâr payı · %{line['share_percent']:.2f}"
                )
                if note:
                    movement_note += f" — {note.strip()}"
                mcur = con.execute(
                    """INSERT INTO partner_movements(
                        tx_date,created_at,partner_id,movement_type,amount,payment_method,
                        payment_target,note,source_type,source_id,actor_name,actor_type,active
                    ) VALUES(?,?,?,?,?,'-','AUTO',?,'PROFIT_DISTRIBUTION',?,?, 'ADMIN',1)""",
                    (
                        preview["end"], now_iso(), line["partner_id"], "ALACAK EKLE",
                        line["amount"], movement_note, did, current_actor_name()
                    )
                )
                mid = mcur.lastrowid
                created_movements.append(mid)

            con.execute(
                """INSERT INTO profit_distribution_lines(
                    distribution_id,partner_id,share_percent,amount,partner_movement_id,created_at
                ) VALUES(?,?,?,?,?,?)""",
                (
                    did, line["partner_id"], line["share_percent"], line["amount"],
                    mid, now_iso()
                )
            )
        con.commit()
    except Exception:
        con.rollback()
        con.close()
        raise
    con.close()

    after = {
        "distribution_id": did,
        **preview,
        "note": (note or "").strip(),
    }
    add_audit(
        "PROFIT_DISTRIBUTION", did, "FINALIZE",
        f"{preview['month']} kâr dağıtımı kesinleştirildi: {preview['distributable_profit']:,.2f} TL",
        {}, after
    )
    for mid in created_movements:
        add_audit(
            "PARTNER_MOVEMENT", mid, "AUTO_CREATE",
            f"{preview['month']} kâr payı alacağı otomatik oluşturuldu."
        )
    schedule_excel_sync()
    return preview


def create_personal_card_claim(tx_date, category, amount, owner, note, actor_name, actor_type):
    """Create a partner receivable without creating an expense yet."""
    tx_date = parse_day(tx_date)
    category = (category or "").strip().upper()
    owner = (owner or "").strip()
    if category not in CATEGORIES:
        raise RuntimeError("Gider kategorisi geçersiz.")
    if owner not in CARD_OWNERS:
        raise RuntimeError("Şahsi kart sahibini seçin.")
    if amount <= 0:
        raise RuntimeError("Tutar 0'dan büyük olmalı.")
    pid = ensure_partner(owner)
    claim_note = (note or "").strip()
    con = db_conn()
    cur = con.execute(
        """INSERT INTO partner_movements(
            tx_date,created_at,partner_id,movement_type,amount,payment_method,payment_target,
            note,source_type,source_id,claim_category,claim_origin,actor_name,actor_type,active
        ) VALUES(?,?,?,?,?,'-','AUTO',?,?,NULL,?,'ŞAHSİ KART',?,?,1)""",
        (tx_date, now_iso(), pid, "ALACAK EKLE", amount, claim_note,
         PERSONAL_CARD_CLAIM_SOURCE, category, actor_name, actor_type)
    )
    mid = cur.lastrowid
    con.commit()
    after = rowdict(con.execute("SELECT * FROM partner_movements WHERE id=?", (mid,)).fetchone())
    con.close()
    add_audit(
        "PARTNER_MOVEMENT", mid, "CARD_CLAIM_CREATE",
        f"Şahsi kart alacağı oluşturuldu: {category} {amount:,.2f} TL. Gider henüz oluşmadı.",
        {}, after, actor_name, actor_type
    )
    rebuild_partner_payment_allocations(pid, actor_name or "SYSTEM")
    return mid


def sync_expense_claim(expense_id):
    """Legacy compatibility only. V10.5 creates personal-card claims directly."""
    return None


def _allocation_plan(partner_id):
    con = db_conn()
    rows = [dict(r) for r in con.execute(
        """SELECT pm.*,p.name partner_name
             FROM partner_movements pm JOIN partners p ON p.id=pm.partner_id
            WHERE pm.active=1 AND pm.partner_id=?
              AND pm.movement_type IN ('ALACAK EKLE','ALACAK ÖDEME')
            ORDER BY pm.tx_date,pm.id""", (partner_id,)
    ).fetchall()]
    con.close()

    claims = []
    desired = {}
    by_id = {r["id"]: r for r in rows}
    for r in rows:
        if r["movement_type"] == "ALACAK EKLE":
            claims.append({"row": r, "remaining": float(r["amount"] or 0), "bucket": _claim_bucket(r.get("source_type"))})
            continue

        left = float(r["amount"] or 0)
        target = (r.get("payment_target") or "AUTO").upper()
        if target == "MASRAF":
            eligible = [c for c in claims if c["bucket"] == "MASRAF" and c["remaining"] > 0.0001]
        elif target == "KAR_PAYI":
            eligible = [c for c in claims if c["bucket"] == "KAR_PAYI" and c["remaining"] > 0.0001]
        else:
            eligible = [c for c in claims if c["remaining"] > 0.0001]

        for c in eligible:
            if left <= 0.0001:
                break
            used = min(c["remaining"], left)
            if used > 0.0001:
                desired[(r["id"], c["row"]["id"])] = round(used, 2)
                c["remaining"] -= used
                left -= used
    return desired, by_id


def personal_card_claim_paid_amount(claim_id):
    con = db_conn()
    row = con.execute(
        "SELECT COALESCE(SUM(amount),0) FROM partner_payment_allocations WHERE active=1 AND claim_movement_id=?",
        (claim_id,)
    ).fetchone()
    con.close()
    return float(row[0] or 0)


def rebuild_partner_payment_allocations(partner_id, actor_name="SYSTEM"):
    """Rebuild payment-to-claim allocation and actual expense rows for one partner.

    Only allocations against PERSONAL_CARD_CLAIM become expenses. Profit-share and
    other receivable payments remain cash outflows but are not economic expenses.
    """
    desired, by_id = _allocation_plan(partner_id)
    con = db_conn()
    partner = con.execute("SELECT * FROM partners WHERE id=?", (partner_id,)).fetchone()
    if not partner:
        con.close(); return
    partner_name = partner["name"]
    existing = {
        (r["payment_movement_id"], r["claim_movement_id"]): dict(r)
        for r in con.execute(
            """SELECT a.* FROM partner_payment_allocations a
                JOIN partner_movements pay ON pay.id=a.payment_movement_id
               WHERE pay.partner_id=?""", (partner_id,)
        ).fetchall()
    }

    try:
        con.execute("BEGIN IMMEDIATE")
        # Disable allocations no longer in the chronological plan.
        for key, old in existing.items():
            if key in desired:
                continue
            con.execute("UPDATE partner_payment_allocations SET active=0,updated_at=? WHERE id=?", (now_iso(), old["id"]))
            if old.get("expense_id"):
                con.execute(
                    """UPDATE expenses SET active=0,canceled_at=?,canceled_by=?,
                       cancel_reason='Ortak ödeme mahsubu yeniden hesaplandı',version=version+1
                       WHERE id=? AND active=1""",
                    (now_iso(), actor_name, old["expense_id"])
                )

        for key, amount in desired.items():
            pay_id, claim_id = key
            pay = by_id.get(pay_id)
            claim = by_id.get(claim_id)
            if not pay or not claim:
                continue
            old = existing.get(key)
            if old:
                aid = old["id"]
                con.execute(
                    "UPDATE partner_payment_allocations SET amount=?,active=1,updated_at=? WHERE id=?",
                    (amount, now_iso(), aid)
                )
            else:
                cur = con.execute(
                    """INSERT INTO partner_payment_allocations(
                        payment_movement_id,claim_movement_id,amount,active,created_at
                    ) VALUES(?,?,?,1,?)""",
                    (pay_id, claim_id, amount, now_iso())
                )
                aid = cur.lastrowid
                old = {"expense_id": None}

            is_card_claim = (claim.get("source_type") or "").upper() in (PERSONAL_CARD_CLAIM_SOURCE, "EXPENSE")
            if is_card_claim:
                category = (claim.get("claim_category") or "").strip().upper()
                if category not in CATEGORIES:
                    category = "EXTRA"
                desc = (claim.get("note") or "").strip()
                note = f"Şahsi kart tahsilatı — {partner_name} — Alacak #{claim_id}"
                if desc:
                    note += f" — {desc}"
                eid = old.get("expense_id") if old else None
                erow = con.execute("SELECT * FROM expenses WHERE id=?", (eid,)).fetchone() if eid else None
                if erow:
                    con.execute(
                        """UPDATE expenses SET tx_date=?,category=?,amount=?,note=?,payment_source='İŞLETME',
                           payment_method='NAKİT',card_owner=?,system_type=?,source_type='PARTNER_PAYMENT_ALLOCATION',
                           source_id=?,active=1,updated_at=?,updated_by=?,canceled_at=NULL,canceled_by=NULL,
                           cancel_reason='',version=version+1 WHERE id=?""",
                        (pay["tx_date"], category, amount, note, partner_name,
                         PERSONAL_CARD_REIMBURSEMENT_TYPE, aid, now_iso(), actor_name, eid)
                    )
                else:
                    cur = con.execute(
                        """INSERT INTO expenses(
                            tx_date,created_at,category,amount,note,payment_source,payment_method,card_owner,
                            system_type,source_type,source_id,actor_name,actor_type,active,created_by
                        ) VALUES(?,?,?,?,?,'İŞLETME','NAKİT',?,?, 'PARTNER_PAYMENT_ALLOCATION',?,?, 'ADMIN',1,?)""",
                        (pay["tx_date"], now_iso(), category, amount, note, partner_name,
                         PERSONAL_CARD_REIMBURSEMENT_TYPE, aid, actor_name, actor_name)
                    )
                    eid = cur.lastrowid
                con.execute("UPDATE partner_payment_allocations SET expense_id=? WHERE id=?", (eid, aid))
            elif old and old.get("expense_id"):
                con.execute(
                    """UPDATE expenses SET active=0,canceled_at=?,canceled_by=?,
                       cancel_reason='Bu ödeme ekonomik gider oluşturmayan alacağa mahsup edildi',version=version+1
                       WHERE id=? AND active=1""",
                    (now_iso(), actor_name, old["expense_id"])
                )
                con.execute("UPDATE partner_payment_allocations SET expense_id=NULL WHERE id=?", (aid,))
        con.commit()
    except Exception:
        con.rollback(); con.close(); raise
    con.close()


def sync_partner_payment_expense(movement_id):
    """Compatibility wrapper; V10.5 rebuilds exact partial-payment allocations."""
    con = db_conn()
    mov = con.execute("SELECT partner_id FROM partner_movements WHERE id=?", (movement_id,)).fetchone()
    con.close()
    if mov:
        rebuild_partner_payment_allocations(int(mov["partner_id"]), current_actor_name() or "SYSTEM")


def migrate_v105_personal_card_logic():
    """One-time safe conversion of V10.4 personal-card expenses to V10.5 claims."""
    if meta_get("v105_personal_card_migrated", "0") == "1":
        return
    con = db_conn()
    try:
        con.execute("BEGIN IMMEDIATE")
        legacy = [dict(r) for r in con.execute(
            """SELECT * FROM expenses
                WHERE active=1 AND system_type='' AND payment_source='ŞAHSİ KART'"""
        ).fetchall()]
        affected = set()
        for exp in legacy:
            linked = con.execute(
                """SELECT * FROM partner_movements
                    WHERE active=1 AND source_type IN ('EXPENSE','PERSONAL_CARD_CLAIM') AND source_id=?
                    ORDER BY id LIMIT 1""", (exp["id"],)
            ).fetchone()
            if linked:
                con.execute(
                    """UPDATE partner_movements SET source_type=?,claim_category=?,claim_origin='ŞAHSİ KART',
                       note=CASE WHEN note='' THEN ? ELSE note END,updated_at=?,updated_by=?,version=version+1
                       WHERE id=?""",
                    (PERSONAL_CARD_CLAIM_SOURCE, exp["category"], exp["note"], now_iso(), "V10.5 MIGRASYON", linked["id"])
                )
                affected.add(int(linked["partner_id"]))
            else:
                prow = con.execute("SELECT id FROM partners WHERE name=? COLLATE NOCASE", (exp["card_owner"],)).fetchone()
                if prow:
                    pid = int(prow["id"])
                else:
                    cur = con.execute(
                        "INSERT INTO partners(name,share_percent,note,active,created_at) VALUES(?,0,?,1,?)",
                        (exp["card_owner"] or "Bilinmeyen Ortak", "V10.5 migrasyon ile oluşturuldu.", now_iso())
                    )
                    pid = cur.lastrowid
                con.execute(
                    """INSERT INTO partner_movements(
                        tx_date,created_at,partner_id,movement_type,amount,payment_method,payment_target,note,
                        source_type,source_id,claim_category,claim_origin,actor_name,actor_type,active
                    ) VALUES(?,?,?,?,?,'-','AUTO',?,?,?,?, 'ŞAHSİ KART','V10.5 MIGRASYON','ADMIN',1)""",
                    (exp["tx_date"], exp["created_at"], pid, "ALACAK EKLE", exp["amount"], exp["note"],
                     PERSONAL_CARD_CLAIM_SOURCE, exp["id"], exp["category"])
                )
                affected.add(pid)
            con.execute(
                """UPDATE expenses SET active=0,canceled_at=?,canceled_by='V10.5 MIGRASYON',
                   cancel_reason='Şahsi kart kaydı bekleyen ortak alacağına taşındı',version=version+1 WHERE id=?""",
                (now_iso(), exp["id"])
            )

        # Old generic partner-payment expenses are replaced by exact claim/category allocations.
        con.execute(
            """UPDATE expenses SET active=0,canceled_at=?,canceled_by='V10.5 MIGRASYON',
               cancel_reason='V10.5 ödeme mahsubu ile yeniden oluşturulacak',version=version+1
               WHERE active=1 AND system_type='PARTNER_PAYMENT'""",
            (now_iso(),)
        )
        for r in con.execute("SELECT DISTINCT partner_id FROM partner_movements WHERE active=1").fetchall():
            affected.add(int(r[0]))
        con.execute(
            "INSERT INTO system_meta(meta_key,meta_value) VALUES('v105_personal_card_migrated','1') "
            "ON CONFLICT(meta_key) DO UPDATE SET meta_value='1'"
        )
        con.commit()
    except Exception:
        con.rollback(); con.close(); raise
    con.close()
    for pid in sorted(affected):
        rebuild_partner_payment_allocations(pid, "V10.5 MIGRASYON")


# ----------------------- Reports -----------------------

def profit_loss(start, end):
    start, end = parse_day(start), parse_day(end)
    con = db_conn()
    c = con.execute("""SELECT
        COALESCE(SUM(CASE WHEN pancafe IS NOT NULL THEN pancafe ELSE 0 END),0) revenue,
        COALESCE(SUM(nakit+pos+iban),0) collections,
        COALESCE(SUM(CASE WHEN pancafe IS NULL THEN 1 ELSE 0 END),0) waiting
        FROM cash_entries WHERE active=1 AND tx_date BETWEEN ? AND ?""", (start,end)).fetchone()
    e = con.execute("""SELECT
        COALESCE(SUM(CASE WHEN system_type IN ('',?,?,?,?,?) THEN amount ELSE 0 END),0) economic_expense,
        COALESCE(SUM(CASE WHEN payment_source='İŞLETME' AND system_type IN ('',?,?,?) THEN amount ELSE 0 END),0) direct_business_paid,
        COALESCE(SUM(CASE WHEN system_type=? THEN amount ELSE 0 END),0) card_reimbursed
        FROM expenses WHERE active=1 AND tx_date BETWEEN ? AND ?""",
        (
            PERSONAL_CARD_REIMBURSEMENT_TYPE, STAFF_ADVANCE_TYPE, STAFF_SALARY_TYPE, STAFF_MEAL_TYPE, PERSONAL_CARD_REIMBURSEMENT_TYPE,
            STAFF_ADVANCE_TYPE, STAFF_SALARY_TYPE, STAFF_MEAL_TYPE,
            PERSONAL_CARD_REIMBURSEMENT_TYPE, start,end
        )).fetchone()
    partner_paid = con.execute(
        """SELECT COALESCE(SUM(amount),0) FROM partner_movements
            WHERE active=1 AND movement_type='ALACAK ÖDEME' AND tx_date BETWEEN ? AND ?""", (start,end)
    ).fetchone()[0]
    pw = con.execute("SELECT COALESCE(SUM(amount),0) FROM partner_movements WHERE active=1 AND movement_type='PARA ÇIKIŞI' AND tx_date BETWEEN ? AND ?", (start,end)).fetchone()[0]
    credit = con.execute("""SELECT
        COALESCE(SUM(CASE WHEN movement_type='BORÇ' THEN amount ELSE 0 END),0),
        COALESCE(SUM(CASE WHEN movement_type='ÖDEME' THEN amount ELSE 0 END),0)
        FROM credit_entries WHERE active=1 AND tx_date BETWEEN ? AND ?""", (start,end)).fetchone()
    partner_debt = con.execute("SELECT COALESCE(SUM(CASE WHEN movement_type='ALACAK EKLE' THEN amount WHEN movement_type='ALACAK ÖDEME' THEN -amount ELSE 0 END),0) FROM partner_movements WHERE active=1 AND tx_date<=?", (end,)).fetchone()[0]
    by_cat = [dict(r) for r in con.execute(
        """SELECT category, SUM(amount) total FROM expenses
            WHERE active=1 AND system_type IN ('',?,?,?,?) AND tx_date BETWEEN ? AND ?
            GROUP BY category ORDER BY total DESC""",
        (PERSONAL_CARD_REIMBURSEMENT_TYPE,STAFF_ADVANCE_TYPE,STAFF_SALARY_TYPE,STAFF_MEAL_TYPE,start,end)
    ).fetchall()]
    pd = con.execute("""SELECT
        COALESCE(SUM(distributable_profit),0),
        COALESCE(SUM(reserve_amount),0),
        COALESCE(MAX(carry_loss_out),0)
        FROM profit_distributions
        WHERE active=1 AND period_end BETWEEN ? AND ?""", (start,end)).fetchone()
    con.close()
    revenue=float(c["revenue"]); economic=float(e["economic_expense"]); direct=float(e["direct_business_paid"]); partner_paid=float(partner_paid or 0)
    pending_card = pending_personal_card_total(end)
    return {
        "start":start,"end":end,"revenue":revenue,"collections":float(c["collections"]),"waiting":int(c["waiting"]),
        "economic_expense":economic,"business_paid":direct,"personal_card":pending_card,"pending_card_debt":pending_card,
        "card_reimbursed":float(e["card_reimbursed"]),"partner_paid":partner_paid,
        "general_cash_out":direct+partner_paid,"partner_withdrawals":float(pw or 0),"partner_debt":float(partner_debt or 0),
        "operating_profit":revenue-economic,"net_cash":float(c["collections"])-direct-partner_paid-float(pw or 0),
        "credit_new":float(credit[0]),"credit_paid":float(credit[1]),"by_category":by_cat,
        "distributed_profit":float(pd[0] or 0),"profit_reserve":float(pd[1] or 0),"profit_carry_out":float(pd[2] or 0)
    }


def monthly_profit_rows(start, end):
    con = db_conn()
    months = []
    cur = date.fromisoformat(start).replace(day=1); last = date.fromisoformat(end).replace(day=1)
    while cur <= last:
        m_start = max(cur, date.fromisoformat(start)); next_m = date(cur.year + (cur.month==12), 1 if cur.month==12 else cur.month+1, 1); m_end=min(next_m-timedelta(days=1), date.fromisoformat(end))
        p=profit_loss(m_start.isoformat(),m_end.isoformat()); p["month"]=cur.strftime("%Y-%m"); months.append(p); cur=next_m
    con.close(); return months


# ----------------------- Excel mirror / backup -----------------------

def _run_admin_secret(mode, password=None):
    if not ADMIN_SECRET_SCRIPT.exists(): raise RuntimeError("admin_secret.ps1 bulunamadı.")
    env=os.environ.copy(); env["KGA_EXCEL_PASSWORD"]=password or ""
    cmd=["powershell.exe","-NoProfile","-ExecutionPolicy","Bypass","-File",str(ADMIN_SECRET_SCRIPT),"-Mode",mode,"-Path",str(ADMIN_SECRET_FILE)]
    r=subprocess.run(cmd,capture_output=True,text=True,env=env,timeout=30,creationflags=WINDOWS_NO_WINDOW)
    if r.returncode != 0: raise RuntimeError((r.stderr or r.stdout or "Excel şifre anahtarı hatası").strip())
    return (r.stdout or "").strip()


def store_excel_password(password):
    global EXCEL_PASSWORD_CACHE
    _run_admin_secret("Protect", password); EXCEL_PASSWORD_CACHE=password


def load_excel_password():
    global EXCEL_PASSWORD_CACHE
    if EXCEL_PASSWORD_CACHE: return EXCEL_PASSWORD_CACHE
    if not ADMIN_SECRET_FILE.exists(): raise RuntimeError("Excel ana şifre anahtarı bulunamadı. KURULUM.bat çalıştırın.")
    EXCEL_PASSWORD_CACHE=_run_admin_secret("Unprotect")
    return EXCEL_PASSWORD_CACHE


def encrypt_plain_excel(plain_path, output_path, password):
    env=os.environ.copy(); env["KGA_EXCEL_PASSWORD"]=password
    cmd=["powershell.exe","-NoProfile","-ExecutionPolicy","Bypass","-File",str(CRYPTO_SCRIPT),"-Mode","Encrypt","-InputPath",str(plain_path),"-OutputPath",str(output_path)]
    r=subprocess.run(cmd,capture_output=True,text=True,env=env,timeout=120,creationflags=WINDOWS_NO_WINDOW)
    if r.returncode != 0: raise RuntimeError((r.stderr or r.stdout or "Excel şifreleme başarısız").strip())


def literal_workbook_text(wb):
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type == "f":
                    cell.data_type = "s"


def style_ws(ws, widths=None):
    fill=PatternFill("solid", fgColor="172033"); font=Font(color="FFFFFF", bold=True)
    for cell in ws[1]: cell.fill=fill; cell.font=font; cell.alignment=Alignment(horizontal="center")
    ws.freeze_panes="A2"
    if widths:
        for col,w in widths.items(): ws.column_dimensions[col].width=w


def build_excel_mirror(path):
    wb=Workbook(); wb.remove(wb.active)
    con=db_conn()

    ws=wb.create_sheet("KONTROL")
    ws["A1"]="KAMPÜS GAME ARENA — MUHASEBE V10.6.3"; ws["A1"].font=Font(size=18,bold=True,color="FFFFFF"); ws["A1"].fill=PatternFill("solid",fgColor="0F172A")
    ws.merge_cells("A1:F1"); ws["A2"]="SQLite ana kayıt · Excel otomatik senkron/yedek"; ws.merge_cells("A2:F2")
    today=accounting_day_iso(); p=profit_loss(*default_range())
    control=[("Gösterge","Değer"),("Son Excel Senkronu",now_iso()),("PanCafe Bekleyen",p["waiting"]),("Ortaklara Borç",p["partner_debt"]),("Otomatik Gün Kilidi",f"Bugün AÇIK · D-1 ve öncesi otomatik kilit"),("Ana Veri Kaynağı","KampusArena_Muhasebe.db")]
    for r,row in enumerate(control,4): ws.cell(r,1,row[0]); ws.cell(r,2,row[1])
    ws["A4"].fill=PatternFill("solid",fgColor="172033"); ws["A4"].font=Font(color="FFFFFF",bold=True); ws["B4"].fill=PatternFill("solid",fgColor="172033"); ws["B4"].font=Font(color="FFFFFF",bold=True)
    ws.column_dimensions["A"].width=28; ws.column_dimensions["B"].width=34

    # Cash
    ws=wb.create_sheet("GUNLUK_KASA"); headers=["ID","Tarih","Saat","Personel/Yönetici","PanCafe Kasa","Nakit","POS","IBAN","Hatalı Masa","İndirim","Yeni Borç","Ödenen Eski Borç","Hesaplanan","Fark","Durum","Açıklama","Kayıt Durumu"]
    ws.append(headers)
    for r in con.execute("SELECT * FROM cash_entries ORDER BY tx_date,id"):
        calc=cash_calc(r["nakit"],r["pos"],r["iban"],r["hatali"],r["indirim"],r["new_debt"],r["old_debt_paid"]); diff,status=cash_status(r["pancafe"],calc)
        ws.append([r["id"],r["tx_date"],r["created_at"][11:19],r["actor_name"],r["pancafe"],r["nakit"],r["pos"],r["iban"],r["hatali"],r["indirim"],r["new_debt"],r["old_debt_paid"],calc,diff,status,r["note"],"AKTİF" if r["active"] else "İPTAL"])
    style_ws(ws,{"B":12,"C":10,"D":20,"P":36})

    ws=wb.create_sheet("GIDERLER"); headers=["ID","Tarih","Saat","Personel/Yönetici","Kategori","Tutar","Ödeme Kaynağı","Ödeme Yöntemi","Şahsi Kart Sahibi","Tür","Açıklama","Kayıt Durumu"]
    ws.append(headers)
    for r in con.execute("SELECT * FROM expenses ORDER BY tx_date,id"):
        typ="ŞAHSİ KART TAHSİLATI" if r["system_type"]==PERSONAL_CARD_REIMBURSEMENT_TYPE else "GERÇEKLEŞEN GİDER"
        ws.append([r["id"],r["tx_date"],r["created_at"][11:19],r["actor_name"],r["category"],r["amount"],r["payment_source"],r["payment_method"],r["card_owner"],typ,r["note"],"AKTİF" if r["active"] else "İPTAL"])
    style_ws(ws,{"D":20,"E":20,"K":44})

    # V10.1 compatibility: some older paths/packages used HARCAMA_HAREKETLERI.
    # GIDERLER remains the canonical sheet; this mirror prevents a legacy worksheet lookup
    # from crashing the application during/after a clean install.
    legacy=wb.create_sheet("HARCAMA_HAREKETLERI")
    legacy.append(headers)
    for r in con.execute("SELECT * FROM expenses ORDER BY tx_date,id"):
        typ="ŞAHSİ KART TAHSİLATI" if r["system_type"]==PERSONAL_CARD_REIMBURSEMENT_TYPE else "GERÇEKLEŞEN GİDER"
        legacy.append([r["id"],r["tx_date"],r["created_at"][11:19],r["actor_name"],r["category"],r["amount"],r["payment_source"],r["payment_method"],r["card_owner"],typ,r["note"],"AKTİF" if r["active"] else "İPTAL"])
    style_ws(legacy,{"D":20,"E":20,"K":44})

    ws=wb.create_sheet("VERESIYE"); ws.append(["ID","Tarih","Saat","Personel/Yönetici","Kişi","Tür","Tutar","Yöntem","Açıklama","Kayıt Durumu"])
    for r in con.execute("SELECT * FROM credit_entries ORDER BY tx_date,id"):
        ws.append([r["id"],r["tx_date"],r["created_at"][11:19],r["actor_name"],r["person"],r["movement_type"],r["amount"],r["payment_method"],r["note"],"AKTİF" if r["active"] else "İPTAL"])
    style_ws(ws,{"D":20,"E":22,"I":42})

    ws=wb.create_sheet("ORTAKLAR"); ws.append(["ID","Ortak","Pay %","Bekleyen Şahsi Kart","Kâr Payı Alacağı","Diğer Alacak","Toplam Ortağa Borç","Alacak Yazılan","Alacak Ödenen","Para Çıkışı","Not","Durum"])
    for ptn in con.execute("SELECT * FROM partners ORDER BY name COLLATE NOCASE"):
        b=partner_claim_breakdown(ptn["id"])
        sums=con.execute("SELECT COALESCE(SUM(CASE WHEN movement_type='PARA ÇIKIŞI' AND active=1 THEN amount ELSE 0 END),0) FROM partner_movements WHERE partner_id=?",(ptn["id"],)).fetchone()
        ws.append([ptn["id"],ptn["name"],ptn["share_percent"],b["expense_left"],b["profit_left"],b["other_left"],b["total_left"],b["claim_add"],b["claim_paid"],sums[0],ptn["note"],"AKTİF" if ptn["active"] else "İPTAL"])
    style_ws(ws,{"B":20,"K":44})

    ws=wb.create_sheet("ORTAK_HAREKETLERI"); ws.append(["ID","Tarih","Saat","Ortak","Tür","Tutar","Ödeme Yöntemi","Mahsup","Kaynak","Kategori","Alım Kaynağı","Kaynak ID","Açıklama","Durum"])
    for r in con.execute("SELECT pm.*,p.name partner_name FROM partner_movements pm JOIN partners p ON p.id=pm.partner_id ORDER BY pm.tx_date,pm.id"):
        ws.append([r["id"],r["tx_date"],r["created_at"][11:19],r["partner_name"],r["movement_type"],r["amount"],r["payment_method"],r["payment_target"],r["source_type"],r["claim_category"],r["claim_origin"],r["source_id"],r["note"],"AKTİF" if r["active"] else "İPTAL"])
    style_ws(ws,{"D":20,"K":44})

    ws=wb.create_sheet("KAR_DAGITIMI"); ws.append(["ID","Dönem","Faaliyet Kârı","Devreden Zarar","Kart Koruma Giriş","Kart Koruma Çıkış","İşletmede Bırakılan","Dağıtılabilir Kâr","Sonraki Aya Zarar","Toplam Pay %","Kesinleştirme","Kesinleştiren","Not"])
    for r in con.execute("SELECT * FROM profit_distributions ORDER BY period_month"):
        ws.append([r["id"],r["period_month"],r["operating_profit"],r["carry_loss_in"],r["card_protection_in"],r["card_protection_out"],r["reserve_amount"],r["distributable_profit"],r["carry_loss_out"],r["total_share"],r["created_at"],r["created_by"],r["note"]])
    style_ws(ws,{"B":14,"I":22,"J":20,"K":44})

    ws=wb.create_sheet("KAR_DAGITIM_DETAY"); ws.append(["Dağıtım ID","Dönem","Ortak","Pay %","Kâr Payı","Ortak Hareket ID"])
    for r in con.execute("""SELECT l.*,d.period_month,p.name partner_name
                            FROM profit_distribution_lines l
                            JOIN profit_distributions d ON d.id=l.distribution_id
                            JOIN partners p ON p.id=l.partner_id
                            ORDER BY d.period_month,l.id"""):
        ws.append([r["distribution_id"],r["period_month"],r["partner_name"],r["share_percent"],r["amount"],r["partner_movement_id"]])
    style_ws(ws,{"B":14,"C":20})

    ws=wb.create_sheet("PERSONEL"); ws.append(["ID","Personel","Aylık Maaş","Varsayılan Vardiya","Durum","Oluşturma"])
    for r in con.execute("SELECT * FROM personnel ORDER BY name COLLATE NOCASE"):
        ws.append([r["id"],r["name"],r["monthly_salary"],SHIFT_DEFINITIONS.get(r["default_shift"],{}).get("label",r["default_shift"] or "HER GÜN SEÇER"),"AKTİF" if r["active"] else "PASİF",r["created_at"]])
    style_ws(ws,{"B":24,"C":16,"D":22,"F":22})

    ws=wb.create_sheet("MAAS_GECMISI"); ws.append(["Personel ID","Personel","Geçerlilik Ayı","Aylık Maaş"])
    for r in con.execute("SELECT h.*,p.name FROM personnel_salary_history h JOIN personnel p ON p.id=h.personnel_id ORDER BY p.name,h.effective_month"):
        ws.append([r["personnel_id"],r["name"],r["effective_month"],r["monthly_salary"]])
    style_ws(ws,{"B":24,"C":18,"D":18})

    ws=wb.create_sheet("PUANTAJ"); ws.append(["ID","Tarih","Personel","Vardiya","Plan Giriş","Plan Çıkış","Gerçek Giriş","Gerçek Çıkış","Çalışılan Saat","Not","Durum"])
    for r in con.execute("SELECT a.*,p.name personnel_name FROM personnel_attendance a JOIN personnel p ON p.id=a.personnel_id ORDER BY a.work_date,a.id"):
        ws.append([r["id"],r["work_date"],r["personnel_name"],SHIFT_DEFINITIONS.get(r["shift_code"],{}).get("label",r["shift_code"]),r["scheduled_start"],r["scheduled_end"],r["actual_start"],r["actual_end"],round((r["worked_minutes"] or 0)/60.0,2),r["note"],"AKTİF" if r["active"] else "İPTAL"])
    style_ws(ws,{"B":12,"C":22,"D":20,"J":40})

    ws=wb.create_sheet("PERSONEL_AVANS"); ws.append(["ID","Tarih","Personel","Tutar","Ödeme","Gider ID","Açıklama","Durum"])
    for r in con.execute("SELECT a.*,p.name personnel_name FROM personnel_advances a JOIN personnel p ON p.id=a.personnel_id ORDER BY a.tx_date,a.id"):
        ws.append([r["id"],r["tx_date"],r["personnel_name"],r["amount"],r["payment_method"],r["expense_id"],r["note"],"AKTİF" if r["active"] else "İPTAL"])
    style_ws(ws,{"B":12,"C":22,"G":40})

    ws=wb.create_sheet("PERSONEL_ODEME"); ws.append(["ID","Dönem","Ödeme Tarihi","Personel","Tür","Tutar","Yöntem","Gider ID","Açıklama","Durum"])
    for r in con.execute("SELECT x.*,p.name personnel_name FROM personnel_payments x JOIN personnel p ON p.id=x.personnel_id ORDER BY x.period_month,x.id"):
        ws.append([r["id"],r["period_month"],r["tx_date"],r["personnel_name"],r["payment_kind"],r["amount"],r["payment_method"],r["expense_id"],r["note"],"AKTİF" if r["active"] else "İPTAL"])
    style_ws(ws,{"B":12,"C":12,"D":22,"I":40})

    ws=wb.create_sheet("GUN_SONU"); ws.append(["Tarih","Durum","Kapatma Saati","Kapatan","Açma Saati","Açan","Not","Kural"])
    for r in con.execute("SELECT * FROM day_closures ORDER BY tx_date DESC"):
        waiting=con.execute("SELECT COUNT(*) FROM cash_entries WHERE active=1 AND tx_date=? AND pancafe IS NULL",(r["tx_date"],)).fetchone()[0]
        if r["is_closed"]:
            durum="OTOMATİK KAPALI" if r["closed_by"]=="OTOMATIK" else "KAPALI"
        elif r["reopened_at"] and accounting_day(datetime.fromisoformat(str(r["reopened_at"]))) == accounting_day():
            durum="GEÇİCİ AÇIK"
        elif waiting:
            durum=f"PANCAFE BEKLİYOR ({waiting})"
        else:
            durum="AÇIK"
        ws.append([r["tx_date"],durum,r["closed_at"],r["closed_by"],r["reopened_at"],r["reopened_by"],r["close_note"],"D-1 OTOMATİK"])
    style_ws(ws,{"A":14,"B":24,"C":22,"E":22,"G":48,"H":18})

    ws=wb.create_sheet("KAR_ZARAR"); ws.append(["Dönem","PanCafe Ciro","Gerçekleşen Gider","Faaliyet Sonucu","Dağıtılan Kâr","İşletmede Bırakılan","Devreden Zarar","İşletmeden Direkt Ödenen","Ortağa NAKİT Alacak Ödemesi","Genel Para Çıkışı","Ortak Para Çıkışı","Tahsilat","Net Nakit"])
    first_date=con.execute("SELECT MIN(tx_date) FROM cash_entries WHERE active=1").fetchone()[0] or accounting_day().replace(day=1).isoformat(); last_date=accounting_day_iso()
    for p in monthly_profit_rows(first_date,last_date): ws.append([p["month"],p["revenue"],p["economic_expense"],p["operating_profit"],p["distributed_profit"],p["profit_reserve"],p["profit_carry_out"],p["business_paid"],p["partner_paid"],p["general_cash_out"],p["partner_withdrawals"],p["collections"],p["net_cash"]])
    style_ws(ws,{"A":14})

    ws=wb.create_sheet("AUDIT"); ws.append(["ID","Zaman","Kullanıcı","Kaynak","Varlık","Kayıt ID","İşlem","Özet"])
    for r in con.execute("SELECT * FROM audit_log ORDER BY id"):
        ws.append([r["id"],r["created_at"],r["actor_name"],r["actor_type"],r["entity_type"],r["entity_id"],r["action"],r["summary"]])
    style_ws(ws,{"B":22,"C":20,"H":54})

    ws=wb.create_sheet("AYARLAR"); ws.append(["Ayar","Değer"]); ws.append(["Ana Veri Kaynağı","SQLite / KampusArena_Muhasebe.db"]); ws.append(["Excel Rolü","Otomatik senkron + yedek + manuel inceleme"]); ws.append(["Excel Şifresi","DOSYA PAROLASI BU EXCEL İÇİNDE SAKLANMAZ"]); ws.append(["Sürüm","V10.6.3"]); style_ws(ws,{"A":26,"B":48})

    con.close(); literal_workbook_text(wb); wb.save(path)


def scheduled_backup(kind, slot, now=None):
    """Independent, atomic backup of SQLite plus available recovery assets."""
    now = now or datetime.now(timezone(timedelta(hours=3)))
    folder = BACKUP_DIR / "planli"
    folder.mkdir(parents=True, exist_ok=True)
    filename = f"{kind}_{slot}.zip"
    destination = folder / filename
    if destination.exists():
        return destination
    partial = folder / (filename + ".partial")
    try:
        with tempfile.TemporaryDirectory(prefix="snapshot_", dir=folder) as working:
            snapshot = Path(working) / "KampusArena_Muhasebe.db"
            source = db_conn()
            target = sqlite3.connect(snapshot)
            try:
                source.backup(target)
                if target.execute("PRAGMA integrity_check").fetchone()[0] != "ok":
                    raise RuntimeError("Veritabanı yedeği doğrulanamadı.")
            finally:
                target.close()
                source.close()
            with zipfile.ZipFile(partial, "w", zipfile.ZIP_DEFLATED) as archive:
                archive.write(snapshot, "KampusArena_Muhasebe.db")
                for path, name in [(EXCEL_FILE,"KampusArena_Muhasebe.xlsx"),
                                   (ADMIN_SECRET_FILE,"admin_excel_secret.dat"),
                                   (RUNTIME_DIR/"session.key","runtime/session.key")]:
                    if path.exists():
                        archive.write(path, name)
                archive.writestr("YEDEK_BILGISI.json",json.dumps({
                    "created_at":now.isoformat(),"kind":kind,"version":"10.6.3",
                    "database":"Tutarlı SQLite snapshot; ana veri kaynağı",
                    "excel_last_sync":meta_get("last_excel_sync","Henüz yok"),
                    "excel_note":"Excel kopyası son başarılı senkron anını içerir.",
                    "recovery":"Programı tamamen kapatın. Geri yükleme önce yeni klasörde doğrulanmalıdır. Excel anahtarı aynı Windows kullanıcısına bağlıdır."
                },ensure_ascii=False,indent=2))
            with zipfile.ZipFile(partial) as archive:
                if archive.testzip() is not None:
                    raise RuntimeError("Yedek ZIP doğrulaması başarısız.")
            os.replace(partial,destination)
        return destination
    finally:
        partial.unlink(missing_ok=True)


def scheduled_backup_tick(now=None):
    now = now or datetime.now(timezone(timedelta(hours=3)))
    with SCHEDULED_BACKUP_LOCK:
        hourly = now.strftime("%Y%m%d_%H")
        due_day = now.date() if (now.hour,now.minute)>=(9,5) else now.date()-timedelta(days=1)
        paths = [scheduled_backup("Saatlik",hourly,now),scheduled_backup("Gunluk",due_day.strftime("%Y%m%d"),now)]
        folder = BACKUP_DIR/"planli"
        for kind, keep in [("Saatlik",24),("Gunluk",30)]:
            files = sorted((f for f in folder.glob(kind+"_*.zip") if re.fullmatch(kind+r"_[0-9]{8}"+(r"_[0-9]{2}" if kind=="Saatlik" else "")+r"\.zip",f.name)),key=lambda f:f.name,reverse=True)
            for old in files[keep:]:
                old.unlink()
        meta_set("last_scheduled_backup",now.isoformat())
        meta_set("scheduled_backup_error","")
        return paths


def scheduled_backup_worker():
    while True:
        try:
            scheduled_backup_tick()
        except Exception as e:
            try:
                meta_set("scheduled_backup_error",str(e))
                add_log("ERROR","PLANLI_YEDEK_HATA",str(e),source="SYSTEM")
            except Exception:
                pass
        time.sleep(60)


def scheduled_backup_info():
    folder = BACKUP_DIR/"planli"
    files = sorted(folder.glob("*.zip"),key=lambda f:f.stat().st_mtime,reverse=True) if folder.exists() else []
    return {"last":meta_get("last_scheduled_backup","Henüz yok"),
            "error":meta_get("scheduled_backup_error",""),
            "files":[{"name":f.name,"size_kb":round(f.stat().st_size/1024,1)} for f in files[:12]]}


@app.post("/admin/backups/create")
@owner_required
def manual_backup_create():
    try:
        now = datetime.now(timezone(timedelta(hours=3)))
        with SCHEDULED_BACKUP_LOCK:
            path = scheduled_backup("Manuel",now.strftime("%Y%m%d_%H%M%S_%f"),now)
        flash("Yedek oluşturuldu: " + path.name,"success")
    except Exception as e:
        flash("Yedek alınamadı: " + str(e),"danger")
    return redirect(url_for("settings"))


@app.get("/admin/backups/download/<filename>")
@owner_required
def backup_download(filename):
    if not re.fullmatch(r"(?:Saatlik_[0-9]{8}_[0-9]{2}|Gunluk_[0-9]{8}|Manuel_[0-9]{8}_[0-9]{6}_[0-9]{6})\.zip",filename):
        abort(404)
    path = BACKUP_DIR/"planli"/filename
    if not path.is_file():
        abort(404)
    return send_file(path,as_attachment=True,download_name=filename)


def backup_if_due(force=False):
    global LAST_BACKUP_TS
    now=time.time()
    if not force and now-LAST_BACKUP_TS < 300: return
    stamp=datetime.now().strftime("%Y%m%d_%H%M%S")
    try:
        if DB_FILE.exists():
            src=db_conn(); dest=sqlite3.connect(BACKUP_DIR / f"KampusArena_DB_{stamp}.db")
            src.backup(dest); dest.close(); src.close()
        if EXCEL_FILE.exists(): shutil.copy2(EXCEL_FILE, BACKUP_DIR / f"KampusArena_Excel_{stamp}.xlsx")
        LAST_BACKUP_TS=now; meta_set("last_backup",now_iso())
        for pattern in ("KampusArena_DB_*.db","KampusArena_Excel_*.xlsx"):
            files=sorted(BACKUP_DIR.glob(pattern),key=lambda p:p.stat().st_mtime,reverse=True)
            for old in files[100:]:
                try: old.unlink()
                except: pass
    except Exception as e:
        add_log("ERROR","BACKUP_HATA",str(e),source="SYSTEM")


def sync_excel_now(force_backup=False, password=None):
    with EXCEL_SYNC_LOCK:
        started=time.perf_counter(); plain=RUNTIME_DIR / f"mirror_{secrets.token_hex(8)}.xlsx"; encrypted=RUNTIME_DIR / f"mirror_enc_{secrets.token_hex(8)}.xlsx"
        try:
            build_excel_mirror(plain)
            pwd=password or load_excel_password()
            encrypt_plain_excel(plain,encrypted,pwd)
            os.replace(encrypted,EXCEL_FILE)
            backup_if_due(force_backup)
            meta_set("last_excel_sync",now_iso()); meta_set("excel_sync_error","")
            add_log("SYNC","EXCEL_SENKRON",f"{int((time.perf_counter()-started)*1000)} ms",source="SYSTEM")
        except Exception as e:
            meta_set("excel_sync_error",str(e)); add_log("ERROR","EXCEL_SENKRON_HATA",str(e),source="SYSTEM"); raise
        finally:
            for p in (plain,encrypted):
                try: p.unlink(missing_ok=True)
                except: pass


def _sync_worker(generation):
    global EXCEL_SYNC_GENERATION, EXCEL_SYNC_TIMER
    with EXCEL_SCHEDULE_LOCK:
        if generation != EXCEL_SYNC_GENERATION:
            return
        EXCEL_SYNC_TIMER = None
    try:
        sync_excel_now(False)
    except Exception:
        pass


def schedule_excel_sync():
    """Debounce expensive Excel mirror generation.

    SQLite is saved immediately. Excel is refreshed after a quiet period,
    so rapid admin/personnel entries collapse into a single export/encryption.
    """
    global EXCEL_SYNC_GENERATION, EXCEL_SYNC_TIMER
    with EXCEL_SCHEDULE_LOCK:
        EXCEL_SYNC_GENERATION += 1
        gen = EXCEL_SYNC_GENERATION
        if EXCEL_SYNC_TIMER is not None:
            try:
                EXCEL_SYNC_TIMER.cancel()
            except Exception:
                pass
        timer = threading.Timer(EXCEL_SYNC_DELAY, _sync_worker, args=(gen,))
        timer.daemon = True
        EXCEL_SYNC_TIMER = timer
        timer.start()


def installation_self_check():
    """V10.2 clean-install sanity checks; does not require the encrypted Excel to be open."""
    problems=[]
    if not CRYPTO_SCRIPT.exists(): problems.append("excel_crypto.ps1 eksik")
    if not ADMIN_SECRET_SCRIPT.exists(): problems.append("admin_secret.ps1 eksik")
    template=BASE_DIR / "KampusArena_Muhasebe_SABLON.kga"
    if not template.exists(): problems.append("KampusArena_Muhasebe_SABLON.kga eksik")
    if problems:
        add_log("ERROR","KURULUM_KONTROL","; ".join(problems),source="SYSTEM")
    return problems


# ----------------------- Query helpers -----------------------

def cash_rows(start,end):
    con=db_conn(); rows=[]
    for r in con.execute("SELECT * FROM cash_entries WHERE active=1 AND tx_date BETWEEN ? AND ? ORDER BY tx_date DESC,id DESC",(start,end)):
        d=dict(r); d["calc"]=cash_calc(d["nakit"],d["pos"],d["iban"],d["hatali"],d["indirim"],d["new_debt"],d["old_debt_paid"]); d["diff"],d["status"]=cash_status(d["pancafe"],d["calc"]); d["time_text"]=d["created_at"][11:19]; rows.append(d)
    con.close(); return rows


def expense_rows(start,end,category="TÜMÜ"):
    con=db_conn(); args=[start,end]; sql="SELECT * FROM expenses WHERE active=1 AND tx_date BETWEEN ? AND ?"
    if category and category!="TÜMÜ": sql+=" AND category=?"; args.append(category)
    sql+=" ORDER BY tx_date DESC,id DESC"; rows=[]
    for r in con.execute(sql,args): d=dict(r); d["time_text"]=d["created_at"][11:19]; rows.append(d)
    con.close(); return rows


def credit_rows(start,end):
    con=db_conn(); rows=[dict(r) for r in con.execute("SELECT * FROM credit_entries WHERE active=1 AND tx_date BETWEEN ? AND ? ORDER BY tx_date DESC,id DESC",(start,end)).fetchall()]; con.close(); return rows


def credit_summary():
    con=db_conn(); rows=[dict(r) for r in con.execute("""SELECT person,
        SUM(CASE WHEN movement_type='BORÇ' THEN amount ELSE 0 END) debt,
        SUM(CASE WHEN movement_type='ÖDEME' THEN amount ELSE 0 END) paid,
        SUM(CASE WHEN movement_type='BORÇ' THEN amount ELSE -amount END) balance,
        MAX(tx_date) last_date FROM credit_entries WHERE active=1 GROUP BY person ORDER BY balance DESC,person COLLATE NOCASE""").fetchall()]; con.close(); return rows


def partner_data():
    con=db_conn(); out=[]
    partners = [dict(p) for p in con.execute(
        "SELECT * FROM partners WHERE active=1 ORDER BY name COLLATE NOCASE"
    ).fetchall()]
    con.close()
    for d in partners:
        con=db_conn()
        s=con.execute("""SELECT
          COALESCE(SUM(CASE WHEN active=1 AND movement_type='PARA ÇIKIŞI' THEN amount ELSE 0 END),0) out_amt,
          COALESCE(SUM(CASE WHEN active=1 AND movement_type='ALACAK EKLE' THEN amount ELSE 0 END),0) claim_add,
          COALESCE(SUM(CASE WHEN active=1 AND movement_type='ALACAK ÖDEME' THEN amount ELSE 0 END),0) claim_paid,
          COALESCE(SUM(CASE WHEN active=1 AND movement_type='ALACAK EKLE' AND source_type IN ('EXPENSE','PERSONAL_CARD_CLAIM') THEN amount ELSE 0 END),0) expense_claim_add,
          COALESCE(SUM(CASE WHEN active=1 AND movement_type='ALACAK EKLE' AND source_type='PROFIT_DISTRIBUTION' THEN amount ELSE 0 END),0) profit_claim_add
          FROM partner_movements WHERE partner_id=?""",(d["id"],)).fetchone()
        con.close()
        b=partner_claim_breakdown(d["id"])
        d.update({
            "out":float(s[0]),"claim_add":float(s[1]),"claim_paid":float(s[2]),
            "expense_claim_add":float(s[3]),"profit_claim_add":float(s[4]),
            "expense_claim_left":b["expense_left"],"profit_claim_left":b["profit_left"],
            "other_claim_left":b["other_left"],"claim_left":b["total_left"]
        })
        out.append(d)
    return out


def partner_movement_rows(start=None,end=None):
    con=db_conn(); args=[]; sql="SELECT pm.*,p.name partner_name FROM partner_movements pm JOIN partners p ON p.id=pm.partner_id WHERE pm.active=1"
    if start and end: sql+=" AND pm.tx_date BETWEEN ? AND ?"; args=[start,end]
    sql+=" ORDER BY pm.tx_date DESC,pm.id DESC"; rows=[dict(r) for r in con.execute(sql,args).fetchall()]; con.close(); return rows


# ----------------------- Flask performance -----------------------

def csrf_token():
    if "csrf_token" not in session:
        session["csrf_token"] = secrets.token_urlsafe(32)
    return session["csrf_token"]


app.jinja_env.globals["csrf_token"] = csrf_token


@app.before_request
def validate_request():
    if request.endpoint == "static":
        return
    with ADMIN_SESSION_LOCK:
        expired_ids = [sid for sid,item in ADMIN_SESSIONS.items() if time.time()-item[1]>=ADMIN_IDLE_SECONDS]
        for sid in expired_ids:
            ADMIN_SESSIONS.pop(sid,None)
    was_admin = bool(session.get("admin"))
    if was_admin:
        try:
            last = admin_session_time()
        except (TypeError, ValueError):
            last = 0
        if not (0 < last <= time.time()) or time.time()-last >= ADMIN_IDLE_SECONDS:
            admin_session_remove()
            session.clear()
            flash("10 dakika işlem yapılmadığı için yönetici oturumu kapatıldı. Tekrar giriş yapın.", "warning")
            if request.endpoint in ("logs_api", "admin_activity", "admin_idle_status"):
                return jsonify(expired=True), 401
            if request.endpoint not in ("admin_login", "logout"):
                return redirect(url_for("admin_login"))
    if request.method in ("GET", "HEAD"):
        csrf_token()
    if session.get("admin"):
        con = db_conn()
        account = con.execute("SELECT active,is_owner,username FROM admin_accounts WHERE id=?", (session.get("admin_id"),)).fetchone()
        con.close()
        if not account or not account["active"]:
            admin_session_remove()
            session.clear()
            if request.endpoint in ("logs_api", "admin_activity", "admin_idle_status"):
                return jsonify(expired=True), 401
        else:
            session["admin_owner"] = bool(account["is_owner"])
            session["admin_username"] = account["username"]
    if request.method == "POST":
        supplied = request.form.get("csrf_token", "")
        expected = session.get("csrf_token", "")
        if not expected or not hmac.compare_digest(supplied, expected):
            abort(400, description="Güvenlik doğrulaması geçersiz. Sayfayı yenileyip tekrar deneyin.")
    if session.get("admin") and request.endpoint not in ("logs_api", "admin_idle_status", "admin_login", "admin_activity"):
        admin_session_touch()
    if request.endpoint == "admin_login" and not admin_accounts_exist() and request.remote_addr not in ("127.0.0.1", "::1"):
        abort(403, description="İlk yönetici hesabını bu bilgisayarda localhost üzerinden oluşturun.")


def admin_session_started():
    sid = secrets.token_urlsafe(32)
    session["admin_sid"] = sid
    with ADMIN_SESSION_LOCK:
        ADMIN_SESSIONS[sid] = (session["admin_id"], time.time())


def admin_session_time():
    with ADMIN_SESSION_LOCK:
        item = ADMIN_SESSIONS.get(session.get("admin_sid"))
        return item[1] if item and item[0] == session.get("admin_id") else 0


def admin_session_touch(age=0):
    with ADMIN_SESSION_LOCK:
        sid = session.get("admin_sid")
        item = ADMIN_SESSIONS.get(sid)
        if item and item[0] == session.get("admin_id"):
            ADMIN_SESSIONS[sid] = (item[0], max(item[1],time.time()-age))


def admin_session_remove():
    with ADMIN_SESSION_LOCK:
        ADMIN_SESSIONS.pop(session.get("admin_sid"),None)


def admin_idle_remaining():
    return max(0, ADMIN_IDLE_SECONDS-(time.time()-admin_session_time()))


app.jinja_env.globals["admin_idle_remaining"] = admin_idle_remaining


@app.post("/admin/activity")
@admin_required
def admin_activity():
    try:
        age = float(request.form.get("activity_age", "0"))
    except ValueError:
        return jsonify(error="Geçersiz etkinlik zamanı"),400
    if not 0 <= age <= 60:
        return jsonify(error="Geçersiz etkinlik zamanı"),400
    admin_session_touch(age)
    return jsonify(remaining=admin_idle_remaining())


@app.get("/admin/idle-status")
def admin_idle_status():
    if not session.get("admin"):
        return jsonify(expired=True), 401
    return jsonify(remaining=admin_idle_remaining())


@app.teardown_request
def cleanup_request(error):
    for con in g.pop("db_connections", []):
        con.close()


@app.errorhandler(400)
@app.errorhandler(403)
def request_error(error):
    return render_template("error.html", message=error.description, code=error.code), error.code


@app.before_request
def perf_start():
    # V10.6.3: web requests never wait for PanCafe / day-close / Excel maintenance.
    g.perf=time.perf_counter()

@app.after_request
def perf_end(resp):
    try:
        ms=int((time.perf_counter()-g.perf)*1000); resp.headers["X-KGA-Time-Ms"]=str(ms)
        # Disk logging every admin request was itself a major slowdown.
        # Keep only genuinely slow requests for diagnostics.
        if ms >= 400 and request.path.startswith("/admin") and request.path not in ("/admin/logs","/api/logs"):
            add_log("SLOW","SAYFA_SURE",f"{request.path} | {ms} ms",current_actor_name(),"PERF")
        if request.path.startswith("/admin"):
            resp.headers["Cache-Control"]="no-store, no-cache, must-revalidate, max-age=0"; resp.headers["Pragma"]="no-cache"; resp.headers["Expires"]="0"
    except Exception: pass
    resp.headers["X-Content-Type-Options"] = "nosniff"
    resp.headers["X-Frame-Options"] = "DENY"
    return resp


# ----------------------- Public/personnel routes -----------------------

@app.route("/")
def entry():
    today=accounting_day_iso()
    cash_date=cash_entry_default_date()
    if is_day_closed(cash_date):
        cash_date=today
    return render_template("entry.html", personnel=personnel_names(), categories=CATEGORIES, card_owners=CARD_OWNERS, today=today, cash_default_date=cash_date, cash_entry_grace=entry_grace_open(cash_date), day=day_status(today))

@app.post("/save-cash")
def save_cash():
    try:
        tx_date=parse_day(request.form.get("tarih")); ensure_day_open(tx_date)
        is_admin=bool(session.get("admin"))
        actor=current_actor_name() if is_admin else validate_personnel(request.form.get("personel"))
        pancafe=None  # V10.4: PanCafe GENEL TOPLAM yalnız otomatik alınır.
        con_chk=db_conn(); existing=con_chk.execute("SELECT id FROM cash_entries WHERE active=1 AND tx_date=? LIMIT 1",(tx_date,)).fetchone(); con_chk.close()
        if existing: raise RuntimeError(f"{tx_date} muhasebe günü için kasa kaydı zaten var. Günlük kasa kaydı tek olmalıdır.")
        vals={k:money(request.form.get(k)) for k in ("nakit","pos","iban","hatali","indirim","new_debt","old_debt_paid")}
        con=db_conn(); cur=con.execute("""INSERT INTO cash_entries(tx_date,created_at,pancafe,nakit,pos,iban,hatali,indirim,new_debt,old_debt_paid,note,actor_name,actor_type,created_by)
          VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",(tx_date,now_iso(),pancafe,vals["nakit"],vals["pos"],vals["iban"],vals["hatali"],vals["indirim"],vals["new_debt"],vals["old_debt_paid"],request.form.get("aciklama","").strip(),actor,"ADMIN" if is_admin else "PERSONEL",actor)); rid=cur.lastrowid; con.commit(); after=rowdict(con.execute("SELECT * FROM cash_entries WHERE id=?",(rid,)).fetchone()); con.close()
        add_audit("CASH",rid,"CREATE","Kasa kaydı oluşturuldu.",{},after,actor,"ADMIN" if is_admin else "PERSONEL"); schedule_excel_sync()
        flash("Kasa kalemleri kaydedildi. PanCafe GENEL TOPLAM 09:00 kapanışından sonra otomatik alınacak.","success")
    except Exception as e: flash(str(e),"danger")
    return redirect(url_for("entry"))

@app.post("/save-expense")
def save_expense():
    try:
        is_admin=bool(session.get("admin"))
        actor=current_actor_name() if is_admin else validate_personnel(request.form.get("personel"))
        tx_date=parse_day(request.form.get("tarih")); ensure_day_open(tx_date)
        category=request.form.get("kategori","").strip().upper(); amount=money(request.form.get("tutar")); source=request.form.get("odeme_kaynagi","İŞLETME").strip().upper(); owner=request.form.get("kart_sahibi","").strip(); note=request.form.get("aciklama","").strip()
        if category not in CATEGORIES or amount<=0: raise RuntimeError("Gider bilgileri geçersiz.")
        if source=="ŞAHSİ KART":
            mid=create_personal_card_claim(tx_date,category,amount,owner,note,actor,"ADMIN" if is_admin else "PERSONEL")
            schedule_excel_sync(); flash(f"Şahsi kart alacağı #{mid} oluşturuldu. Giderlere henüz eklenmedi; {owner}'a NAKİT ödeme yapıldıkça gider oluşacak.","success")
        else:
            con=db_conn(); cur=con.execute("""INSERT INTO expenses(tx_date,created_at,category,amount,note,payment_source,payment_method,card_owner,system_type,source_type,source_id,actor_name,actor_type,created_by)
             VALUES(?,?,?,?,?,'İŞLETME','-',?,'','',NULL,?,?,?)""",(tx_date,now_iso(),category,amount,note,"",actor,"ADMIN" if is_admin else "PERSONEL",actor)); rid=cur.lastrowid; con.commit(); after=rowdict(con.execute("SELECT * FROM expenses WHERE id=?",(rid,)).fetchone()); con.close()
            add_audit("EXPENSE",rid,"CREATE","Gider kaydı oluşturuldu.",{},after,actor,"ADMIN" if is_admin else "PERSONEL"); schedule_excel_sync(); flash("Gider kaydedildi.","success")
    except Exception as e: flash(str(e),"danger")
    return redirect(url_for("entry"))


@app.post("/save-credit")
def save_credit():
    try:
        tx_date=parse_day(request.form.get("tarih")); ensure_day_open(tx_date)
        is_admin=bool(session.get("admin")); actor=current_actor_name() if is_admin else validate_personnel(request.form.get("personel"))
        person=(request.form.get("kisi") or "").strip(); typ=(request.form.get("tur") or "").strip().upper(); amount=money(request.form.get("tutar")); method=(request.form.get("yontem") or "-").strip().upper()
        if not person: raise RuntimeError("Veresiye kişi adı boş olamaz.")
        if typ not in CREDIT_TYPES: raise RuntimeError("Veresiye türü geçersiz.")
        if amount<=0: raise RuntimeError("Tutar 0'dan büyük olmalı.")
        if typ=="BORÇ": method="-"
        elif method not in PAYMENT_METHODS or method=="-": raise RuntimeError("Ödemede NAKİT/POS/IBAN seçin.")
        con=db_conn(); cur=con.execute("INSERT INTO credit_entries(tx_date,created_at,person,movement_type,amount,payment_method,note,actor_name,actor_type,created_by) VALUES(?,?,?,?,?,?,?,?,?,?)",(tx_date,now_iso(),person,typ,amount,method,request.form.get("aciklama","").strip(),actor,"ADMIN" if is_admin else "PERSONEL",actor)); rid=cur.lastrowid; con.commit(); after=rowdict(con.execute("SELECT * FROM credit_entries WHERE id=?",(rid,)).fetchone()); con.close()
        add_audit("CREDIT",rid,"CREATE",f"Veresiye {typ} kaydı oluşturuldu.",{},after,actor,"ADMIN" if is_admin else "PERSONEL"); schedule_excel_sync(); flash("Veresiye kaydı oluşturuldu.","success")
    except Exception as e: flash(str(e),"danger")
    return redirect(url_for("entry"))


@app.route("/personel/mesai")
def personnel_attendance_page():
    people = personnel_rows(True)
    defaults = {str(p["id"]): (p.get("default_shift") or "") for p in people}
    return render_template(
        "personnel_attendance.html",
        personnel=people,
        shifts=SHIFT_DEFINITIONS,
        defaults=defaults,
        today=accounting_day_iso(),
    )


@app.post("/personel/mesai/kaydet")
def personnel_attendance_save():
    try:
        person = get_personnel_by_name(request.form.get("personel"))
        if not person:
            raise RuntimeError("Personel bulunamadı.")
        work_date = parse_day(request.form.get("tarih"))
        ensure_day_open(work_date)
        shift_code = (request.form.get("vardiya") or "").strip()
        info = shift_info(shift_code)
        actual_start = (request.form.get("gercek_giris") or info["start"]).strip()
        actual_end = (request.form.get("gercek_cikis") or info["end"]).strip()
        mins = _clock_minutes(actual_start, actual_end)
        note = (request.form.get("aciklama") or "").strip()
        con = db_conn()
        old = con.execute(
            "SELECT * FROM personnel_attendance WHERE personnel_id=? AND work_date=? AND active=1",
            (person["id"], work_date)
        ).fetchone()
        if old:
            before = dict(old)
            con.execute(
                """UPDATE personnel_attendance SET shift_code=?,scheduled_start=?,scheduled_end=?,actual_start=?,actual_end=?,
                   worked_minutes=?,note=?,updated_at=?,updated_by=? WHERE id=?""",
                (shift_code, info["start"], info["end"], actual_start, actual_end, mins, note, now_iso(), person["name"], old["id"])
            )
            rid = old["id"]
            action = "UPDATE"
            summary = f"{work_date} puantaj kaydı personel tarafından güncellendi."
        else:
            before = {}
            cur = con.execute(
                """INSERT INTO personnel_attendance(personnel_id,work_date,shift_code,scheduled_start,scheduled_end,
                   actual_start,actual_end,worked_minutes,note,created_at,created_by)
                   VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                (person["id"], work_date, shift_code, info["start"], info["end"], actual_start, actual_end, mins, note, now_iso(), person["name"])
            )
            rid = cur.lastrowid
            action = "CREATE"
            summary = f"{work_date} puantaj kaydı personel tarafından oluşturuldu."
        con.commit()
        after = rowdict(con.execute("SELECT * FROM personnel_attendance WHERE id=?", (rid,)).fetchone())
        con.close()
        # Puantaj finansal hareket değildir; Excel senkronu tetiklenmez. Böylece günlük kullanım paneli yavaşlatmaz.
        add_audit("PERSONNEL_ATTENDANCE", rid, action, summary, before, after, person["name"], "PERSONEL")
        flash(f"Mesai kaydedildi: {mins/60:.2f} saat.", "success")
    except Exception as e:
        flash(str(e), "danger")
    return redirect(url_for("personnel_attendance_page"))


# ----------------------- Admin login/dashboard -----------------------

@app.route("/admin-login", methods=["GET","POST"])
def admin_login():
    setup=not admin_accounts_exist()
    if request.method=="POST":
        try:
            username=(request.form.get("username") or "").strip(); password=request.form.get("password") or ""
            if setup:
                confirm=request.form.get("password_confirm") or ""
                if password!=confirm: raise RuntimeError("Şifreler aynı değil.")
                aid=create_admin(username,password,True,"İLK_KURULUM"); session.clear(); session.update(admin=True,admin_id=aid,admin_username=username,admin_owner=True,admin_last_activity=time.time()); admin_session_started(); add_audit("ADMIN",aid,"CREATE","Ana yönetici hesabı oluşturuldu.",{}, {"username":username},username,"ADMIN"); schedule_excel_sync(); flash("Ana yönetici hesabı oluşturuldu.","success"); return redirect(url_for("admin"))
            account=authenticate_admin(username,password)
            if not account: raise RuntimeError("Yönetici kullanıcı adı veya şifre yanlış.")
            session.clear(); session.update(admin=True,admin_id=account["id"],admin_username=account["username"],admin_owner=bool(account["is_owner"]),admin_last_activity=time.time()); admin_session_started(); add_log("INFO","ADMIN_GIRIS",username,username,"ADMIN"); return redirect(url_for("admin"))
        except Exception as e: flash(str(e),"danger")
    return render_template("login.html", setup_mode=setup)

@app.route("/logout")
def logout():
    admin_session_remove()
    session.clear(); return redirect(url_for("entry"))

@app.route("/admin")
@admin_required
def admin():
    month=request.args.get("month") or accounting_day().strftime("%Y-%m"); start,end=month_bounds(month); p=profit_loss(start,end); cash=cash_rows(start,end)[:8]; exp=expense_rows(start,end)[:8]; today=accounting_day_iso()
    yesterday=(accounting_day()-timedelta(days=1)).isoformat()
    return render_template(
        "admin.html",month=month,p=p,recent_cash=cash,recent_expenses=exp,has_personnel=bool(personnel_names()),
        day=day_status(today),today=today,
        yesterday=yesterday,yesterday_status=day_status(yesterday),
        last_sync=meta_get("last_excel_sync","Henüz yok"),
        sync_error=meta_get("excel_sync_error",""),
        last_backup=meta_get("last_backup","Henüz yok"),
        pancafe_auto=pancafe_auto_status()
    )


# ----------------------- Cash admin -----------------------

@app.route("/admin/cash")
@admin_required
def admin_cash():
    month=request.args.get("month") or accounting_day().strftime("%Y-%m"); start,end=month_bounds(month); rows=cash_rows(start,end)
    totals={"pancafe":sum(r["pancafe"] or 0 for r in rows),"calc":sum(r["calc"] for r in rows),"diff":sum(r["diff"] or 0 for r in rows),"waiting":sum(1 for r in rows if r["pancafe"] is None),"new_debt":sum(r["new_debt"] for r in rows),"old_debt_paid":sum(r["old_debt_paid"] for r in rows)}
    return render_template("cash.html",month=month,rows=rows,totals=totals,pancafe_auto=pancafe_auto_status(),accounting_today=accounting_day_iso())

@app.post("/admin/cash/pancafe-auto/<tx_date>")
@admin_required
def cash_pancafe_auto(tx_date):
    try:
        tx_date=parse_day(tx_date)
        result=auto_fill_pancafe_for_day(tx_date,force=True)
        if not result.get("ok"):
            raise RuntimeError(result.get("message") or "PanCafe otomatik çekilemedi.")
        if result.get("changed"):
            auto_close_due_days()
            flash(f"PanCafe GENEL TOPLAM otomatik alındı: {result['total']:,.2f} TL","success")
        else:
            flash(result.get("message") or "PanCafe zaten dolu.","success")
    except Exception as e:
        flash(str(e),"danger")
    return redirect(url_for("admin_cash",month=tx_date[:7] if tx_date else accounting_day_iso()[:7]))


@app.post("/admin/cash/pancafe/<int:rid>")
@admin_required
def cash_set_pancafe(rid):
    flash("V10.6.3'te PanCafe GENEL TOPLAM elle girilmez; PanCafe'den otomatik alınır.","warning")
    return redirect(url_for("admin_cash",month=request.form.get("month") or accounting_day_iso()[:7]))

@app.route("/admin/cash/edit/<int:rid>", methods=["GET","POST"])
@admin_required
def cash_edit(rid):
    con=db_conn(); row=con.execute("SELECT * FROM cash_entries WHERE id=? AND active=1",(rid,)).fetchone(); con.close()
    if not row: flash("Kayıt bulunamadı.","danger"); return redirect(url_for("admin_cash"))
    if request.method=="POST":
        try:
            ensure_day_open(row["tx_date"]); new_date=parse_day(request.form.get("tarih")); ensure_day_open(new_date); expected=int(request.form.get("version") or row["version"]); pancafe=row["pancafe"] if new_date==row["tx_date"] else None; vals={k:money(request.form.get(k)) for k in ("nakit","pos","iban","hatali","indirim","new_debt","old_debt_paid")}
            con=db_conn(); before=dict(con.execute("SELECT * FROM cash_entries WHERE id=?",(rid,)).fetchone()); cur=con.execute("""UPDATE cash_entries SET tx_date=?,pancafe=?,nakit=?,pos=?,iban=?,hatali=?,indirim=?,new_debt=?,old_debt_paid=?,note=?,updated_at=?,updated_by=?,version=version+1 WHERE id=? AND version=? AND active=1""",(new_date,pancafe,vals["nakit"],vals["pos"],vals["iban"],vals["hatali"],vals["indirim"],vals["new_debt"],vals["old_debt_paid"],request.form.get("aciklama","").strip(),now_iso(),current_actor_name(),rid,expected))
            if cur.rowcount!=1: con.rollback(); con.close(); raise RuntimeError("Kayıt başka bir yönetici tarafından değiştirildi. Sayfayı yenileyin.")
            con.commit(); after=rowdict(con.execute("SELECT * FROM cash_entries WHERE id=?",(rid,)).fetchone()); con.close(); add_audit("CASH",rid,"UPDATE","Kasa kaydı düzenlendi.",before,after); schedule_excel_sync(); flash("Kasa kaydı güncellendi.","success"); return redirect(url_for("admin_cash"))
        except sqlite3.IntegrityError:
            flash("Seçilen muhasebe gününde zaten bir kasa kaydı var. Farklı bir gün seçin.","danger")
        except Exception as e: flash(str(e),"danger")
    return render_template("cash_edit.html",data=dict(row))

@app.post("/admin/cash/cancel/<int:rid>")
@admin_required
def cash_cancel(rid):
    try:
        reason=(request.form.get("reason") or "").strip();
        if not reason: raise RuntimeError("İptal nedeni zorunlu.")
        con=db_conn(); row=con.execute("SELECT * FROM cash_entries WHERE id=? AND active=1",(rid,)).fetchone();
        if not row: con.close(); raise RuntimeError("Kayıt bulunamadı.")
        ensure_day_open(row["tx_date"]); before=dict(row); con.execute("UPDATE cash_entries SET active=0,canceled_at=?,canceled_by=?,cancel_reason=?,version=version+1 WHERE id=?",(now_iso(),current_actor_name(),reason,rid)); con.commit(); after=rowdict(con.execute("SELECT * FROM cash_entries WHERE id=?",(rid,)).fetchone()); con.close(); add_audit("CASH",rid,"CANCEL",f"Kasa kaydı iptal edildi: {reason}",before,after); schedule_excel_sync(); flash("Kasa kaydı silinmedi; iptal edildi ve audit geçmişine yazıldı.","success")
    except Exception as e: flash(str(e),"danger")
    return redirect(url_for("admin_cash"))


# ----------------------- Expenses admin -----------------------

@app.route("/admin/expenses")
@admin_required
def admin_expenses():
    month=request.args.get("month") or accounting_day().strftime("%Y-%m"); start,end=month_bounds(month); category=request.args.get("category","TÜMÜ"); rows=expense_rows(start,end,category); p=profit_loss(start,end); partners=partner_data();
    totals={c:0.0 for c in CATEGORIES}
    for r in rows:
        if r["category"] in totals: totals[r["category"]]+=r["amount"]
    return render_template("expenses.html",month=month,category=category,categories=CATEGORIES,rows=rows,totals=totals,p=p,partners=partners,card_owners=CARD_OWNERS,today=accounting_day_iso())

@app.post("/admin/expenses/add")
@admin_required
def expense_add():
    try:
        tx_date=parse_day(request.form.get("tarih")); ensure_day_open(tx_date); category=request.form.get("kategori","").strip().upper(); amount=money(request.form.get("tutar")); source=request.form.get("odeme_kaynagi","İŞLETME").strip().upper(); owner=request.form.get("kart_sahibi","").strip(); note=request.form.get("aciklama","").strip()
        if category not in CATEGORIES or amount<=0: raise RuntimeError("Gider bilgileri geçersiz.")
        if source=="ŞAHSİ KART":
            mid=create_personal_card_claim(tx_date,category,amount,owner,note,current_actor_name(),"ADMIN")
            schedule_excel_sync(); flash(f"Şahsi kart alacağı #{mid} oluşturuldu. Giderlere ödeme yapılana kadar eklenmeyecek.","success")
        else:
            con=db_conn(); cur=con.execute("INSERT INTO expenses(tx_date,created_at,category,amount,note,payment_source,payment_method,card_owner,actor_name,actor_type,created_by) VALUES(?,?,?,?,?,'İŞLETME','-','',?, 'ADMIN',?)",(tx_date,now_iso(),category,amount,note,current_actor_name(),current_actor_name())); rid=cur.lastrowid; con.commit(); after=rowdict(con.execute("SELECT * FROM expenses WHERE id=?",(rid,)).fetchone()); con.close(); add_audit("EXPENSE",rid,"CREATE","Yönetici gider kaydı oluşturdu.",{},after); schedule_excel_sync(); flash("Gider eklendi.","success")
    except Exception as e: flash(str(e),"danger")
    return redirect(url_for("admin_expenses"))


@app.route("/admin/expenses/edit/<int:rid>",methods=["GET","POST"])
@admin_required
def expense_edit(rid):
    con=db_conn(); row=con.execute("SELECT * FROM expenses WHERE id=? AND active=1",(rid,)).fetchone(); con.close()
    if not row: flash("Gider bulunamadı.","danger"); return redirect(url_for("admin_expenses"))
    if row["system_type"]==PERSONAL_CARD_REIMBURSEMENT_TYPE: flash("Bu gider şahsi kart alacağına yapılan NAKİT ödemeden otomatik oluşur. Ortaklar ekranındaki ödeme/alacak hareketini düzenleyin.","warning"); return redirect(url_for("admin_partners"))
    if row["system_type"]: flash("Bu sistem kaydı doğrudan düzenlenemez.","warning"); return redirect(url_for("admin_partners"))
    if request.method=="POST":
        try:
            ensure_day_open(row["tx_date"]); tx_date=parse_day(request.form.get("tarih")); ensure_day_open(tx_date); category=request.form.get("kategori","").strip().upper(); amount=money(request.form.get("tutar")); source=request.form.get("odeme_kaynagi","İŞLETME").strip().upper(); owner=request.form.get("kart_sahibi","").strip(); expected=int(request.form.get("version") or row["version"]); note=request.form.get("aciklama","").strip()
            if category not in CATEGORIES or amount<=0: raise RuntimeError("Bilgiler geçersiz.")
            if source=="ŞAHSİ KART":
                mid=create_personal_card_claim(tx_date,category,amount,owner,note,current_actor_name(),"ADMIN")
                con=db_conn(); before=dict(con.execute("SELECT * FROM expenses WHERE id=?",(rid,)).fetchone()); con.execute("UPDATE expenses SET active=0,canceled_at=?,canceled_by=?,cancel_reason='Şahsi kart alacağına dönüştürüldü',version=version+1 WHERE id=?",(now_iso(),current_actor_name(),rid)); con.commit(); after=rowdict(con.execute("SELECT * FROM expenses WHERE id=?",(rid,)).fetchone()); con.close(); add_audit("EXPENSE",rid,"CONVERT_TO_CARD_CLAIM",f"Gider şahsi kart alacağı #{mid} olarak taşındı.",before,after); schedule_excel_sync(); flash("Kayıt şahsi kart alacağına dönüştürüldü; giderlerden çıkarıldı.","success"); return redirect(url_for("admin_partners"))
            con=db_conn(); before=dict(con.execute("SELECT * FROM expenses WHERE id=?",(rid,)).fetchone()); cur=con.execute("UPDATE expenses SET tx_date=?,category=?,amount=?,note=?,payment_source='İŞLETME',card_owner='',updated_at=?,updated_by=?,version=version+1 WHERE id=? AND version=? AND active=1",(tx_date,category,amount,note,now_iso(),current_actor_name(),rid,expected))
            if cur.rowcount!=1: con.rollback(); con.close(); raise RuntimeError("Kayıt başka yönetici tarafından değiştirildi.")
            con.commit(); after=rowdict(con.execute("SELECT * FROM expenses WHERE id=?",(rid,)).fetchone()); con.close(); add_audit("EXPENSE",rid,"UPDATE","Gider kaydı düzenlendi.",before,after); schedule_excel_sync(); flash("Gider güncellendi.","success"); return redirect(url_for("admin_expenses"))
        except Exception as e: flash(str(e),"danger")
    return render_template("expense_edit.html",data=dict(row),categories=CATEGORIES,card_owners=CARD_OWNERS)


@app.post("/admin/expenses/cancel/<int:rid>")
@admin_required
def expense_cancel(rid):
    try:
        reason=(request.form.get("reason") or "").strip();
        if not reason: raise RuntimeError("İptal nedeni zorunlu.")
        con=db_conn(); row=con.execute("SELECT * FROM expenses WHERE id=? AND active=1",(rid,)).fetchone();
        if not row: con.close(); raise RuntimeError("Gider bulunamadı.")
        if row["system_type"]: con.close(); raise RuntimeError("Otomatik oluşan gideri Ortaklar ekranındaki bağlı hareketten iptal edin.")
        ensure_day_open(row["tx_date"]); before=dict(row); con.execute("UPDATE expenses SET active=0,canceled_at=?,canceled_by=?,cancel_reason=?,version=version+1 WHERE id=?",(now_iso(),current_actor_name(),reason,rid)); con.commit(); after=rowdict(con.execute("SELECT * FROM expenses WHERE id=?",(rid,)).fetchone()); con.close(); add_audit("EXPENSE",rid,"CANCEL",f"Gider iptal edildi: {reason}",before,after); schedule_excel_sync(); flash("Gider iptal edildi; kayıt audit geçmişinde korunuyor.","success")
    except Exception as e: flash(str(e),"danger")
    return redirect(url_for("admin_expenses"))


# ----------------------- Credit admin -----------------------

@app.route("/admin/credit")
@admin_required
def admin_credit():
    month=request.args.get("month") or accounting_day().strftime("%Y-%m"); start,end=month_bounds(month); return render_template("credit.html",month=month,rows=credit_rows(start,end),summary=credit_summary(),today=accounting_day_iso())

@app.post("/admin/credit/add")
@admin_required
def credit_add():
    # Reuse direct DB logic with admin actor
    try:
        tx_date=parse_day(request.form.get("tarih")); ensure_day_open(tx_date); person=request.form.get("kisi","").strip(); typ=request.form.get("tur","").strip().upper(); amount=money(request.form.get("tutar")); method=request.form.get("yontem","-").strip().upper()
        if not person or typ not in CREDIT_TYPES or amount<=0: raise RuntimeError("Veresiye bilgileri geçersiz.")
        if typ=="BORÇ": method="-"
        elif method not in PAYMENT_METHODS or method=="-": raise RuntimeError("Ödemede NAKİT/POS/IBAN seçin.")
        con=db_conn(); cur=con.execute("INSERT INTO credit_entries(tx_date,created_at,person,movement_type,amount,payment_method,note,actor_name,actor_type,created_by) VALUES(?,?,?,?,?,?,?,?, 'ADMIN',?)",(tx_date,now_iso(),person,typ,amount,method,request.form.get("aciklama","").strip(),current_actor_name(),current_actor_name())); rid=cur.lastrowid; con.commit(); after=rowdict(con.execute("SELECT * FROM credit_entries WHERE id=?",(rid,)).fetchone()); con.close(); add_audit("CREDIT",rid,"CREATE",f"Veresiye {typ} eklendi.",{},after); schedule_excel_sync(); flash("Veresiye hareketi eklendi.","success")
    except Exception as e: flash(str(e),"danger")
    return redirect(url_for("admin_credit"))

@app.route("/admin/credit/edit/<int:rid>",methods=["GET","POST"])
@admin_required
def credit_edit(rid):
    con=db_conn(); row=con.execute("SELECT * FROM credit_entries WHERE id=? AND active=1",(rid,)).fetchone(); con.close()
    if not row: flash("Veresiye kaydı bulunamadı.","danger"); return redirect(url_for("admin_credit"))
    if request.method=="POST":
        try:
            ensure_day_open(row["tx_date"]); tx_date=parse_day(request.form.get("tarih")); ensure_day_open(tx_date); person=request.form.get("kisi","").strip(); typ=request.form.get("tur","").strip().upper(); amount=money(request.form.get("tutar")); method=request.form.get("yontem","-").strip().upper(); expected=int(request.form.get("version") or row["version"])
            if not person or typ not in CREDIT_TYPES or amount<=0: raise RuntimeError("Bilgiler geçersiz.")
            if typ=="BORÇ": method="-"
            elif method not in PAYMENT_METHODS or method=="-": raise RuntimeError("Ödemede NAKİT/POS/IBAN seçin.")
            con=db_conn(); before=dict(con.execute("SELECT * FROM credit_entries WHERE id=?",(rid,)).fetchone()); cur=con.execute("UPDATE credit_entries SET tx_date=?,person=?,movement_type=?,amount=?,payment_method=?,note=?,updated_at=?,updated_by=?,version=version+1 WHERE id=? AND version=? AND active=1",(tx_date,person,typ,amount,method,request.form.get("aciklama","").strip(),now_iso(),current_actor_name(),rid,expected))
            if cur.rowcount!=1: con.rollback(); con.close(); raise RuntimeError("Kayıt başka bir yönetici tarafından değiştirildi.")
            con.commit(); after=rowdict(con.execute("SELECT * FROM credit_entries WHERE id=?",(rid,)).fetchone()); con.close(); add_audit("CREDIT",rid,"UPDATE","Veresiye kaydı düzenlendi.",before,after); schedule_excel_sync(); flash("Veresiye güncellendi.","success"); return redirect(url_for("admin_credit"))
        except Exception as e: flash(str(e),"danger")
    return render_template("credit_edit.html",data=dict(row))

@app.post("/admin/credit/cancel/<int:rid>")
@admin_required
def credit_cancel(rid):
    try:
        reason=request.form.get("reason","").strip();
        if not reason: raise RuntimeError("İptal nedeni zorunlu.")
        con=db_conn(); row=con.execute("SELECT * FROM credit_entries WHERE id=? AND active=1",(rid,)).fetchone();
        if not row: con.close(); raise RuntimeError("Kayıt bulunamadı.")
        ensure_day_open(row["tx_date"]); before=dict(row); con.execute("UPDATE credit_entries SET active=0,canceled_at=?,canceled_by=?,cancel_reason=?,version=version+1 WHERE id=?",(now_iso(),current_actor_name(),reason,rid)); con.commit(); after=rowdict(con.execute("SELECT * FROM credit_entries WHERE id=?",(rid,)).fetchone()); con.close(); add_audit("CREDIT",rid,"CANCEL",f"Veresiye iptal edildi: {reason}",before,after); schedule_excel_sync(); flash("Veresiye hareketi iptal edildi.","success")
    except Exception as e: flash(str(e),"danger")
    return redirect(url_for("admin_credit"))


# ----------------------- Partners admin -----------------------

@app.route("/admin/partners")
@admin_required
def admin_partners():
    partners=partner_data()
    movements=partner_movement_rows()
    profit_month=request.args.get("profit_month") or (accounting_day().replace(day=1)-timedelta(days=1)).strftime("%Y-%m")
    try:
        reserve=money(request.args.get("reserve"))
    except RuntimeError as e:
        flash(str(e), "danger"); reserve=0
    try:
        profit_preview=profit_distribution_preview(profit_month,reserve)
    except Exception as e:
        profit_preview={"month":profit_month,"error":str(e),"lines":[],"ready":False}
    return render_template(
        "partners.html",partners=partners,movements=movements,
        movement_types=PARTNER_MOVEMENT_TYPES,payment_methods=["NAKİT"],
        payment_targets=["AUTO","MASRAF","KAR_PAYI"],
        today=accounting_day_iso(),
        total_share=sum(p["share_percent"] for p in partners),
        total_out=sum(p["out"] for p in partners),
        total_claim_add=sum(p["claim_add"] for p in partners),
        total_claim_paid=sum(p["claim_paid"] for p in partners),
        total_claim_left=sum(p["claim_left"] for p in partners),
        total_expense_claim_left=sum(p["expense_claim_left"] for p in partners),
        total_profit_claim_left=sum(p["profit_claim_left"] for p in partners),
        profit_preview=profit_preview,
        profit_history=profit_distribution_history()
    )


@app.post("/admin/partners/profit-distribute")
@owner_required
def profit_distribute():
    try:
        period_month=(request.form.get("profit_month") or "").strip()
        reserve=money(request.form.get("reserve"))
        note=request.form.get("note","").strip()
        result=finalize_profit_distribution(period_month,reserve,note)
        flash(
            f"{period_month} kâr dağıtımı kesinleşti. Ortaklara {result['distributable_profit']:,.2f} TL kâr payı alacağı yazıldı.",
            "success"
        )
    except Exception as e:
        flash(str(e),"danger")
    return redirect(url_for("admin_partners",profit_month=request.form.get("profit_month"),reserve=request.form.get("reserve")))


@app.post("/admin/partners/add")
@admin_required
def partner_add():
    try:
        name=request.form.get("name","").strip(); share=money(request.form.get("share"));
        if share>100: raise RuntimeError("Ortak payı %100 üzerinde olamaz.")

        if not name: raise RuntimeError("Ortak adı gerekli.")
        con=db_conn(); row=con.execute("SELECT * FROM partners WHERE name=? COLLATE NOCASE",(name,)).fetchone()
        if row and row["active"]: con.close(); raise RuntimeError("Bu ortak zaten var.")
        if row:
            before=dict(row); con.execute("UPDATE partners SET active=1,share_percent=?,note=?,updated_at=?,updated_by=?,canceled_at=NULL,canceled_by=NULL,cancel_reason='',version=version+1 WHERE id=?",(share,request.form.get("note","").strip(),now_iso(),current_actor_name(),row["id"])); pid=row["id"]
        else:
            cur=con.execute("INSERT INTO partners(name,share_percent,note,active,created_at) VALUES(?,?,?,1,?)",(name,share,request.form.get("note","").strip(),now_iso())); pid=cur.lastrowid; before={}
        con.commit(); after=rowdict(con.execute("SELECT * FROM partners WHERE id=?",(pid,)).fetchone()); con.close(); add_audit("PARTNER",pid,"CREATE" if not before else "REACTIVATE","Ortak kaydı oluşturuldu/aktif edildi.",before,after); schedule_excel_sync(); flash("Ortak kaydedildi.","success")
    except Exception as e: flash(str(e),"danger")
    return redirect(url_for("admin_partners"))

@app.route("/admin/partners/edit/<int:rid>",methods=["GET","POST"])
@admin_required
def partner_edit(rid):
    con=db_conn(); row=con.execute("SELECT * FROM partners WHERE id=? AND active=1",(rid,)).fetchone(); con.close()
    if not row: flash("Ortak bulunamadı.","danger"); return redirect(url_for("admin_partners"))
    if request.method=="POST":
        try:
            name=request.form.get("name","").strip(); share=money(request.form.get("share"));
            if not name or share>100: raise RuntimeError("Ortak adı ve %0–100 arasında pay girin.")
            expected=int(request.form.get("version") or row["version"]); con=db_conn(); before=dict(con.execute("SELECT * FROM partners WHERE id=?",(rid,)).fetchone()); cur=con.execute("UPDATE partners SET name=?,share_percent=?,note=?,updated_at=?,updated_by=?,version=version+1 WHERE id=? AND version=? AND active=1",(name,share,request.form.get("note","").strip(),now_iso(),current_actor_name(),rid,expected))
            if cur.rowcount!=1: con.rollback(); con.close(); raise RuntimeError("Kayıt başka yönetici tarafından değiştirildi.")
            con.commit(); after=rowdict(con.execute("SELECT * FROM partners WHERE id=?",(rid,)).fetchone()); con.close(); add_audit("PARTNER",rid,"UPDATE","Ortak bilgileri güncellendi.",before,after); schedule_excel_sync(); flash("Ortak güncellendi.","success"); return redirect(url_for("admin_partners"))
        except Exception as e: flash(str(e),"danger")
    return render_template("partner_edit.html",data=dict(row))

@app.post("/admin/partners/cancel/<int:rid>")
@admin_required
def partner_cancel(rid):
    try:
        reason=request.form.get("reason","").strip();
        if not reason: raise RuntimeError("İptal nedeni zorunlu.")
        if abs(partner_balance(rid))>0.01: raise RuntimeError("Alacağı/borcu sıfır olmayan ortak iptal edilemez.")
        con=db_conn(); row=con.execute("SELECT * FROM partners WHERE id=? AND active=1",(rid,)).fetchone();
        if not row: con.close(); raise RuntimeError("Ortak bulunamadı.")
        before=dict(row); con.execute("UPDATE partners SET active=0,canceled_at=?,canceled_by=?,cancel_reason=?,version=version+1 WHERE id=?",(now_iso(),current_actor_name(),reason,rid)); con.commit(); after=rowdict(con.execute("SELECT * FROM partners WHERE id=?",(rid,)).fetchone()); con.close(); add_audit("PARTNER",rid,"CANCEL",f"Ortak iptal edildi: {reason}",before,after); schedule_excel_sync(); flash("Ortak iptal edildi.","success")
    except Exception as e: flash(str(e),"danger")
    return redirect(url_for("admin_partners"))

@app.post("/admin/partners/movement/add")
@admin_required
def partner_movement_add():
    try:
        tx_date=parse_day(request.form.get("date")); ensure_day_open(tx_date); pid=int(request.form.get("partner_id")); typ=request.form.get("type","").strip().upper(); amount=money(request.form.get("amount")); method=request.form.get("payment_method","-").strip().upper(); target=request.form.get("payment_target","AUTO").strip().upper()
        if typ not in PARTNER_MOVEMENT_TYPES or amount<=0: raise RuntimeError("Ortak hareket bilgileri geçersiz.")
        if typ=="ALACAK ÖDEME":
            if method != "NAKİT": raise RuntimeError("Ortak alacak ödemesi yalnız NAKİT olabilir.")
            if target not in ("AUTO","MASRAF","KAR_PAYI"): raise RuntimeError("Mahsup türü geçersiz.")
            b=partner_claim_breakdown(pid,until_date=tx_date)
            available=b["total_left"] if target=="AUTO" else (b["expense_left"] if target=="MASRAF" else b["profit_left"])
            if amount-available>0.01: raise RuntimeError(f"Ödeme seçilen alacak bakiyesinden fazla. Kullanılabilir: {available:,.2f} TL")
        else:
            method="-"; target="AUTO"
        con=db_conn(); cur=con.execute("INSERT INTO partner_movements(tx_date,created_at,partner_id,movement_type,amount,payment_method,payment_target,note,actor_name,actor_type,active) VALUES(?,?,?,?,?,?,?,?,?, 'ADMIN',1)",(tx_date,now_iso(),pid,typ,amount,method,target,request.form.get("note","").strip(),current_actor_name())); rid=cur.lastrowid; con.commit(); after=rowdict(con.execute("SELECT * FROM partner_movements WHERE id=?",(rid,)).fetchone()); con.close(); add_audit("PARTNER_MOVEMENT",rid,"CREATE",f"Ortak hareketi {typ} oluşturuldu.",{},after); sync_partner_payment_expense(rid); schedule_excel_sync(); flash("Ortak hareketi eklendi." + (" Şahsi kart alacağına mahsup edilen kısım kendi kategorisiyle Giderler'e eklendi." if typ=="ALACAK ÖDEME" else ""),"success")
    except Exception as e: flash(str(e),"danger")
    return redirect(url_for("admin_partners"))

@app.route("/admin/partners/movement/edit/<int:rid>",methods=["GET","POST"])
@admin_required
def partner_movement_edit(rid):
    con=db_conn(); row=con.execute("SELECT * FROM partner_movements WHERE id=? AND active=1",(rid,)).fetchone(); con.close()
    if not row: flash("Hareket bulunamadı.","danger"); return redirect(url_for("admin_partners"))
    if row["source_type"]=="PROFIT_DISTRIBUTION": flash("Bu kayıt kesinleşmiş kâr dağıtımından otomatik oluşur ve doğrudan düzenlenemez.","warning"); return redirect(url_for("admin_partners"))
    if request.method=="POST":
        try:
            ensure_day_open(row["tx_date"]); old_pid=int(row["partner_id"]); tx_date=parse_day(request.form.get("date")); ensure_day_open(tx_date); pid=int(request.form.get("partner_id")); typ=request.form.get("type","").strip().upper(); amount=money(request.form.get("amount")); method=request.form.get("payment_method","-").strip().upper(); target=request.form.get("payment_target","AUTO").strip().upper(); expected=int(request.form.get("version") or row["version"]); claim_category=request.form.get("claim_category",row["claim_category"] or "").strip().upper()
            if row["source_type"]==PERSONAL_CARD_CLAIM_SOURCE:
                typ="ALACAK EKLE"; method="-"; target="AUTO"
                if claim_category not in CATEGORIES: raise RuntimeError("Şahsi kart alacağında kategori seçin.")
                already_paid=personal_card_claim_paid_amount(rid)
                if amount + 0.01 < already_paid:
                    raise RuntimeError(f"Bu alacağın {already_paid:,.2f} TL'si zaten NAKİT ödendi. Alacak tutarı bunun altına indirilemez.")
                if already_paid > 0.01 and pid != int(row["partner_id"]):
                    raise RuntimeError("Kısmen ödenmiş şahsi kart alacağının ortağı değiştirilemez. Önce ödeme hareketini iptal/düzeltin.")
                if already_paid > 0.01 and tx_date != row["tx_date"]:
                    raise RuntimeError("Kısmen ödenmiş şahsi kart alacağının tarihi değiştirilemez. Önce ödeme hareketini iptal/düzeltin.")
            if typ not in PARTNER_MOVEMENT_TYPES or amount<=0: raise RuntimeError("Bilgiler geçersiz.")
            if typ=="ALACAK ÖDEME":
                if method != "NAKİT": raise RuntimeError("Ortak alacak ödemesi yalnız NAKİT olabilir.")
                if target not in ("AUTO","MASRAF","KAR_PAYI"): raise RuntimeError("Mahsup türü geçersiz.")
                b=partner_claim_breakdown(pid,until_date=tx_date,exclude_movement_id=rid)
                available=b["total_left"] if target=="AUTO" else (b["expense_left"] if target=="MASRAF" else b["profit_left"])
                if amount-available>0.01: raise RuntimeError(f"Ödeme seçilen alacak bakiyesinden fazla. Kullanılabilir: {available:,.2f} TL")
            else:
                method="-"; target="AUTO"
            con=db_conn(); before=dict(con.execute("SELECT * FROM partner_movements WHERE id=?",(rid,)).fetchone()); cur=con.execute("UPDATE partner_movements SET tx_date=?,partner_id=?,movement_type=?,amount=?,payment_method=?,payment_target=?,note=?,claim_category=?,claim_origin=CASE WHEN source_type=? THEN 'ŞAHSİ KART' ELSE claim_origin END,updated_at=?,updated_by=?,version=version+1 WHERE id=? AND version=? AND active=1",(tx_date,pid,typ,amount,method,target,request.form.get("note","").strip(),claim_category,PERSONAL_CARD_CLAIM_SOURCE,now_iso(),current_actor_name(),rid,expected))
            if cur.rowcount!=1: con.rollback(); con.close(); raise RuntimeError("Kayıt başka yönetici tarafından değiştirildi.")
            con.commit(); after=rowdict(con.execute("SELECT * FROM partner_movements WHERE id=?",(rid,)).fetchone()); con.close(); add_audit("PARTNER_MOVEMENT",rid,"UPDATE","Ortak hareketi güncellendi.",before,after); rebuild_partner_payment_allocations(old_pid,current_actor_name()); rebuild_partner_payment_allocations(pid,current_actor_name()); schedule_excel_sync(); flash("Ortak hareketi güncellendi.","success"); return redirect(url_for("admin_partners"))
        except Exception as e: flash(str(e),"danger")
    return render_template("partner_movement_edit.html",data=dict(row),partners=partner_data(),movement_types=PARTNER_MOVEMENT_TYPES,payment_methods=["NAKİT"],payment_targets=["AUTO","MASRAF","KAR_PAYI"],categories=CATEGORIES)

@app.post("/admin/partners/movement/cancel/<int:rid>")
@admin_required
def partner_movement_cancel(rid):
    try:
        reason=request.form.get("reason","").strip();
        if not reason: raise RuntimeError("İptal nedeni zorunlu.")
        con=db_conn(); row=con.execute("SELECT * FROM partner_movements WHERE id=? AND active=1",(rid,)).fetchone();
        if not row: con.close(); raise RuntimeError("Hareket bulunamadı.")
        if row["source_type"]=="EXPENSE": con.close(); raise RuntimeError("Bu eski tip alacak doğrudan iptal edilemez.")
        if row["source_type"]=="PROFIT_DISTRIBUTION": con.close(); raise RuntimeError("Kesinleşmiş kâr payı alacağı doğrudan iptal edilemez.")
        if row["source_type"]==PERSONAL_CARD_CLAIM_SOURCE:
            paid=con.execute("SELECT COALESCE(SUM(amount),0) FROM partner_payment_allocations WHERE active=1 AND claim_movement_id=?",(rid,)).fetchone()[0]
            if float(paid or 0)>0.01:
                con.close(); raise RuntimeError(f"Bu şahsi kart alacağının {float(paid):,.2f} TL'si ödendi. Önce bağlı NAKİT ödeme hareketini iptal/düzeltin.")
        ensure_day_open(row["tx_date"]); before=dict(row); con.execute("UPDATE partner_movements SET active=0,canceled_at=?,canceled_by=?,cancel_reason=?,version=version+1 WHERE id=?",(now_iso(),current_actor_name(),reason,rid)); con.commit(); after=rowdict(con.execute("SELECT * FROM partner_movements WHERE id=?",(rid,)).fetchone()); con.close(); add_audit("PARTNER_MOVEMENT",rid,"CANCEL",f"Ortak hareketi iptal edildi: {reason}",before,after); sync_partner_payment_expense(rid); schedule_excel_sync(); flash("Ortak hareketi iptal edildi.","success")
    except Exception as e: flash(str(e),"danger")
    return redirect(url_for("admin_partners"))


# ----------------------- Day close -----------------------

@app.route("/admin/day-close")
@admin_required
def day_close_page():
    tx_date=request.args.get("date") or (accounting_day()-timedelta(days=1)).isoformat()
    tx_date=parse_day(tx_date)
    auto_result={}  # Day-lock screen is read-only; maintenance runs in the worker.
    con=db_conn()
    counts={
        "cash":con.execute("SELECT COUNT(*) FROM cash_entries WHERE active=1 AND tx_date=?",(tx_date,)).fetchone()[0],
        "expense":con.execute("SELECT COUNT(*) FROM expenses WHERE active=1 AND tx_date=?",(tx_date,)).fetchone()[0],
        "credit":con.execute("SELECT COUNT(*) FROM credit_entries WHERE active=1 AND tx_date=?",(tx_date,)).fetchone()[0],
        "partner":con.execute("SELECT COUNT(*) FROM partner_movements WHERE active=1 AND tx_date=?",(tx_date,)).fetchone()[0],
    }
    recent=[dict(r) for r in con.execute(
        """SELECT * FROM day_closures
           ORDER BY tx_date DESC LIMIT 14"""
    ).fetchall()]
    con.close()
    return render_template(
        "day_close.html",
        tx_date=tx_date,
        status=day_status(tx_date),
        waiting=day_waiting_pancafe(tx_date),
        counts=counts,
        recent=recent,
        auto_result=auto_result,
        today=accounting_day_iso(),
        cutoff=(accounting_day()-timedelta(days=1)).isoformat()
    )


@app.post("/admin/day-close/close")
@admin_required
def day_close_do():
    flash(
        "V10.6.3'te manuel Günü Kapat kaldırıldı. Sistem bugünü açık tutar; dün ve daha eski günleri otomatik kilitler.",
        "warning"
    )
    return redirect(url_for("day_close_page",date=request.form.get("date")))


@synchronized(DAY_AUTO_LOCK)
def reopen_days(start, end, reason):
    start, end = parse_day(start), parse_day(end)
    first, last = date.fromisoformat(start), date.fromisoformat(end)
    today = accounting_day()
    reason = (reason or "").strip()
    if not reason:
        raise RuntimeError("Gün açma nedeni zorunlu.")
    if first > last:
        raise RuntimeError("Başlangıç tarihi bitiş tarihinden sonra olamaz.")
    if last >= today:
        raise RuntimeError("Yalnız geçmiş muhasebe günleri açılabilir. Bugün zaten açıktır.")
    count = (last-first).days+1
    if count > 366:
        raise RuntimeError("Tek işlemde en fazla 366 gün açılabilir.")
    actor, stamp = current_actor_name(), now_iso()
    # SQLite backup captures WAL data too. A failed backup prevents any change.
    backup_path = BACKUP_DIR / ("GunAcma_" + datetime.now().strftime("%Y%m%d_%H%M%S_%f") + "_" + secrets.token_hex(4) + ".db")
    con = db_conn()
    try:
        with sqlite3.connect(backup_path) as backup:
            con.backup(backup)
        con.execute("BEGIN IMMEDIATE")
        changed, skipped = [], []
        for offset in range(count):
            day = (first+timedelta(days=offset)).isoformat()
            before = rowdict(con.execute("SELECT * FROM day_closures WHERE tx_date=?", (day,)).fetchone()) or {}
            reopened_day = None
            try:
                reopened_day = accounting_day(datetime.fromisoformat(str(before.get("reopened_at") or "")))
            except ValueError:
                pass
            if before and not before["is_closed"] and reopened_day == today:
                skipped.append(day)
                continue
            con.execute("""INSERT INTO day_closures(tx_date,is_closed,closed_at,closed_by,close_note,reopened_at,reopened_by)
                VALUES(?,0,'','',?,?,?) ON CONFLICT(tx_date) DO UPDATE SET
                is_closed=0,reopened_at=excluded.reopened_at,reopened_by=excluded.reopened_by,
                close_note=COALESCE(day_closures.close_note,'') || excluded.close_note""",
                (day, " | Geçici yeniden açma: " + reason, stamp, actor))
            after = rowdict(con.execute("SELECT * FROM day_closures WHERE tx_date=?", (day,)).fetchone())
            con.execute("""INSERT INTO audit_log(created_at,actor_name,actor_type,entity_type,entity_id,action,summary,before_json,after_json)
                VALUES(?,?,'ADMIN','DAY',NULL,?,?,?,?)""",
                (stamp,actor,"BULK_REOPEN" if count>1 else "REOPEN",f"{day} düzeltme için geçici açıldı: {reason}",
                 json.dumps(before,ensure_ascii=False),json.dumps(after,ensure_ascii=False)))
            changed.append(day)
        con.commit()
        return {"changed": changed, "skipped": skipped, "backup": backup_path.name}
    except Exception:
        con.rollback()
        raise
    finally:
        con.close()


@app.post("/admin/day-close/reopen-bulk")
@owner_required
def day_reopen_bulk():
    try:
        result = reopen_days(request.form.get("start"), request.form.get("end"), request.form.get("reason"))
        schedule_excel_sync()
        flash(f"{len(result['changed'])} gün bugünlük açıldı; {len(result['skipped'])} gün zaten açıktı. "
              f"İşlem öncesi veritabanı yedeği: {result['backup']}. Açılan günler yarın tekrar kilitlenir.", "success")
    except Exception as e:
        flash(str(e), "danger")
    return redirect(url_for("day_close_page"))


@app.post("/admin/day-close/reopen")
@owner_required
def day_reopen_do():
    try:
        tx_date = parse_day(request.form.get("date"))
        result = reopen_days(tx_date, tx_date, request.form.get("reason"))
        schedule_excel_sync()
        flash(f"{tx_date} bugünlük açık. Yarın tekrar kilitlenir. İşlem öncesi yedek: {result['backup']}", "success")
    except Exception as e:
        flash(str(e), "danger")
    return redirect(url_for("day_close_page",date=request.form.get("date")))


# ----------------------- Personnel / payroll admin -----------------------

@app.route("/admin/personnel")
@admin_required
def admin_personnel():
    month = request.args.get("month") or accounting_day().strftime("%Y-%m")
    try:
        month_bounds(month)
    except Exception:
        month = accounting_day().strftime("%Y-%m")
    rows = personnel_month_summary(month)
    totals = personnel_month_totals(rows)
    attendance = personnel_attendance_rows(month)
    advances, payments = personnel_financial_rows(month)
    return render_template(
        "personnel_admin.html",
        month=month,
        rows=rows,
        totals=totals,
        attendance=attendance,
        advances=advances,
        payments=payments,
        shifts=SHIFT_DEFINITIONS,
        today=accounting_day_iso(),
        divisor=PERSONNEL_HOURS_DIVISOR,
        meal_hours=PERSONNEL_MEAL_BLOCK_HOURS,
        meal_amount=PERSONNEL_MEAL_AMOUNT,
    )


@app.post("/admin/personnel/attendance/<int:rid>/update")
@admin_required
def admin_personnel_attendance_update(rid):
    month = request.form.get("month") or accounting_day().strftime("%Y-%m")
    try:
        work_date = parse_day(request.form.get("tarih"))
        shift_code = (request.form.get("vardiya") or "").strip()
        info = shift_info(shift_code)
        actual_start = (request.form.get("gercek_giris") or "").strip()
        actual_end = (request.form.get("gercek_cikis") or "").strip()
        mins = _clock_minutes(actual_start, actual_end)
        note = (request.form.get("aciklama") or "").strip()
        con = db_conn()
        old = con.execute("SELECT * FROM personnel_attendance WHERE id=? AND active=1", (rid,)).fetchone()
        if not old:
            con.close(); raise RuntimeError("Puantaj kaydı bulunamadı.")
        before = dict(old)
        con.execute(
            """UPDATE personnel_attendance SET work_date=?,shift_code=?,scheduled_start=?,scheduled_end=?,actual_start=?,actual_end=?,
               worked_minutes=?,note=?,updated_at=?,updated_by=? WHERE id=?""",
            (work_date, shift_code, info["start"], info["end"], actual_start, actual_end, mins, note, now_iso(), current_actor_name(), rid)
        )
        con.commit(); after = rowdict(con.execute("SELECT * FROM personnel_attendance WHERE id=?", (rid,)).fetchone()); con.close()
        add_audit("PERSONNEL_ATTENDANCE", rid, "UPDATE", "Puantaj kaydı yönetici tarafından düzeltildi.", before, after)
        flash("Puantaj güncellendi.", "success")
    except Exception as e:
        flash(str(e), "danger")
    return redirect(url_for("admin_personnel", month=month))


@app.post("/admin/personnel/attendance/<int:rid>/cancel")
@admin_required
def admin_personnel_attendance_cancel(rid):
    month = request.form.get("month") or accounting_day().strftime("%Y-%m")
    try:
        reason = (request.form.get("reason") or "").strip()
        if not reason: raise RuntimeError("İptal nedeni zorunlu.")
        con = db_conn(); old = con.execute("SELECT * FROM personnel_attendance WHERE id=? AND active=1", (rid,)).fetchone()
        if not old: con.close(); raise RuntimeError("Puantaj kaydı bulunamadı.")
        before = dict(old)
        con.execute("UPDATE personnel_attendance SET active=0,canceled_at=?,canceled_by=?,cancel_reason=? WHERE id=?", (now_iso(), current_actor_name(), reason, rid))
        con.commit(); after = rowdict(con.execute("SELECT * FROM personnel_attendance WHERE id=?", (rid,)).fetchone()); con.close()
        add_audit("PERSONNEL_ATTENDANCE", rid, "CANCEL", f"Puantaj iptal edildi: {reason}", before, after)
        flash("Puantaj kaydı iptal edildi.", "success")
    except Exception as e: flash(str(e), "danger")
    return redirect(url_for("admin_personnel", month=month))


@app.post("/admin/personnel/advance/add")
@admin_required
def admin_personnel_advance_add():
    month = request.form.get("month") or accounting_day().strftime("%Y-%m")
    try:
        pid = int(request.form.get("personnel_id") or 0)
        tx_date = parse_day(request.form.get("tarih")); ensure_day_open(tx_date)
        amount = money(request.form.get("tutar")); note = (request.form.get("aciklama") or "").strip()
        if amount <= 0: raise RuntimeError("Avans 0'dan büyük olmalı.")
        con = db_conn(); p = con.execute("SELECT * FROM personnel WHERE id=? AND active=1", (pid,)).fetchone()
        if not p: con.close(); raise RuntimeError("Personel bulunamadı.")
        cur = con.execute(
            "INSERT INTO personnel_advances(personnel_id,tx_date,amount,payment_method,note,created_at,created_by) VALUES(?,?,?,'NAKİT',?,?,?)",
            (pid, tx_date, amount, note, now_iso(), current_actor_name())
        )
        aid = cur.lastrowid
        ecur = con.execute(
            """INSERT INTO expenses(tx_date,created_at,category,amount,note,payment_source,payment_method,card_owner,system_type,source_type,source_id,actor_name,actor_type,created_by)
               VALUES(?,?,?,?,?,'İŞLETME','NAKİT','',?,'PERSONNEL_ADVANCE',?,?, 'ADMIN',?)""",
            (tx_date, now_iso(), "ELEMAN", amount, f"{p['name']} · AVANS" + (f" · {note}" if note else ""), STAFF_ADVANCE_TYPE, aid, current_actor_name(), current_actor_name())
        )
        eid = ecur.lastrowid
        con.execute("UPDATE personnel_advances SET expense_id=? WHERE id=?", (eid, aid))
        con.commit(); after = rowdict(con.execute("SELECT * FROM personnel_advances WHERE id=?", (aid,)).fetchone()); con.close()
        add_audit("PERSONNEL_ADVANCE", aid, "CREATE", f"{p['name']} için {amount:.2f} TL NAKİT avans verildi ve ELEMAN giderine işlendi.", {}, after)
        schedule_excel_sync(); flash("Avans kaydedildi; Giderler > ELEMAN > AVANS olarak eklendi.", "success")
    except Exception as e: flash(str(e), "danger")
    return redirect(url_for("admin_personnel", month=month))


@app.post("/admin/personnel/advance/<int:rid>/cancel")
@admin_required
def admin_personnel_advance_cancel(rid):
    month = request.form.get("month") or accounting_day().strftime("%Y-%m")
    try:
        reason = (request.form.get("reason") or "").strip()
        if not reason: raise RuntimeError("İptal nedeni zorunlu.")
        con = db_conn(); row = con.execute("SELECT * FROM personnel_advances WHERE id=? AND active=1", (rid,)).fetchone()
        if not row: con.close(); raise RuntimeError("Avans bulunamadı.")
        ensure_day_open(row["tx_date"]); before = dict(row)
        con.execute("UPDATE personnel_advances SET active=0,canceled_at=?,canceled_by=?,cancel_reason=? WHERE id=?", (now_iso(), current_actor_name(), reason, rid))
        if row["expense_id"]:
            con.execute("UPDATE expenses SET active=0,canceled_at=?,canceled_by=?,cancel_reason=?,version=version+1 WHERE id=? AND active=1", (now_iso(), current_actor_name(), "Bağlı personel avansı iptal edildi: "+reason, row["expense_id"]))
        con.commit(); after = rowdict(con.execute("SELECT * FROM personnel_advances WHERE id=?", (rid,)).fetchone()); con.close()
        add_audit("PERSONNEL_ADVANCE", rid, "CANCEL", f"Personel avansı iptal edildi: {reason}", before, after)
        schedule_excel_sync(); flash("Avans ve bağlı ELEMAN gideri iptal edildi.", "success")
    except Exception as e: flash(str(e), "danger")
    return redirect(url_for("admin_personnel", month=month))


def _create_personnel_payment(pid, month, kind, tx_date, note=""):
    month_bounds(month)
    ensure_day_open(tx_date)
    con = db_conn()
    try:
        con.execute("BEGIN IMMEDIATE")
        summary = current_personnel_due(month, pid, con)
        if kind == "MAAŞ":
            amount = float(summary["salary_due"])
            method = "BANKA"
            system_type = STAFF_SALARY_TYPE
            desc = f"{summary['name']} · {month} MAAŞ · {summary['worked_hours']:.2f} saat"
        elif kind == "YEMEK":
            amount = float(summary["meal_due"])
            method = "YEMEK KARTI"
            system_type = STAFF_MEAL_TYPE
            desc = f"{summary['name']} · {month} YEMEK · {summary['meal_units']} adet x {PERSONNEL_MEAL_AMOUNT:.0f} TL"
        else:
            raise RuntimeError("Ödeme türü geçersiz.")
        if amount <= 0.005:
            raise RuntimeError(f"{kind} için ödenecek bakiye yok.")
        cur = con.execute(
            """INSERT INTO personnel_payments(personnel_id,period_month,tx_date,payment_kind,amount,payment_method,note,created_at,created_by)
               VALUES(?,?,?,?,?,?,?,?,?)""",
            (pid, month, tx_date, kind, amount, method, note, now_iso(), current_actor_name())
        )
        pay_id = cur.lastrowid
        ecur = con.execute(
            """INSERT INTO expenses(tx_date,created_at,category,amount,note,payment_source,payment_method,card_owner,system_type,source_type,source_id,actor_name,actor_type,created_by)
               VALUES(?,?,?,?,?,'İŞLETME',?,'',?,'PERSONNEL_PAYMENT',?,?, 'ADMIN',?)""",
            (tx_date, now_iso(), "ELEMAN", amount, desc + (f" · {note}" if note else ""), method, system_type, pay_id, current_actor_name(), current_actor_name())
        )
        eid = ecur.lastrowid
        con.execute("UPDATE personnel_payments SET expense_id=? WHERE id=?", (eid, pay_id))
        con.commit(); after = rowdict(con.execute("SELECT * FROM personnel_payments WHERE id=?", (pay_id,)).fetchone())
    except Exception:
        con.rollback()
        raise
    finally:
        con.close()
    add_audit("PERSONNEL_PAYMENT", pay_id, "CREATE", f"{summary['name']} {month} {kind}: {amount:.2f} TL / {method}; ELEMAN gideri oluşturuldu.", {}, after)
    schedule_excel_sync()
    return amount, method, summary["name"]


@app.post("/admin/personnel/pay")
@admin_required
def admin_personnel_pay():
    month = request.form.get("month") or accounting_day().strftime("%Y-%m")
    try:
        pid = int(request.form.get("personnel_id") or 0)
        kind = (request.form.get("kind") or "").strip().upper()
        tx_date = parse_day(request.form.get("tarih"))
        amount, method, name = _create_personnel_payment(pid, month, kind, tx_date, (request.form.get("aciklama") or "").strip())
        flash(f"{name}: {amount:.2f} TL {kind} / {method} ödendi; Giderler > ELEMAN'a ayrı kayıt eklendi.", "success")
    except Exception as e: flash(str(e), "danger")
    return redirect(url_for("admin_personnel", month=month))


@app.post("/admin/personnel/payment/<int:rid>/cancel")
@admin_required
def admin_personnel_payment_cancel(rid):
    month = request.form.get("month") or accounting_day().strftime("%Y-%m")
    try:
        reason = (request.form.get("reason") or "").strip()
        if not reason: raise RuntimeError("İptal nedeni zorunlu.")
        con = db_conn(); row = con.execute("SELECT * FROM personnel_payments WHERE id=? AND active=1", (rid,)).fetchone()
        if not row: con.close(); raise RuntimeError("Personel ödemesi bulunamadı.")
        ensure_day_open(row["tx_date"]); before = dict(row)
        con.execute("UPDATE personnel_payments SET active=0,canceled_at=?,canceled_by=?,cancel_reason=? WHERE id=?", (now_iso(), current_actor_name(), reason, rid))
        if row["expense_id"]:
            con.execute("UPDATE expenses SET active=0,canceled_at=?,canceled_by=?,cancel_reason=?,version=version+1 WHERE id=? AND active=1", (now_iso(), current_actor_name(), "Bağlı personel ödemesi iptal edildi: "+reason, row["expense_id"]))
        con.commit(); after = rowdict(con.execute("SELECT * FROM personnel_payments WHERE id=?", (rid,)).fetchone()); con.close()
        add_audit("PERSONNEL_PAYMENT", rid, "CANCEL", f"Personel ödemesi iptal edildi: {reason}", before, after)
        schedule_excel_sync(); flash("Ödeme ve bağlı ELEMAN gideri iptal edildi.", "success")
    except Exception as e: flash(str(e), "danger")
    return redirect(url_for("admin_personnel", month=month))


# ----------------------- Reports / export -----------------------

@app.route("/admin/reports")
@admin_required
def reports():
    start=request.args.get("start") or default_range()[0]; end=request.args.get("end") or default_range()[1]
    try: start,end=parse_day(start),parse_day(end)
    except: start,end=default_range()
    p=profit_loss(start,end); monthly=monthly_profit_rows(start,end); return render_template("reports.html",start=start,end=end,p=p,monthly=monthly)


def report_workbook(start,end):
    wb=Workbook(); wb.remove(wb.active)
    def header(ws, values): ws.append(values); style_ws(ws)
    ws=wb.create_sheet("KASA"); header(ws,["ID","Tarih","Saat","Personel/Yönetici","PanCafe","Nakit","POS","IBAN","Hatalı","İndirim","Yeni Borç","Eski Borç Ödeme","Hesaplanan","Fark","Durum","Açıklama"])
    for r in cash_rows(start,end): ws.append([r["id"],r["tx_date"],r["time_text"],r["actor_name"],r["pancafe"],r["nakit"],r["pos"],r["iban"],r["hatali"],r["indirim"],r["new_debt"],r["old_debt_paid"],r["calc"],r["diff"],r["status"],r["note"]])
    ws=wb.create_sheet("GIDERLER"); header(ws,["ID","Tarih","Saat","Personel/Yönetici","Kategori","Tutar","Ödeme Kaynağı","Ödeme Yöntemi","Kart Sahibi","Tür","Açıklama"])
    for r in expense_rows(start,end): ws.append([r["id"],r["tx_date"],r["time_text"],r["actor_name"],r["category"],r["amount"],r["payment_source"],r["payment_method"],r["card_owner"],"ŞAHSİ KART TAHSİLATI" if r["system_type"]==PERSONAL_CARD_REIMBURSEMENT_TYPE else "GERÇEKLEŞEN GİDER",r["note"]])
    ws=wb.create_sheet("VERESIYE"); header(ws,["ID","Tarih","Kişi","Tür","Tutar","Yöntem","Personel/Yönetici","Açıklama"])
    for r in credit_rows(start,end): ws.append([r["id"],r["tx_date"],r["person"],r["movement_type"],r["amount"],r["payment_method"],r["actor_name"],r["note"]])
    ws=wb.create_sheet("ORTAK_HAREKETLERI"); header(ws,["ID","Tarih","Ortak","Tür","Tutar","Ödeme Yöntemi","Mahsup","Kaynak","Kategori","Alım Kaynağı","Açıklama"])
    for r in partner_movement_rows(start,end): ws.append([r["id"],r["tx_date"],r["partner_name"],r["movement_type"],r["amount"],r["payment_method"],r.get("payment_target","AUTO"),r["source_type"],r.get("claim_category",""),r.get("claim_origin",""),r["note"]])
    ws=wb.create_sheet("KAR_DAGITIMI"); header(ws,["ID","Dönem","Faaliyet Kârı","Devreden Zarar","Kart Koruma Giriş","Kart Koruma Çıkış","İşletmede Bırakılan","Dağıtılabilir Kâr","Sonraki Aya Zarar","Toplam Pay %","Kesinleştirme","Kesinleştiren","Not"])
    con_pd=db_conn()
    for r in con_pd.execute("SELECT * FROM profit_distributions WHERE active=1 AND period_end BETWEEN ? AND ? ORDER BY period_month",(start,end)):
        ws.append([r["id"],r["period_month"],r["operating_profit"],r["carry_loss_in"],r["card_protection_in"],r["card_protection_out"],r["reserve_amount"],r["distributable_profit"],r["carry_loss_out"],r["total_share"],r["created_at"],r["created_by"],r["note"]])
    con_pd.close()
    ws=wb.create_sheet("KAR_ZARAR"); header(ws,["Dönem","PanCafe Ciro","Gerçekleşen Gider","Faaliyet Sonucu","Dağıtılan Kâr","İşletmede Bırakılan","Devreden Zarar","İşletmeden Direkt Ödenen","Ortağa NAKİT Alacak Ödemesi","Genel Para Çıkışı","Ortak Para Çıkışı","Tahsilat","Net Nakit"])
    for p in monthly_profit_rows(start,end): ws.append([p["month"],p["revenue"],p["economic_expense"],p["operating_profit"],p["distributed_profit"],p["profit_reserve"],p["profit_carry_out"],p["business_paid"],p["partner_paid"],p["general_cash_out"],p["partner_withdrawals"],p["collections"],p["net_cash"]])
    ws=wb.create_sheet("PUANTAJ"); header(ws,["Tarih","Personel","Vardiya","Gerçek Giriş","Gerçek Çıkış","Çalışılan Saat","Not"]); con_staff=db_conn()
    for r in con_staff.execute("SELECT a.*,p.name personnel_name FROM personnel_attendance a JOIN personnel p ON p.id=a.personnel_id WHERE a.active=1 AND a.work_date BETWEEN ? AND ? ORDER BY a.work_date,a.id",(start,end)):
        ws.append([r["work_date"],r["personnel_name"],SHIFT_DEFINITIONS.get(r["shift_code"],{}).get("label",r["shift_code"]),r["actual_start"],r["actual_end"],round((r["worked_minutes"] or 0)/60.0,2),r["note"]])
    ws=wb.create_sheet("PERSONEL_AVANS"); header(ws,["Tarih","Personel","Tutar","Ödeme","Açıklama"] )
    for r in con_staff.execute("SELECT a.*,p.name personnel_name FROM personnel_advances a JOIN personnel p ON p.id=a.personnel_id WHERE a.active=1 AND a.tx_date BETWEEN ? AND ? ORDER BY a.tx_date,a.id",(start,end)):
        ws.append([r["tx_date"],r["personnel_name"],r["amount"],r["payment_method"],r["note"]])
    ws=wb.create_sheet("PERSONEL_ODEME"); header(ws,["Dönem","Tarih","Personel","Tür","Tutar","Yöntem","Açıklama"] )
    for r in con_staff.execute("SELECT x.*,p.name personnel_name FROM personnel_payments x JOIN personnel p ON p.id=x.personnel_id WHERE x.active=1 AND x.tx_date BETWEEN ? AND ? ORDER BY x.tx_date,x.id",(start,end)):
        ws.append([r["period_month"],r["tx_date"],r["personnel_name"],r["payment_kind"],r["amount"],r["payment_method"],r["note"]])
    con_staff.close()
    ws=wb.create_sheet("AUDIT"); header(ws,["Zaman","Kullanıcı","Kaynak","Varlık","ID","İşlem","Özet"]); con=db_conn()
    for r in con.execute("SELECT * FROM audit_log WHERE substr(created_at,1,10) BETWEEN ? AND ? ORDER BY id",(start,end)): ws.append([r["created_at"],r["actor_name"],r["actor_type"],r["entity_type"],r["entity_id"],r["action"],r["summary"]])
    con.close(); return wb

@app.get("/admin/export")
@admin_required
def export_page():
    return redirect(url_for("reports",start=request.args.get("start") or default_range()[0],end=request.args.get("end") or default_range()[1]))

@app.post("/admin/export")
@admin_required
def export_download():
    try:
        start=parse_day(request.form.get("start")); end=parse_day(request.form.get("end")); wb=report_workbook(start,end); literal_workbook_text(wb); bio=BytesIO(); wb.save(bio); bio.seek(0); return send_file(bio,as_attachment=True,download_name=f"KampusArena_Rapor_{start}_{end}.xlsx",mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e: flash(str(e),"danger"); return redirect(url_for("reports"))


# ----------------------- Audit / logs -----------------------

@app.route("/admin/audit")
@admin_required
def audit_page():
    entity=request.args.get("entity","TÜMÜ"); action=request.args.get("action","TÜMÜ"); con=db_conn(); args=[]; sql="SELECT * FROM audit_log WHERE 1=1"
    if entity!="TÜMÜ": sql+=" AND entity_type=?"; args.append(entity)
    if action!="TÜMÜ": sql+=" AND action=?"; args.append(action)
    sql+=" ORDER BY id DESC LIMIT 500"; rows=[dict(r) for r in con.execute(sql,args).fetchall()]; entities=[r[0] for r in con.execute("SELECT DISTINCT entity_type FROM audit_log ORDER BY entity_type").fetchall()]; actions=[r[0] for r in con.execute("SELECT DISTINCT action FROM audit_log ORDER BY action").fetchall()]; con.close(); return render_template("audit.html",rows=rows,entity=entity,action=action,entities=entities,actions=actions)

@app.route("/admin/logs")
@admin_required
def logs_page():
    con=db_conn(); rows=[dict(r) for r in con.execute("SELECT * FROM app_log ORDER BY id DESC LIMIT 300").fetchall()]; con.close(); return render_template("logs.html",logs=rows)

@app.route("/api/logs")
@admin_required
def logs_api():
    after=int(request.args.get("after",0) or 0); con=db_conn(); rows=[dict(r) for r in con.execute("SELECT * FROM app_log WHERE id>? ORDER BY id ASC LIMIT 200",(after,)).fetchall()]; con.close(); return jsonify(rows)


# ----------------------- Settings -----------------------

@app.route("/admin/settings",methods=["GET","POST"])
@admin_required
def settings():
    if request.method=="POST":
        # Excel password change, owner only
        if not session.get("admin_owner"):
            flash("Excel ana şifresini yalnız ana yönetici değiştirebilir.","danger")
        else:
            try:
                old=request.form.get("old_password") or ""; new=request.form.get("new_password") or ""; confirm=request.form.get("confirm_password") or ""
                if old!=load_excel_password(): raise RuntimeError("Mevcut Excel şifresi yanlış.")
                if len(new)<4 or len(new)>15 or new!=confirm: raise RuntimeError("Yeni şifre 4-15 karakter olmalı ve tekrar ile aynı olmalı.")
                store_excel_password(new); sync_excel_now(True,password=new); add_audit("SYSTEM",None,"EXCEL_PASSWORD_CHANGE","Excel ana şifresi değiştirildi."); flash("Excel ana şifresi değiştirildi.","success")
            except Exception as e: flash(str(e),"danger")
    return render_template("settings.html",backup_info=scheduled_backup_info(),personnel=personnel_rows(True),shifts=SHIFT_DEFINITIONS,admin_accounts=get_admin_accounts(),current_admin_username=session.get("admin_username"),current_admin_owner=bool(session.get("admin_owner")))

@app.post("/admin/settings/personnel/add")
@admin_required
def personnel_add():
    try:
        name=request.form.get("name","").strip(); salary=money(request.form.get("monthly_salary")); default_shift=(request.form.get("default_shift") or "").strip()
        if not name: raise RuntimeError("Personel adı boş olamaz.")
        if salary < 0: raise RuntimeError("Maaş negatif olamaz.")
        if default_shift and default_shift not in SHIFT_DEFINITIONS: raise RuntimeError("Varsayılan vardiya geçersiz.")
        con=db_conn(); row=con.execute("SELECT * FROM personnel WHERE name=? COLLATE NOCASE",(name,)).fetchone()
        if row:
            con.execute("UPDATE personnel SET active=1,monthly_salary=?,default_shift=? WHERE id=?",(salary,default_shift,row["id"])); pid=row["id"]
        else:
            cur=con.execute("INSERT INTO personnel(name,monthly_salary,default_shift,active,created_at) VALUES(?,?,?,1,?)",(name,salary,default_shift,now_iso())); pid=cur.lastrowid
        record_salary_rate(con, pid, salary)
        con.commit(); con.close(); add_audit("PERSONNEL",pid,"CREATE","Personel eklendi/aktif edildi; maaş ve vardiya bilgisi kaydedildi."); schedule_excel_sync(); flash("Personel eklendi.","success")
    except Exception as e: flash(str(e),"danger")
    return redirect(url_for("settings"))

@app.post("/admin/settings/personnel/update/<int:rid>")
@admin_required
def personnel_update(rid):
    try:
        salary=money(request.form.get("monthly_salary")); default_shift=(request.form.get("default_shift") or "").strip()
        if salary < 0: raise RuntimeError("Maaş negatif olamaz.")
        if default_shift and default_shift not in SHIFT_DEFINITIONS: raise RuntimeError("Varsayılan vardiya geçersiz.")
        con=db_conn(); before=rowdict(con.execute("SELECT * FROM personnel WHERE id=?",(rid,)).fetchone())
        if not before: con.close(); raise RuntimeError("Personel bulunamadı.")
        con.execute("UPDATE personnel SET monthly_salary=?,default_shift=? WHERE id=?",(salary,default_shift,rid)); record_salary_rate(con, rid, salary); con.commit(); after=rowdict(con.execute("SELECT * FROM personnel WHERE id=?",(rid,)).fetchone()); con.close()
        add_audit("PERSONNEL",rid,"UPDATE","Personel maaşı / varsayılan vardiyası güncellendi.",before,after); schedule_excel_sync(); flash("Personel bilgisi güncellendi.","success")
    except Exception as e: flash(str(e),"danger")
    return redirect(url_for("settings"))

@app.post("/admin/settings/personnel/delete/<int:rid>")
@admin_required
def personnel_remove(rid):
    con=db_conn(); con.execute("UPDATE personnel SET active=0 WHERE id=?",(rid,)); con.commit(); con.close(); add_audit("PERSONNEL",rid,"DEACTIVATE","Personel aktif listeden çıkarıldı."); schedule_excel_sync(); flash("Personel aktif listeden çıkarıldı; eski kayıtları korunur.","success"); return redirect(url_for("settings"))

@app.post("/admin/settings/admin/add")
@owner_required
def admin_account_add():
    try:
        password=request.form.get("password") or ""; confirm=request.form.get("password_confirm") or ""
        if password!=confirm: raise RuntimeError("Şifreler aynı değil.")
        aid=create_admin(request.form.get("username"),password,False,session.get("admin_username","")); add_audit("ADMIN",aid,"CREATE","Yeni yönetici hesabı eklendi."); flash("Yönetici hesabı eklendi.","success")
    except Exception as e: flash(str(e),"danger")
    return redirect(url_for("settings"))

@app.post("/admin/settings/admin/delete/<int:rid>")
@owner_required
def admin_account_remove(rid):
    try:
        if rid==session.get("admin_id"): raise RuntimeError("Kendi hesabınızı devre dışı bırakamazsınız.")
        con=db_conn(); row=con.execute("SELECT * FROM admin_accounts WHERE id=?",(rid,)).fetchone()
        if not row: con.close(); raise RuntimeError("Hesap bulunamadı.")
        if row["is_owner"]: con.close(); raise RuntimeError("Ana yönetici devre dışı bırakılamaz.")
        con.execute("UPDATE admin_accounts SET active=0 WHERE id=?",(rid,)); con.commit(); con.close(); add_audit("ADMIN",rid,"DEACTIVATE","Yönetici hesabı devre dışı bırakıldı."); flash("Yönetici hesabı devre dışı bırakıldı.","success")
    except Exception as e: flash(str(e),"danger")
    return redirect(url_for("settings"))

@app.post("/admin/settings/admin/password")
@admin_required
def admin_password_change():
    try:
        acc=authenticate_admin(session.get("admin_username"),request.form.get("current_password") or "")
        if not acc: raise RuntimeError("Mevcut yönetici şifresi yanlış.")
        new=request.form.get("new_password") or ""; confirm=request.form.get("confirm_password") or ""
        if len(new)<8 or new!=confirm: raise RuntimeError("Yeni şifre en az 8 karakter ve tekrar ile aynı olmalı.")
        salt,digest=password_digest(new); con=db_conn(); con.execute("UPDATE admin_accounts SET password_salt=?,password_hash=? WHERE id=?",(salt,digest,session["admin_id"])); con.commit(); con.close(); add_audit("ADMIN",session["admin_id"],"PASSWORD_CHANGE","Yönetici panel şifresi değiştirildi."); flash("Yönetici şifreniz değiştirildi.","success")
    except Exception as e: flash(str(e),"danger")
    return redirect(url_for("settings"))


def pancafe_background_worker():
    """Maintenance runs independently from the web UI.

    It may copy/read the large PanCafe DB, but admin/personnel page requests
    never wait on this worker.
    """
    while True:
        started = time.perf_counter()
        try:
            auto_close_due_days()
        except Exception as e:
            try:
                add_log("ERROR","PANCAFE_ARKAPLAN",str(e),source="PANCAFE")
            except Exception:
                pass
        elapsed = time.perf_counter() - started
        time.sleep(max(5, 60 - elapsed))


# ----------------------- Excel download / health -----------------------

@app.route("/download-excel")
@admin_required
def download_excel():
    try:
        sync_excel_now(True)
        return send_file(EXCEL_FILE,as_attachment=True,download_name="KampusArena_Muhasebe.xlsx")
    except Exception as e: flash(str(e),"danger"); return redirect(url_for("admin"))



def another_accounting_server_running():
    """Prevent a second app.py process from touching the same SQLite DB."""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(0.35)
            return s.connect_ex(("127.0.0.1", 8788)) == 0
    except Exception:
        return False


if __name__ == "__main__" and another_accounting_server_running():
    print("Kampus Arena Muhasebe zaten calisiyor: http://127.0.0.1:8788")
    sys.exit(0)

init_db()
migrate_v105_personal_card_logic()

# V10.6.3: do not block server startup with a PanCafe snapshot/day-close pass.
# pancafe_background_worker starts immediately below and performs maintenance
# independently from page requests.
installation_self_check()

if __name__ == "__main__":
    add_log("INFO","SERVER_START","Kampüs Arena Muhasebe V10.6.3 başladı",source="SYSTEM")
    threading.Thread(target=pancafe_background_worker,daemon=True,name="PanCafeAuto").start()
    threading.Thread(target=scheduled_backup_worker,daemon=True,name="ScheduledBackups").start()
    app.run(host="0.0.0.0",port=8788,debug=False,threaded=True)
