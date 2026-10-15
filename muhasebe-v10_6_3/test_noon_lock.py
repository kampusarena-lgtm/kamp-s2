import importlib.util, pathlib, tempfile, shutil, sqlite3, unittest, sys
from datetime import datetime as RealDateTime, timedelta
ROOT=pathlib.Path(__file__).resolve().parent
TEMP=pathlib.Path(tempfile.mkdtemp(prefix="noon-lock-tests-"))
shutil.copytree(ROOT/"app",TEMP/"app",dirs_exist_ok=True)
spec=importlib.util.spec_from_file_location("tested_accounting_app",TEMP/"app"/"app.py")
a=importlib.util.module_from_spec(spec);sys.modules[spec.name]=a;spec.loader.exec_module(a)
class Frozen(RealDateTime):
    value=RealDateTime(2026,10,16,10)
    @classmethod
    def now(cls,tz=None):
        return cls.value.replace(tzinfo=tz) if tz else cls.value
a.datetime=Frozen
a.schedule_excel_sync=lambda:None
a.app.config["TESTING"]=True
class NoonLockTests(unittest.TestCase):
    def setUp(self):
        Frozen.value=RealDateTime(2026,10,16,10)
        a.DB_FILE=TEMP/(self._testMethodName+".db")
        a.BACKUP_DIR=TEMP/(self._testMethodName+"_backups")
        a.BACKUP_DIR.mkdir()
        a.init_db();a.migrate_v105_personal_card_logic()
        self.client=a.app.test_client()
    def time(self,h,m=0,s=0):
        Frozen.value=RealDateTime(2026,10,16,h,m,s)
    def query(self,sql,args=()):
        con=a.db_conn()
        try:return [dict(r) for r in con.execute(sql,args).fetchall()]
        finally:con.close()
    def cash(self,day="2026-10-15",pancafe=375):
        con=a.db_conn()
        cur=con.execute("INSERT INTO cash_entries(tx_date,created_at,pancafe,nakit,pos,note) VALUES(?,?,?,?,?,?)",(day,a.now_iso(),pancafe,150,225,"preserve me"))
        rid=cur.lastrowid;con.commit();con.close();return rid
    def close(self,day):
        con=a.db_conn();con.execute("INSERT INTO day_closures(tx_date,is_closed,closed_at,closed_by,close_note) VALUES(?,1,'ORIGINAL_TIME','ORIGINAL_OWNER','previously handled')",(day,));con.commit();con.close()
    def login(self):
        self.client.get("/admin-login")
        with self.client.session_transaction() as s:token=s["csrf_token"]
        self.client.post("/admin-login",data={"csrf_token":token,"username":"Owner","password":"test-password","password_confirm":"test-password"})
        self.client.get("/admin")
    def post(self,path,data):
        with self.client.session_transaction() as s:token=s["csrf_token"]
        return self.client.post(path,data={"csrf_token":token,**data})
    def test_0859_0900_1159_1200_boundaries(self):
        for h,m,s,allowed in [(8,59,59,True),(9,0,0,True),(11,59,59,True),(12,0,0,False),(13,0,0,False)]:
            with self.subTest(time=(h,m,s)):
                self.time(h,m,s)
                if allowed:a.ensure_day_open("2026-10-15")
                else:self.assertRaises(RuntimeError,a.ensure_day_open,"2026-10-15")
        self.time(8,59);self.assertRaises(RuntimeError,a.ensure_day_open,"2026-10-14")
    def test_panCafe_bounds_stay_0900(self):
        start,end=a.accounting_bounds("2026-10-15")
        self.assertEqual(start,RealDateTime(2026,10,15,9))
        self.assertEqual(end,RealDateTime(2026,10,16,9))
        self.time(8,59);self.assertEqual(a.accounting_day_iso(),"2026-10-15")
        self.time(9);self.assertEqual(a.accounting_day_iso(),"2026-10-16")
    def test_worker_skips_noon_grace_and_closes_at_noon(self):
        self.cash();before=self.query("select * from cash_entries")
        self.time(9);a.auto_close_due_days()
        self.assertEqual(self.query("select * from day_closures"),[])
        self.assertEqual(before,self.query("select * from cash_entries"))
        self.time(11,59,59);a.auto_close_due_days()
        self.assertEqual(self.query("select * from day_closures"),[])
        self.time(12);a.auto_close_due_days()
        self.assertEqual(self.query("select is_closed from day_closures")[0]["is_closed"],1)
        self.assertEqual(before,self.query("select * from cash_entries"))
    def test_missing_panCafe_does_not_block_entry_before_noon(self):
        self.cash(pancafe=None);original=a.auto_fill_pancafe_for_day;calls=[]
        a.auto_fill_pancafe_for_day=lambda day,force=False:calls.append(day)
        try:
            self.time(9);a.auto_close_due_days()
            self.assertIn("2026-10-15",calls)
            a.ensure_day_open("2026-10-15")
            status=a.day_status("2026-10-15")
            self.assertTrue(status["entry_grace"]);self.assertFalse(status["effective_locked"])
            self.assertEqual(status["waiting"],1)
            self.time(12);a.auto_close_due_days()
            self.assertRaises(RuntimeError,a.ensure_day_open,"2026-10-15")
            self.assertTrue(a.day_status("2026-10-15")["effective_locked"])
        finally:a.auto_fill_pancafe_for_day=original
    def test_previous_handled_locks_not_reopened(self):
        self.cash("2026-10-14");self.close("2026-10-14");self.close("2026-10-15")
        before=self.query("select * from day_closures order by tx_date")
        self.time(10);a.auto_close_due_days()
        self.assertEqual(before,self.query("select * from day_closures order by tx_date"))
        for day in ("2026-10-14","2026-10-15"):
            self.assertRaises(RuntimeError,a.ensure_day_open,day)
        self.assertNotIn('value="2026-10-15" max=',self.client.get("/").text.split('id="cash-date"')[1].split(">")[0])
    def test_default_cash_date_and_other_forms(self):
        self.time(10);text=self.client.get("/").text
        self.assertIn('id="cash-date" type="date" name="tarih" value="2026-10-15"',text)
        self.assertIn('id="expense-date" type="date" name="tarih" value="2026-10-16"',text)
        self.assertIn('id="credit-date" type="date" name="tarih" value="2026-10-16"',text)
        self.time(12);text=self.client.get("/").text
        self.assertIn('id="cash-date" type="date" name="tarih" value="2026-10-16"',text)
    def test_late_cash_save_and_edit_before_noon_then_reject(self):
        self.login();self.time(10)
        self.post("/save-cash",{"tarih":"2026-10-15","nakit":"150","pos":"225"})
        row=self.query("select * from cash_entries")[0];self.assertEqual(row["tx_date"],"2026-10-15")
        self.post("/admin/cash/edit/"+str(row["id"]),{"tarih":"2026-10-15","nakit":"175","pos":"225","version":str(row["version"])})
        row=self.query("select * from cash_entries")[0];self.assertEqual(row["nakit"],175)
        self.time(12)
        self.post("/admin/cash/edit/"+str(row["id"]),{"tarih":"2026-10-15","nakit":"200","pos":"225","version":str(row["version"])})
        self.assertEqual(row,self.query("select * from cash_entries")[0])
    def test_expense_and_credit_allowed_before_noon_only(self):
        self.login();self.time(10)
        self.post("/save-expense",{"tarih":"2026-10-15","kategori":"MARKET","tutar":"25"})
        self.post("/save-credit",{"tarih":"2026-10-15","kisi":"Test","tur":"BORÇ","tutar":"10"})
        self.assertEqual(len(self.query("select * from expenses")),1)
        self.assertEqual(len(self.query("select * from credit_entries")),1)
        self.time(12)
        self.post("/save-expense",{"tarih":"2026-10-15","kategori":"MARKET","tutar":"30"})
        self.post("/save-credit",{"tarih":"2026-10-15","kisi":"Test","tur":"BORÇ","tutar":"20"})
        self.assertEqual(len(self.query("select * from expenses")),1)
        self.assertEqual(len(self.query("select * from credit_entries")),1)
    def test_month_year_leap_day_and_future(self):
        for value,day in [(RealDateTime(2026,11,1,10),"2026-10-31"),(RealDateTime(2027,1,1,10),"2026-12-31"),(RealDateTime(2028,3,1,10),"2028-02-29")]:
            Frozen.value=value;a.ensure_day_open(day)
            self.assertEqual(a.cash_entry_default_date(),day)
            self.assertRaises(RuntimeError,a.ensure_day_open,(value.date()+timedelta(days=1)).isoformat())
    def test_explicit_owner_reopen_remains_available(self):
        self.login()
        with a.app.test_request_context():
            from flask import session
            session.update(admin=True,admin_username="Owner")
            a.reopen_days("2026-10-13","2026-10-13","Manual historical correction")
        self.time(12);a.ensure_day_open("2026-10-13")
    def test_templates_and_status_consistent(self):
        self.login();self.time(10)
        for template in (TEMP/"app"/"templates").glob("*.html"):a.app.jinja_env.parse(template.read_text(encoding="utf-8-sig"))
        text=self.client.get("/admin/day-close?date=2026-10-15").text
        self.assertIn("12:00’YE KADAR KAYIT AÇIK",text)
        self.assertIn("Kayıt girişi 12:00’ye kadar zaten açık",text)
        self.assertFalse(a.day_status("2026-10-15")["effective_locked"])
        self.time(12);self.assertTrue(a.day_status("2026-10-15")["effective_locked"])
    def test_import_has_no_new_schema_or_historical_migration(self):
        self.cash("2026-10-01");self.close("2026-10-01")
        before=self.query("select * from cash_entries");locks=self.query("select * from day_closures")
        a.init_db();a.migrate_v105_personal_card_logic()
        self.assertEqual(before,self.query("select * from cash_entries"))
        self.assertEqual(locks,self.query("select * from day_closures"))
if __name__=="__main__":unittest.main(verbosity=2)
