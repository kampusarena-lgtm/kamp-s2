from pathlib import Path
import zipfile,re
root=Path(__file__).resolve().parents[1]
archive=root/"KampusArena_V10_6_3_VERI_KORUYAN_GUNCELLEME.zip"
prefix="KampusArena_V10_6_3_Guncelleme/"
files=["app.py","templates/base.html","templates/day_close.html","templates/admin.html","templates/entry.html","templates/settings.html","static/admin-idle.js"]
with zipfile.ZipFile(archive) as z:
    assert z.testzip() is None
    assert len(z.namelist())==10
    assert not any(p.endswith((".db",".dat",".xlsx",".key")) for p in z.namelist())
    for name in files:
        assert z.read(prefix+name)==(root/"muhasebe-v10_6_3"/"app"/name).read_bytes(), name
    ps=z.read(prefix+"GUNCELLE.ps1").decode("utf-8-sig")
    selected=re.search(r"\$files = @\(([^\n]+)\)",ps).group(1)
    actual=[x.replace(chr(92),"/") for x in re.findall(r"'([^']+)'",selected)]
    assert set(actual)==set(files),actual
print("Release ZIP CRC, tested-source equality, data-file exclusion and update whitelist passed")
