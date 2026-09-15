"""Build a curated, verifiable project ZIP without moving original materials."""
import hashlib
import json
import shutil
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'deliverables/20260916_resume_portfolio'
PREFIX = 'PPT模板替换_简历作品集'
ZIP = OUT / f'{PREFIX}.zip'

def digest(path):
    with path.open('rb') as f:
        return hashlib.file_digest(f, 'sha256').hexdigest()

files = {}
for path in ROOT.iterdir():
    if path.is_file() and (path.suffix in {'.py', '.ps1', '.bat', '.spec', '.md'} or path.name in {'requirements.txt', '.gitignore'}):
        files[path.name] = path
for folder in ['prototype_tool', 'tests', 'scripts', 'docs']:
    for path in (ROOT / folder).rglob('*'):
        if path.is_file() and '__pycache__' not in path.parts:
            files[path.relative_to(ROOT).as_posix()] = path
for name in ['EverMemOS_Introduction.pptx', '手把手带你玩转腾讯版自研小龙虾 WorkBuddy.pptx', 'OpenClaw TVP研讨会-PPT模板.pptx', '上海-架构师城市沙龙-PPT模板.pptx']:
    files[name] = ROOT / name
exe = OUT / 'app/PPTTemplateReplacer.exe'
if not exe.is_file():
    raise RuntimeError('Current-source executable is missing')
for folder in ['app', 'screenshots', 'evidence', 'demo']:
    for path in (OUT / folder).rglob('*'):
        if path.is_file():
            files[path.relative_to(OUT).as_posix()] = path
for name in ['交付说明.md', '截图说明.md']:
    files[name] = OUT / name
shutil.copy2(ROOT / 'docs/简历项目描述.md', OUT / '简历项目描述.md')
manifest = [{'path': name, 'bytes': p.stat().st_size, 'sha256': digest(p)} for name, p in sorted(files.items())]
manifest_path = OUT / '文件清单.json'
manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
with zipfile.ZipFile(ZIP, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as z:
    for name, p in sorted(files.items()):
        z.write(p, f'{PREFIX}/{name}')
    z.write(manifest_path, f'{PREFIX}/文件清单.json')
with zipfile.ZipFile(ZIP) as z:
    assert z.testzip() is None, 'ZIP CRC validation failed'
    assert len(z.namelist()) == len(files) + 1
(OUT / 'SHA256.txt').write_text(f'{digest(ZIP)}  {ZIP.name}\n', encoding='utf-8')
print(json.dumps({'zip': str(ZIP), 'files': len(files) + 1, 'bytes': ZIP.stat().st_size, 'crc': 'OK'}, ensure_ascii=False))
