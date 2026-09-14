"""
Standalone converter unit tests — no pytest, no pytest import.
Run directly: PYTHONPATH=. python3 backend/tests/test_converters.py

Tests the core pure-logic helpers in the converter modules.
"""
import sys
import os
import re
import tempfile
from pathlib import Path

failures = 0

def ok(cond, msg):
    global failures
    if cond:
        print(f"  \u2713 {msg}")
    else:
        print(f"  \u2717 FAIL: {msg}")
        failures += 1


# ─── 1. PlantUML config ─────────────────────────────────────────────────────

print("\n=== PlantUML config helpers ===")

from backend.converters.plantuml_converter import PlantUMLConfig, PlantUMLConverter

cfg = PlantUMLConfig(dpi=300, format='png', timeout=60)
ok(cfg.dpi == 300, "PlantUMLConfig dpi default")
ok(cfg.format == 'png', "PlantUMLConfig format default")
ok(cfg.timeout == 60, "PlantUMLConfig timeout default")

ok('.puml' in PlantUMLConverter.SUPPORTED_EXTENSIONS, "SUPPORTED_EXTENSIONS includes .puml")
ok('.plantuml' in PlantUMLConverter.SUPPORTED_EXTENSIONS, "SUPPORTED_EXTENSIONS includes .plantuml")
ok('.pu' in PlantUMLConverter.SUPPORTED_EXTENSIONS, "SUPPORTED_EXTENSIONS includes .pu")


# ─── 2. PlantUML Chinese-font preprocessing ──────────────────────────────────

print("\n=== PlantUML preprocessing (Chinese font) ===")

# Instantiate without __init__ side-effects by creating a minimal subclass
class FakeLogger:
    def info(self, m): pass
    def warning(self, m): pass
    def debug(self, m): pass
    def error(self, m): pass

class TestablePlantUMLConverter(PlantUMLConverter):
    def __init__(self):
        # skip BaseConverter.__init__ — no filesystem deps
        self.logger = FakeLogger()
        self.plantuml_config = PlantUMLConfig()
        self._dependency_status = None

conv = TestablePlantUMLConverter()

# Test _build_preprocessed_content (instance method)
content_chinese = "@startuml\nAlice -> Bob: 你好\n@enduml"
content_english = "@startuml\nAlice -> Bob: Hello\n@enduml"

built_cn = conv._build_preprocessed_content(content_chinese)
ok('Microsoft YaHei' in built_cn, "Chinese content gets font config")

built_theme = conv._build_preprocessed_content("!theme plain\n@startuml\nBob: 中文\n@enduml")
ok('skinparam defaultFontName "Microsoft YaHei"' in built_theme, "Font config injected after !theme plain")

# _build_preprocessed_content is called by _preprocess_plantuml_file ONLY when Chinese is detected.
# We test the full _preprocess_plantuml_file path for English (no injection should happen).
with tempfile.NamedTemporaryFile(mode='w', suffix='.puml', delete=False, encoding='utf-8') as f:
    f.write("@startuml\nAlice -> Bob: Hello\n@enduml\n")
    english_puml = f.name

result_en = conv._preprocess_plantuml_file(english_puml)
# Should return the original path unchanged (no temp file created)
ok(result_en == english_puml, "English-only file: _preprocess_plantuml_file returns original path unchanged")
os.unlink(english_puml)


# ─── 3. PlantUML error parsing ───────────────────────────────────────────────

print("\n=== PlantUML error parsing ===")

patterns = [
    ('Syntax error at line 5', '语法错误'),
    ('Cannot find Graphviz', 'Graphviz未安装或未找到'),
    ('OutOfMemoryError', '内存不足'),
    ('FileNotFoundException', '文件未找到'),
]

for err, expected in patterns:
    parsed = conv._parse_plantuml_error(err)
    ok(expected in parsed, f"Error parse: '{err[:20]}' -> contains '{expected}'")


# ─── 4. Batik config ─────────────────────────────────────────────────────────

print("\n=== Batik config ===")

from backend.converters.batik_converter import BatikConfig, BatikConverter

bcfg = BatikConfig(dpi=300, quality=1.0, timeout=60)
ok(bcfg.dpi == 300, "BatikConfig dpi default")
ok(bcfg.quality == 1.0, "BatikConfig quality default")
ok(bcfg.timeout == 60, "BatikConfig timeout default")

ok('.svg' in BatikConverter.SUPPORTED_EXTENSIONS, "SUPPORTED_EXTENSIONS includes .svg")


# ─── 5. Batik SVG preprocessing (SVG 2.0 syntax fixes) ──────────────────────

print("\n=== Batik SVG preprocessing ===")

class TestableBatikConverter(BatikConverter):
    def __init__(self):
        self.logger = FakeLogger()
        self.batik_config = BatikConfig()
        self._dependency_status = type('S', (), {'batik_jar_path': None, 'batik_lib_path': None, 'is_ready': False})()

batik = TestableBatikConverter()

# SVG with orient="auto-start-reverse" → should be fixed
with tempfile.NamedTemporaryFile(mode='w', suffix='.svg', delete=False, encoding='utf-8') as f:
    f.write('<svg xmlns="http://www.w3.org/2000/svg"><marker orient="auto-start-reverse">test</marker></svg>')
    test_svg = f.name

result = batik._preprocess_svg_for_batik(test_svg)
ok(result != test_svg, "SVG with auto-start-reverse creates temp file")
with open(result, 'r') as f:
    content = f.read()
ok('orient="auto"' in content and 'auto-start-reverse' not in content,
   "auto-start-reverse fixed to auto")
os.unlink(test_svg)
if result != test_svg:
    os.unlink(result)

# Clean SVG → unchanged
with tempfile.NamedTemporaryFile(mode='w', suffix='.svg', delete=False, encoding='utf-8') as f:
    f.write('<svg xmlns="http://www.w3.org/2000/svg"><rect/></svg>')
    clean_svg = f.name

result2 = batik._preprocess_svg_for_batik(clean_svg)
ok(result2 == clean_svg, "Clean SVG returned unchanged")
os.unlink(clean_svg)


# ─── 6. DiagramToPng file-type detection ─────────────────────────────────────

print("\n=== DiagramToPng file-type detection ===")

from backend.converters.diagram_to_png import DiagramToPngConverter

d2p = DiagramToPngConverter.__new__(DiagramToPngConverter)
d2p.logger = FakeLogger()
d2p.tools_status = {}

for ext, expected_type in [
    ('.svg', 'svg'),
    ('.drawio', 'drawio'),
    ('.mmd', 'mermaid'),
    ('.puml', 'plantuml'),
    ('.plantuml', 'plantuml'),
    ('.pu', 'plantuml'),
    ('.jpg', None),
    ('.png', None),
]:
    p = Path(f'/tmp/test{ext}')
    detected = d2p._get_file_type(p)
    ok(detected == expected_type, f"_get_file_type({ext}) = {expected_type}")


# ─── 7. dep_check lib_available ───────────────────────────────────────────────

print("\n=== dep_check lib_available ===")

from backend.converters.dep_check import lib_available

ok(isinstance(lib_available('os'), bool), "lib_available('os') returns bool")
ok(isinstance(lib_available('sys'), bool), "lib_available('sys') returns bool")

# These may or may not be installed — just check it doesn't crash
for lib in ['PIL', 'Pillow', 'fitz', 'PyMuPDF', 'pdf2image', 'reportlab', 'svglib']:
    try:
        result = lib_available(lib)
        ok(isinstance(result, bool), f"lib_available('{lib}') returns bool")
    except Exception as e:
        ok(False, f"lib_available('{lib}') raised: {e}")


# ─── 8. dep_check install_hint_for ───────────────────────────────────────────

print("\n=== dep_check install_hint_for ===")

from backend.converters.dep_check import install_hint_for
import backend.converters.dep_check as dc

# Capture real platform
_real_platform = dc.sys.platform

for cmd in ['pandoc', 'tesseract', 'java', 'mmdc', 'drawio', 'graphviz', 'soffice']:
    hint = install_hint_for(cmd)
    ok(isinstance(hint, str) and len(hint) > 3,
       f"install_hint_for('{cmd}') returns non-empty str")

for plat, name in [('win32', 'Windows'), ('darwin', 'macOS'), ('linux', 'Linux')]:
    dc.sys.platform = plat
    hint = install_hint_for('graphviz')
    ok(len(hint) > 0, f"graphviz hint on {name} is non-empty")
    hint2 = install_hint_for('soffice')
    ok(len(hint2) > 0, f"soffice hint on {name} is non-empty")

dc.sys.platform = _real_platform  # restore


# ─── Summary ─────────────────────────────────────────────────────────────────

print(f"\n{'='*50}")
if failures == 0:
    print("ALL TESTS PASSED")
else:
    print(f"TESTS FAILED: {failures}")
sys.exit(failures)
