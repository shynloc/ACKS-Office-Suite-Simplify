#!/usr/bin/env python3
"""打包技能：skills/acks-office + src/acks_office → dist/acks-office-skill-<版本>.zip

压缩包的根目录是 acks-office/，解压到任意 Agent 的技能目录即可使用。程序代码放在 lib/ 里，
入口脚本优先使用它，所以不必再单独安装 acks-office；依赖由 doctor 给出安装步骤。

用法：python tools/build_skill_bundle.py [--out dist]
"""

import argparse
import re
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SKILL = ROOT / "skills" / "acks-office"
PACKAGE = ROOT / "src" / "acks_office"
SKIP_DIRS = {"__pycache__", "tests", "lib"}
SKIP_NAMES = {".DS_Store"}


def _version(path: Path, pattern: str) -> str:
    match = re.search(pattern, path.read_text(encoding="utf-8"), re.M)
    if not match:
        raise SystemExit(f"在 {path} 中找不到版本号")
    return match.group(1)


def _files(base: Path):
    for path in sorted(base.rglob("*")):
        rel = path.relative_to(base)
        if (path.is_file() and not SKIP_DIRS & set(rel.parts[:-1])
                and path.name not in SKIP_NAMES and path.suffix not in (".pyc", ".pyo")):
            yield path, rel.as_posix()


def build(out_dir: Path) -> Path:
    version = _version(PACKAGE / "__init__.py", r'^__version__ = "([^"]+)"')
    for path, pattern in ((SKILL / "SKILL.md", r'^  version: "([^"]+)"'),
                          (SKILL / "scripts" / "acks.py", r'^VERSION = "([^"]+)"')):
        found = _version(path, pattern)
        if found != version:
            raise SystemExit(f"版本号不一致：{path.relative_to(ROOT)} 是 {found}，程序是 {version}")

    out_dir.mkdir(parents=True, exist_ok=True)
    target = out_dir / f"acks-office-skill-{version}.zip"
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as bundle:
        for path, rel in _files(SKILL):
            bundle.write(path, f"acks-office/{rel}")
        bundle.write(ROOT / "LICENSE", "acks-office/LICENSE")
        for path, rel in _files(PACKAGE):
            bundle.write(path, f"acks-office/lib/acks_office/{rel}")
    return target


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--out", default=str(ROOT / "dist"), help="输出目录，默认 dist/")
    target = build(Path(parser.parse_args().out))
    print(target)
    return 0


if __name__ == "__main__":
    sys.exit(main())
