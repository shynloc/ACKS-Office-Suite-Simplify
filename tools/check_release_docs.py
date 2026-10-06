#!/usr/bin/env python3
"""Check versioned public documentation before building or publishing."""

import argparse
import ast
import re
from pathlib import Path
from urllib.parse import unquote, urlsplit

ROOT = Path(__file__).resolve().parent.parent
REPOSITORY = "/shynloc/ACKS-Office-Suite-Simplify/"


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--tag", help="Expected publication tag")
    parser.add_argument("--ref-type", help="GitHub ref type for production publishing")
    args = parser.parse_args()
    errors = []
    source = (ROOT / "src/acks_office/__init__.py").read_text(encoding="utf-8")
    version = re.search(r'^__version__ = "([^"]+)"', source, re.M).group(1)
    if args.tag and args.tag != "v" + version:
        errors.append("Publication tag must match package version: v" + version)
    if args.ref_type and args.ref_type != "tag":
        errors.append("Production publishing requires a version tag.")
    config = (ROOT / "pyproject.toml").read_text(encoding="utf-8")
    if not re.search(r'^readme = "README.md"$', config, re.M):
        errors.append("PyPI description must use the shared README.md.")
    documentation = "https://github.com" + REPOSITORY + "blob/v" + version + "/README.md"
    if 'Documentation = "' + documentation + '"' not in config:
        errors.append("Package Documentation URL must match the release README.")
    for filename in ("README.md", "INSTALL_GUIDE.md", "SECURITY.md"):
        text = (ROOT / filename).read_text(encoding="utf-8")
        match = re.search(r"文档对应版本：\*\*([^*]+)\*\*", text)
        if not match or match.group(1) != version:
            errors.append(filename + ": documented version does not match " + version)
        pins = re.findall(r"acks-office(?:\[[^\]]+\])?==([0-9A-Za-z.+_-]+)", text)
        if any(pin != version for pin in pins):
            errors.append(filename + ": installation commands pin an outdated version.")
        artifacts = re.findall(r"acks-office-skill-([0-9A-Za-z.+_-]+)\.zip", text)
        artifacts += re.findall(r"acks_office-([0-9A-Za-z.+_-]+)-py3-none-any\.whl", text)
        if any(artifact != version for artifact in artifacts):
            errors.append(filename + ": download filenames use an outdated version.")
        # Both Markdown and HTML links must render correctly on GitHub and PyPI.
        links = re.findall(r'\]\(([^)]+)\)', text)
        links += re.findall(r'(?:href|src)="([^"]+)"', text)
        for link in links:
            if link.startswith("#"):
                continue
            parsed = urlsplit(link)
            if parsed.scheme not in ("https", "http", "mailto"):
                errors.append(filename + ": relative public link: " + link)
                continue
            if parsed.netloc == "github.com" and parsed.path.startswith(REPOSITORY + "blob/"):
                ref, _, path = unquote(parsed.path[len(REPOSITORY + "blob/"):]).partition("/")
                if ref != "v" + version:
                    errors.append(filename + ": source link uses an outdated release: " + link)
                if not (ROOT / path).is_file():
                    errors.append(filename + ": source link points to a missing file: " + path)
            if parsed.netloc == "github.com" and parsed.path.startswith(REPOSITORY + "releases/tag/"):
                if parsed.path != REPOSITORY + "releases/tag/v" + version:
                    errors.append(filename + ": Release link uses an outdated version: " + link)
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    listing = re.search(r"^子命令：(.*)$", readme, re.M)
    documented = set(re.findall(r"`([^`]+)`", listing.group(1))) if listing else set()
    tree = ast.parse((ROOT / "src/acks_office/cli.py").read_text(encoding="utf-8"))
    commands = {
        call.args[0].value
        for call in ast.walk(tree)
        if isinstance(call, ast.Call) and isinstance(call.func, ast.Attribute)
        and call.func.attr == "add_parser" and call.args
        and isinstance(call.args[0], ast.Constant) and isinstance(call.args[0].value, str)
    }
    if documented != commands:
        errors.append("README CLI command list must match the actual parser.")
    for filename, pattern in (
        ("skills/acks-office/SKILL.md", r'^  version: "([^"]+)"'),
        ("skills/acks-office/scripts/acks.py", r'^VERSION = "([^"]+)"'),
    ):
        match = re.search(pattern, (ROOT / filename).read_text(encoding="utf-8"), re.M)
        if not match or match.group(1) != version:
            errors.append(filename + ": version does not match " + version)
    changelog = (ROOT / "CHANGELOG.md").read_text(encoding="utf-8")
    if not re.search(r"^## \[" + re.escape(version) + r"\]", changelog, re.M):
        errors.append("CHANGELOG must include the package version.")
    if errors:
        raise SystemExit("\n".join(errors))
    print("Release documentation checks passed for " + version)


if __name__ == "__main__":
    main()
