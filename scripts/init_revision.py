#!/usr/bin/env python3
"""Create an isolated PPT revision workspace without editing source files.

Requires a new, nonexistent destination. Preserves source bytes and copies the
included configuration templates. Python 3.9+, standard library only.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open('rb') as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b''):
            h.update(chunk)
    return h.hexdigest()


def validate_pptx(path: Path) -> Path:
    path = path.resolve(strict=True)
    if not path.is_file() or path.suffix.lower() != '.pptx':
        raise ValueError(f'Expected a .pptx file: {path}')
    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        if not {'[Content_Types].xml', 'ppt/presentation.xml'} <= names:
            raise ValueError(f'Not a recognizable PPTX package: {path}')
    return path


def initialize(source: Path, project_dir: Path, order_reference: Optional[Path] = None) -> Dict[str, Any]:
    source = validate_pptx(source)
    reference = validate_pptx(order_reference) if order_reference else None
    target = project_dir.absolute()
    if target.exists() or target.is_symlink():
        raise FileExistsError(f'Project directory already exists; refusing to overwrite: {target}')
    if not target.name:
        raise ValueError('Invalid project directory.')
    originals = [(source, 'baseline/source_editable.pptx', 'source_editable')]
    if reference:
        originals.append((reference, 'baseline/order_reference.pptx', 'order_reference'))
    hashes = {str(p): sha256_file(p) for p, _, _ in originals}
    target.parent.mkdir(parents=True, exist_ok=True)
    target.mkdir(exist_ok=False)
    try:
        for directory in ('baseline', 'work', 'assets', 'output', 'qc', 'handoff', 'config'):
            (target / directory).mkdir()
        entries = []
        for original, relative, role in originals:
            copied = target / relative
            shutil.copy2(original, copied)
            if sha256_file(copied) != hashes[str(original)]:
                raise RuntimeError(f'Copy hash mismatch for {original}')
            entries.append({'role': role, 'source_path': str(original), 'backup_path': relative,
                            'size_bytes': original.stat().st_size, 'sha256': hashes[str(original)]})
        working = target / 'work/current_editable.pptx'
        shutil.copy2(target / 'baseline/source_editable.pptx', working)
        template_root = Path(__file__).resolve().parents[1] / 'templates'
        if template_root.exists():
            for file in template_root.iterdir():
                if file.is_file():
                    shutil.copy2(file, target / 'config' / file.name)
        for original, _, _ in originals:
            if sha256_file(original) != hashes[str(original)]:
                raise RuntimeError(f'Input changed during initialization: {original}')
        manifest = {'schema_version': 1, 'created_at_utc': datetime.now(timezone.utc).isoformat(),
                    'project_dir': str(target), 'original_files': entries,
                    'working_copy': 'work/current_editable.pptx',
                    'state': 'initialized_not_edited',
                    'notes': ['Source files were read and copied only.',
                              'Baseline files are protected by workflow convention; do not edit them.',
                              'Fill config/project.yaml and map slides before modification.',
                              'No slide rendering, evidence review, or visual verification has occurred.']}
        (target / 'baseline_manifest.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
        return manifest
    except Exception as exc:
        (target / 'INITIALIZATION_FAILED.txt').write_text(
            f'Initialization did not complete. Source files were not intentionally modified.\nError: {exc}\n', encoding='utf-8')
        raise


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--source', type=Path, required=True)
    parser.add_argument('--project-dir', type=Path, required=True)
    parser.add_argument('--order-reference', type=Path)
    args = parser.parse_args(argv)
    try:
        result = initialize(args.source, args.project_dir, args.order_reference)
        print(json.dumps({'project_dir': result['project_dir'], 'state': result['state'],
                          'working_copy': result['working_copy']}, ensure_ascii=False))
        return 0
    except (OSError, ValueError, RuntimeError, zipfile.BadZipFile) as exc:
        print(f'Initialization failed: {exc}', file=sys.stderr)
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
