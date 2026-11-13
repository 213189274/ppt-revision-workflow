#!/usr/bin/env python3
"""Read-only PPTX structural preflight. No rendering, OCR, or network access.

Exit codes: 0 structural checks passed (warnings may remain); 2 findings with
hard errors; 1 execution/input/report-writing failure. Visual and factual review
are always separate. Python 3.9+; standard library only.
"""
from __future__ import annotations

import argparse
import difflib
import hashlib
import json
import math
import posixpath
import re
import sys
import zipfile
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import unquote
from xml.etree import ElementTree as ET

NS = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
}
MAX_MEMBERS = 100_000
MAX_XML_BYTES = 20_000_000
MAX_XML_TOTAL = 200_000_000
MAX_HASHED_ASSET_BYTES = 256_000_000


class AuditError(ValueError):
    """Unreadable or unsupported source package."""


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open('rb') as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b''):
            h.update(chunk)
    return h.hexdigest()


def _xml(zf: zipfile.ZipFile, name: str) -> ET.Element:
    info = zf.getinfo(name)
    if info.file_size > MAX_XML_BYTES:
        raise AuditError(f'XML part too large: {name}')
    data = zf.read(name)
    if b'<!DOCTYPE' in data.upper() or b'<!ENTITY' in data.upper():
        raise AuditError(f'DTD/entity declarations are not accepted: {name}')
    try:
        return ET.fromstring(data)
    except ET.ParseError as exc:
        raise AuditError(f'Invalid XML in {name}: {exc}') from exc


def _source_for_rels(name: str) -> str:
    if name == '_rels/.rels':
        return ''
    directory, basename = posixpath.split(name)
    if posixpath.basename(directory) != '_rels' or not basename.endswith('.rels'):
        raise AuditError(f'Unexpected relationship part path: {name}')
    return posixpath.join(posixpath.dirname(directory), basename[:-5])


def _target_part(source: str, target: str) -> str:
    target = unquote(target.split('#', 1)[0])
    if target.startswith('/'):
        resolved = posixpath.normpath(target.lstrip('/'))
    else:
        resolved = posixpath.normpath(posixpath.join(posixpath.dirname(source), target))
    if resolved == '..' or resolved.startswith('../'):
        raise AuditError(f'Relationship escapes package root: {source} -> {target}')
    return resolved


def _rels_for_part(part: str) -> str:
    return posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part) + '.rels')


def _normal_text(text: str) -> str:
    return re.sub(r'\s+', '', text)


def _hash_asset(zf: zipfile.ZipFile, name: str) -> Optional[str]:
    if zf.getinfo(name).file_size > MAX_HASHED_ASSET_BYTES:
        return None
    h = hashlib.sha256()
    with zf.open(name) as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b''):
            h.update(chunk)
    return h.hexdigest()


def _bounds(shape: ET.Element) -> Optional[Tuple[int, int, int, int]]:
    transform = shape.find('p:spPr/a:xfrm', NS)
    if transform is None:
        transform = shape.find('p:xfrm', NS)
    if transform is None:
        return None
    offset = transform.find('a:off', NS)
    extent = transform.find('a:ext', NS)
    if offset is None or extent is None:
        return None
    try:
        return (int(offset.attrib['x']), int(offset.attrib['y']),
                int(extent.attrib['cx']), int(extent.attrib['cy']))
    except (ValueError, KeyError):
        return None


def inspect_pptx(path: Path, expected_pages: Optional[int] = None,
                 max_mb: Optional[float] = None, flag_text: Optional[List[str]] = None,
                 deep_crc: bool = False) -> Dict[str, Any]:
    path = path.resolve(strict=True)
    if not path.is_file():
        raise AuditError(f'Not a file: {path}')
    errors: List[str] = []
    warnings: List[str] = []
    original_hash = sha256_file(path)
    size = path.stat().st_size
    report: Dict[str, Any] = {
        'schema_version': 1,
        'checked_at_utc': datetime.now(timezone.utc).isoformat(),
        'input_path': str(path),
        'size_bytes': size,
        'size_mb_decimal': round(size / 1_000_000, 3),
        'sha256': original_hash,
        'errors': errors,
        'warnings': warnings,
        'limitations': [
            'Structural inspection only; no slide rendering or OCR.',
            'Declared fonts do not prove font installation, substitution behavior, or visual glyphs.',
            'Bounds are top-level unrotated shape candidates, not a complete overlap detector.',
            'Object counts do not prove complete editability or factual accuracy.',
            'Comparison mappings are suggestions and require visual confirmation.',
        ],
    }
    if max_mb is not None and size > int(max_mb * 1_000_000):
        errors.append(f'File exceeds {max_mb:g} MB decimal ({int(max_mb * 1_000_000)} bytes).')

    with zipfile.ZipFile(path) as zf:
        infos = zf.infolist()
        if len(infos) > MAX_MEMBERS:
            raise AuditError('Too many ZIP members.')
        names = [i.filename for i in infos]
        name_set = set(names)
        duplicates = [name for name, count in Counter(names).items() if count > 1]
        if duplicates:
            errors.append(f'Duplicate ZIP member names: {duplicates[:10]}')
        xml_total = sum(i.file_size for i in infos if i.filename.endswith(('.xml', '.rels')))
        if xml_total > MAX_XML_TOTAL:
            raise AuditError('Total XML size exceeds safety limit.')
        for required in ('[Content_Types].xml', 'ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels'):
            if required not in name_set:
                raise AuditError(f'Missing required part: {required}')
        _xml(zf, '[Content_Types].xml')
        for name in names:
            if name.startswith('/') or '..' in name.split('/') or '\\' in name:
                errors.append(f'Unsafe/noncanonical ZIP member path: {name}')
        if deep_crc:
            bad_part = zf.testzip()
            report['crc_check'] = 'all_members'
            if bad_part:
                errors.append(f'CRC failure: {bad_part}')
        else:
            report['crc_check'] = 'parts_read_during_inspection_only'

        relationships: Dict[str, Dict[str, Dict[str, str]]] = {}
        external: List[Dict[str, str]] = []
        missing: List[Dict[str, str]] = []
        for name in names:
            if not name.endswith('.rels'):
                continue
            source = _source_for_rels(name)
            rel_map: Dict[str, Dict[str, str]] = {}
            for rel in _xml(zf, name):
                rid = rel.attrib.get('Id', '')
                target = rel.attrib.get('Target', '')
                rtype = rel.attrib.get('Type', '')
                mode = rel.attrib.get('TargetMode', 'Internal')
                if not rid or not target:
                    errors.append(f'Incomplete relationship in {name}')
                    continue
                if rid in rel_map:
                    errors.append(f'Duplicate relationship ID {rid} in {name}')
                item = {'type': rtype, 'mode': mode, 'target': target}
                if mode.lower() == 'external':
                    external.append({'source': source, 'type': rtype, 'target': target})
                else:
                    resolved = _target_part(source, target)
                    item['part'] = resolved
                    if resolved not in name_set:
                        missing.append({'source': source, 'target': resolved})
                rel_map[rid] = item
            relationships[source] = rel_map
        if missing:
            errors.append(f'{len(missing)} internal relationship target(s) missing.')
        report['missing_internal_targets'] = missing
        report['external_relationships'] = external
        if external:
            warnings.append('External relationships exist; none were fetched or verified.')

        presentation = _xml(zf, 'ppt/presentation.xml')
        size_el = presentation.find('p:sldSz', NS)
        if size_el is None:
            raise AuditError('Missing slide size declaration.')
        width, height = int(size_el.attrib['cx']), int(size_el.attrib['cy'])
        if width <= 0 or height <= 0:
            raise AuditError('Invalid slide dimensions.')
        report['slide_size_emu'] = {'width': width, 'height': height}
        report['aspect_ratio'] = round(width / height, 8)
        slide_ids = presentation.findall('p:sldIdLst/p:sldId', NS)
        report['slide_count'] = len(slide_ids)
        if not slide_ids:
            errors.append('Presentation contains no slides.')
        if expected_pages is not None and len(slide_ids) != expected_pages:
            errors.append(f'Expected {expected_pages} slides, found {len(slide_ids)}.')
        font_counts: Counter = Counter()
        asset_cache: Dict[str, Optional[str]] = {}
        slides: List[Dict[str, Any]] = []
        p_rels = relationships['ppt/presentation.xml']
        for page_no, sid in enumerate(slide_ids, start=1):
            rid = sid.attrib.get(f'{{{NS["r"]}}}id', '')
            relation = p_rels.get(rid, {})
            part = relation.get('part', '')
            if not part or part not in name_set:
                errors.append(f'Slide {page_no} has no readable part for relationship {rid}.')
                continue
            slide = _xml(zf, part)
            all_texts = [t.text or '' for t in slide.findall('.//a:t', NS)]
            text = '\n'.join(all_texts)
            normalized = _normal_text(text)
            native_shapes = [s for s in slide.findall('.//p:sp', NS) if s.findall('.//a:t', NS)]
            tables = slide.findall('.//a:tbl', NS)
            charts = slide.findall('.//c:chart', NS)
            pictures = slide.findall('.//p:pic', NS)
            s_fonts: Counter = Counter()
            for tag in ('latin', 'ea', 'cs'):
                for font in slide.findall(f'.//a:{tag}', NS):
                    name = font.attrib.get('typeface', '')
                    if name:
                        s_fonts[f'{tag}:{name}'] += 1
                        font_counts[f'{tag}:{name}'] += 1
            bold_yes = sum(1 for r in slide.findall('.//a:rPr', NS) if r.attrib.get('b') in ('1', 'true'))
            declared_runs = len(slide.findall('.//a:rPr', NS))
            s_rels = relationships.get(part, {})
            hashes = set()
            image_parts = set()
            for blip in slide.findall('.//a:blip', NS):
                embed = blip.attrib.get(f'{{{NS["r"]}}}embed')
                if not embed:
                    continue
                image_part = s_rels.get(embed, {}).get('part')
                if not image_part or image_part not in name_set:
                    errors.append(f'Slide {page_no}: missing embedded image relationship {embed}.')
                    continue
                image_parts.add(image_part)
                if image_part not in asset_cache:
                    asset_cache[image_part] = _hash_asset(zf, image_part)
                if asset_cache[image_part]:
                    hashes.add(asset_cache[image_part])
                else:
                    warnings.append(f'Large asset not hashed: {image_part}')
            tree = slide.find('p:cSld/p:spTree', NS)
            bound_candidates = []
            full_pictures = 0
            if tree is not None:
                for item in tree:
                    if item.tag not in {f'{{{NS["p"]}}}{tag}' for tag in ('sp', 'pic', 'graphicFrame', 'cxnSp')}:
                        continue
                    bounds = _bounds(item)
                    if bounds is None:
                        continue
                    x, y, w, h = bounds
                    label = item.find('.//p:cNvPr', NS)
                    label_text = label.attrib.get('name', '') if label is not None else ''
                    if x < -width * .005 or y < -height * .005 or x + w > width * 1.005 or y + h > height * 1.005:
                        bound_candidates.append({'name': label_text, 'bounds_emu': list(bounds)})
                    if item.tag == f'{{{NS["p"]}}}pic' and w >= width * .95 and h >= height * .95:
                        full_pictures += 1
            raster_background = bool(slide.findall('p:cSld/p:bg//a:blip', NS))
            has_notes = any(r.get('type', '').endswith('/notesSlide') and r.get('part') in name_set for r in s_rels.values())
            flagged = []
            for term in flag_text or []:
                if term in text:
                    p = text.index(term)
                    flagged.append({'term': term, 'context': text[max(0, p - 40):p + len(term) + 60]})
            item_report = {
                'page': page_no,
                'slide_id': sid.attrib.get('id'),
                'slide_part': part,
                'hidden': slide.attrib.get('show') == '0',
                'native_text_shapes': len(native_shapes),
                'native_text_characters': len(normalized),
                'text_sample': text[:240],
                'text_sha256': hashlib.sha256(normalized.encode('utf-8')).hexdigest() if normalized else None,
                'tables': len(tables),
                'charts': len(charts),
                'pictures': len(pictures),
                'full_slide_picture_candidates': full_pictures,
                'raster_background': raster_background,
                'image_backed_candidate': bool(full_pictures or raster_background) and not normalized,
                'declared_fonts': dict(sorted(s_fonts.items())),
                'explicit_bold_runs': bold_yes,
                'explicit_run_style_count': declared_runs,
                'has_notes': has_notes,
                'top_level_bounds_candidates': bound_candidates,
                'flagged_text': flagged,
                'image_parts': sorted(image_parts),
                'image_sha256': sorted(hashes),
            }
            slides.append(item_report)
            if bound_candidates:
                warnings.append(f'Slide {page_no}: {len(bound_candidates)} top-level bounds candidate(s); review visually.')
            if flagged:
                warnings.append(f'Slide {page_no}: flagged text found; no replacement was performed.')
        report['slides'] = slides
        report['declared_font_summary'] = dict(sorted(font_counts.items()))
        report['summary'] = {
            'slides_inspected': len(slides),
            'native_text_shapes': sum(s['native_text_shapes'] for s in slides),
            'tables': sum(s['tables'] for s in slides),
            'charts': sum(s['charts'] for s in slides),
            'pictures': sum(s['pictures'] for s in slides),
            'image_backed_candidate_pages': [s['page'] for s in slides if s['image_backed_candidate']],
            'hidden_pages': [s['page'] for s in slides if s['hidden']],
            'pages_without_notes': [s['page'] for s in slides if not s['has_notes']],
        }
    report['input_unchanged_during_inspection'] = sha256_file(path) == original_hash
    if not report['input_unchanged_during_inspection']:
        errors.append('Input changed during inspection; repeat with a stable file.')
    report['structural_status'] = 'failed' if errors else 'passed_with_warnings' if warnings else 'passed'
    report['visual_status'] = 'not_checked'
    report['factual_status'] = 'not_checked'
    return report


def compare_reports(primary: Dict[str, Any], other: Dict[str, Any]) -> Dict[str, Any]:
    page_count_equal = primary['slide_count'] == other['slide_count']
    ratio_equal = math.isclose(primary['aspect_ratio'], other['aspect_ratio'], rel_tol=1e-5)
    if not page_count_equal:
        primary['errors'].append('Compared presentations have different slide counts.')
    if not ratio_equal:
        primary['errors'].append('Compared presentations have different page aspect ratios.')
    candidates = []
    for slide in primary['slides']:
        scores = []
        for target in other['slides']:
            score, reasons = 0.0, []
            if slide['text_sha256'] and slide['text_sha256'] == target['text_sha256']:
                score, reasons = 1.0, ['identical_native_text']
            elif slide['text_sample'] and target['text_sample']:
                sim = difflib.SequenceMatcher(None, _normal_text(slide['text_sample']), _normal_text(target['text_sample'])).ratio()
                score = sim * .65
                reasons.append('partial_text_similarity_only')
            a, b = set(slide['image_sha256']), set(target['image_sha256'])
            if a and b and a & b:
                jaccard = len(a & b) / len(a | b)
                score = max(score, jaccard * .9)
                reasons.append('shared_image_assets')
            if score >= .25:
                scores.append({'candidate_page': target['page'], 'score': round(score, 4), 'basis': reasons})
        candidates.append({'page': slide['page'], 'candidates': sorted(scores, key=lambda x: x['score'], reverse=True)[:3]})
    primary['structural_status'] = 'failed' if primary['errors'] else primary['structural_status']
    return {
        'other_file': other['input_path'],
        'other_sha256': other['sha256'],
        'page_count_equal': page_count_equal,
        'aspect_ratio_equal': ratio_equal,
        'exact_page_size_equal': primary['slide_size_emu'] == other['slide_size_emu'],
        'candidate_mapping': candidates,
        'warning': 'Matching suggestions only. Rasterized and native slides may share no hashes or text. Confirm order and visuals manually.',
    }


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('pptx', type=Path)
    parser.add_argument('--report', type=Path, required=True)
    parser.add_argument('--expected-pages', type=int)
    parser.add_argument('--max-mb', type=float, help='Decimal MB: 1 MB = 1,000,000 bytes.')
    parser.add_argument('--compare', type=Path)
    parser.add_argument('--flag-text', action='append', default=[], help='Literal text to flag; never changes source.')
    parser.add_argument('--deep-crc', action='store_true')
    parser.add_argument('--overwrite-report', action='store_true')
    args = parser.parse_args(argv)
    try:
        if args.expected_pages is not None and args.expected_pages <= 0:
            raise AuditError('--expected-pages must be positive.')
        if args.max_mb is not None and (not math.isfinite(args.max_mb) or args.max_mb <= 0):
            raise AuditError('--max-mb must be finite and positive.')
        source = args.pptx.resolve(strict=True)
        dest = args.report.resolve()
        if dest.suffix.lower() != '.json':
            raise AuditError('Report path must end in .json.')
        if dest == source or (args.compare and dest == args.compare.resolve()):
            raise AuditError('Report path must not be an input file.')
        if dest.exists() and not args.overwrite_report:
            raise AuditError('Report already exists; use a new path or --overwrite-report.')
        report = inspect_pptx(source, args.expected_pages, args.max_mb, args.flag_text, args.deep_crc)
        if args.compare:
            other = inspect_pptx(args.compare)
            report['comparison'] = compare_reports(report, other)
            if other['errors']:
                report['errors'].append('Comparison file has structural errors; inspect it separately.')
                report['structural_status'] = 'failed'
        dest.parent.mkdir(parents=True, exist_ok=True)
        mode = 'w' if args.overwrite_report else 'x'
        with dest.open(mode, encoding='utf-8') as out:
            json.dump(report, out, ensure_ascii=False, indent=2)
            out.write('\n')
        print(json.dumps({'report': str(dest), 'slide_count': report['slide_count'],
                          'size_bytes': report['size_bytes'], 'errors': len(report['errors']),
                          'warnings': len(report['warnings']), 'visual_status': 'not_checked'}, ensure_ascii=False))
        return 2 if report['errors'] else 0
    except (AuditError, OSError, zipfile.BadZipFile, KeyError, ValueError, RuntimeError) as exc:
        print(f'Preflight failed: {exc}', file=sys.stderr)
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
