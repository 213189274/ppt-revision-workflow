"""Structural fixture tests only. Never reads or changes user presentation files."""
from __future__ import annotations

import importlib.util
import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

ROOT = Path(__file__).resolve().parents[1]


def load_module(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


audit = load_module('skill_pptx_preflight', ROOT / 'scripts/pptx_preflight.py')
init = load_module('skill_init_revision', ROOT / 'scripts/init_revision.py')

P = audit.NS['p']
A = audit.NS['a']
R = audit.NS['r']
RELS = audit.NS['rel']


def fixture(path, texts=None, order=None, image=False, broken=False, external=False,
            out_of_bounds=False, size=(12192000, 6858000), dtd=False):
    """Make a small OOXML structure fixture, not a visually rendered PPT."""
    texts = ['测试标题', '第二页'] if texts is None else texts
    order = list(range(1, len(texts) + 1)) if order is None else order
    width, height = size
    pres = f'<p:presentation xmlns:p="{P}" xmlns:r="{R}"><p:sldIdLst>'
    for i in order:
        pres += f'<p:sldId id="{255 + i}" r:id="rId{i}"/>'
    pres += f'</p:sldIdLst><p:sldSz cx="{width}" cy="{height}"/></p:presentation>'
    rels = f'<Relationships xmlns="{RELS}">' + ''.join(
        f'<Relationship Id="rId{i}" Type="{R}/slide" Target="slides/slide{i}.xml"/>'
        for i in range(1, len(texts) + 1)) + '</Relationships>'
    with zipfile.ZipFile(path, 'w', compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>')
        z.writestr('ppt/presentation.xml', pres)
        z.writestr('ppt/_rels/presentation.xml.rels', rels)
        for i, text in enumerate(texts, 1):
            x = width + 1000 if out_of_bounds else 0
            body = ''
            if text:
                body += (f'<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/></p:nvSpPr>'
                         f'<p:spPr><a:xfrm><a:off x="{x}" y="200000"/><a:ext cx="2000000" cy="700000"/></a:xfrm></p:spPr>'
                         '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r>'
                         '<a:rPr sz="2800" b="1"><a:latin typeface="Microsoft YaHei"/><a:ea typeface="微软雅黑"/></a:rPr>'
                         f'<a:t>{escape(text)}</a:t></a:r></a:p></p:txBody></p:sp>')
            srels = []
            if image:
                body += (f'<p:pic><p:nvPicPr><p:cNvPr id="3" name="Full background"/></p:nvPicPr>'
                         '<p:blipFill><a:blip r:embed="rImage"/></p:blipFill>'
                         f'<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{width}" cy="{height}"/></a:xfrm></p:spPr></p:pic>')
                srels.append(f'<Relationship Id="rImage" Type="{R}/image" Target="../media/fixture.bin"/>')
            if external:
                srels.append(f'<Relationship Id="rLink" Type="{R}/hyperlink" Target="https://example.invalid/evidence" TargetMode="External"/>')
            xml = f'<p:sld xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}"><p:cSld><p:spTree>{body}</p:spTree></p:cSld></p:sld>'
            if dtd:
                xml = '<!DOCTYPE sample [<!ENTITY x "blocked">]>' + xml
            z.writestr(f'ppt/slides/slide{i}.xml', xml)
            if srels:
                z.writestr(f'ppt/slides/_rels/slide{i}.xml.rels', f'<Relationships xmlns="{RELS}">' + ''.join(srels) + '</Relationships>')
        if image and not broken:
            # Opaque bytes to exercise relationship/hashing code; not a generated image.
            z.writestr('ppt/media/fixture.bin', b'nonvisual-unit-test-asset')


class WorkflowTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.file = self.root / 'source.pptx'

    def tearDown(self):
        self.temp.cleanup()

    def test_order_comes_from_presentation_not_filename(self):
        fixture(self.file, order=[2, 1])
        report = audit.inspect_pptx(self.file)
        self.assertEqual(report['slide_count'], 2)
        self.assertEqual(report['slides'][0]['slide_part'], 'ppt/slides/slide2.xml')
        self.assertEqual(report['errors'], [])

    def test_font_and_bold_are_declared_not_visual_verification(self):
        fixture(self.file)
        report = audit.inspect_pptx(self.file)
        self.assertIn('ea:微软雅黑', report['declared_font_summary'])
        self.assertEqual(report['slides'][0]['explicit_bold_runs'], 1)
        self.assertEqual(report['visual_status'], 'not_checked')

    def test_missing_media_target_is_error(self):
        fixture(self.file, image=True, broken=True)
        report = audit.inspect_pptx(self.file)
        self.assertTrue(report['errors'])
        self.assertTrue(report['missing_internal_targets'])

    def test_image_only_candidate(self):
        fixture(self.file, texts=[''], image=True)
        report = audit.inspect_pptx(self.file)
        self.assertTrue(report['slides'][0]['image_backed_candidate'])
        self.assertEqual(report['summary']['native_text_shapes'], 0)

    def test_native_text_over_background_not_image_only(self):
        fixture(self.file, texts=['可编辑标题'], image=True)
        report = audit.inspect_pptx(self.file)
        self.assertFalse(report['slides'][0]['image_backed_candidate'])
        self.assertEqual(report['summary']['native_text_shapes'], 1)

    def test_flag_does_not_change_source(self):
        fixture(self.file, texts=['保留模拟与拟合，检查拟使用'])
        before = audit.sha256_file(self.file)
        report = audit.inspect_pptx(self.file, flag_text=['拟'])
        self.assertTrue(report['slides'][0]['flagged_text'])
        self.assertEqual(before, audit.sha256_file(self.file))
        self.assertTrue(report['input_unchanged_during_inspection'])

    def test_bounds_are_candidates(self):
        fixture(self.file, out_of_bounds=True)
        report = audit.inspect_pptx(self.file)
        self.assertTrue(report['slides'][0]['top_level_bounds_candidates'])
        self.assertEqual(report['errors'], [])

    def test_external_links_are_not_opened(self):
        fixture(self.file, external=True)
        report = audit.inspect_pptx(self.file)
        self.assertEqual(len(report['external_relationships']), 2)
        self.assertEqual(report['errors'], [])

    def test_max_mb_and_expected_count(self):
        fixture(self.file)
        report = audit.inspect_pptx(self.file, expected_pages=3, max_mb=.0001)
        self.assertEqual(len(report['errors']), 2)

    def test_comparison_mismatch_and_candidates(self):
        fixture(self.file)
        other = self.root / 'other.pptx'
        fixture(other, texts=['测试标题'])
        a, b = audit.inspect_pptx(self.file), audit.inspect_pptx(other)
        result = audit.compare_reports(a, b)
        self.assertFalse(result['page_count_equal'])
        self.assertTrue(a['errors'])
        self.assertEqual(result['candidate_mapping'][0]['candidates'][0]['candidate_page'], 1)

    def test_reject_dtd(self):
        fixture(self.file, dtd=True)
        with self.assertRaises(audit.AuditError):
            audit.inspect_pptx(self.file)

    def test_reject_non_zip(self):
        self.file.write_bytes(b'not a presentation')
        with self.assertRaises(zipfile.BadZipFile):
            audit.inspect_pptx(self.file)

    def test_cli_writes_json_without_overwriting_existing_report(self):
        fixture(self.file)
        dest = self.root / 'report.json'
        code = audit.main([str(self.file), '--report', str(dest)])
        self.assertEqual(code, 0)
        data = json.loads(dest.read_text(encoding='utf-8'))
        self.assertEqual(data['slide_count'], 2)
        before = dest.read_bytes()
        code = audit.main([str(self.file), '--report', str(dest)])
        self.assertEqual(code, 1)
        self.assertEqual(dest.read_bytes(), before)

    def test_init_copies_and_preserves_original(self):
        fixture(self.file)
        before = audit.sha256_file(self.file)
        target = self.root / 'revision'
        result = init.initialize(self.file, target, self.file)
        self.assertEqual(before, audit.sha256_file(self.file))
        self.assertEqual(before, audit.sha256_file(target / 'baseline/source_editable.pptx'))
        self.assertEqual(before, audit.sha256_file(target / 'work/current_editable.pptx'))
        self.assertEqual(result['state'], 'initialized_not_edited')
        self.assertTrue((target / 'config/project.yaml').exists())

    def test_init_refuses_existing_directory(self):
        fixture(self.file)
        target = self.root / 'revision'
        target.mkdir()
        keep = target / 'keep.txt'
        keep.write_text('do not overwrite')
        with self.assertRaises(FileExistsError):
            init.initialize(self.file, target)
        self.assertEqual(keep.read_text(), 'do not overwrite')

    def test_invalid_source_does_not_create_workspace(self):
        self.file.write_bytes(b'bad')
        target = self.root / 'revision'
        with self.assertRaises(zipfile.BadZipFile):
            init.initialize(self.file, target)
        self.assertFalse(target.exists())


if __name__ == '__main__':
    unittest.main()
