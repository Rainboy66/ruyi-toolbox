"""Public synthetic fixtures: never read private documents."""
from dataclasses import replace
from pathlib import Path
import sys
import tempfile
import unittest
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from pypdf import PdfReader
from reportlab.pdfgen import canvas
from pdf_organizer.model import PlanDocument, Sheet, inspect_sources, sha256
from pdf_organizer.exporter import export_pdf, split_pdf
from pdf_organizer.presets import save_preset, load_preset, apply_preset, batch_apply


def fixture(path, count, prefix='P'):
    c = canvas.Canvas(str(path), pagesize=(595.275591, 841.889764), invariant=1)
    for i in range(count):
        c.setFont('Helvetica', 20)
        c.drawString(48, 740, f'{prefix}{i+1:03}')
        c.rect(30, 30, 535, 780)
        c.showPage()
    c.save()


class PublicPDFTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.source = self.folder / 'source.pdf'
        fixture(self.source, 20)
        self.doc = PlanDocument()
        self.doc.import_sources(inspect_sources([self.source]))

    def render(self, plan, name='output.pdf'):
        target = self.folder / name
        export_pdf(plan, self.doc.sources, target)
        return PdfReader(target)

    def test_hundred_pages_edit_order_rotate_and_unchanged_sources(self):
        paths = [self.folder / f'{p}.pdf' for p in 'ABC']
        for p, count, label in zip(paths, (40,35,25), 'ABC'):
            fixture(p, count, label)
        d = PlanDocument(); d.import_sources(inspect_sources(paths))
        before = [sha256(p) for p in paths]
        d.move((0,77),90); d.select((4,5,39,60,88)); d.delete()
        d.select((3,45)); d.rotate()
        out = self.folder / 'edited.pdf'; export_pdf(d.plan,d.sources,out)
        pages = PdfReader(out).pages
        self.assertEqual(len(pages),95)
        for page, ref in zip(pages,d.plan.pages):
            label = Path(d.sources[ref.source_id].path).stem
            self.assertIn(f'{label}{ref.page_index+1:03}',page.extract_text())
            self.assertEqual(page.rotation,ref.rotate)
        self.assertEqual(before,[sha256(p) for p in paths])

    def test_plan_save_reload_and_fifty_step_history(self):
        d=self.doc; initial=d.plan; d.select((0,))
        for _ in range(30):d.rotate()
        for _ in range(30):self.assertTrue(d.undo())
        self.assertEqual(initial,d.plan)
        for _ in range(30):self.assertTrue(d.redo())
        path=self.folder/'plan.ruyiplan.json';d.save(path)
        self.assertEqual(d.plan,PlanDocument.load(path).plan)
        for _ in range(60):d.rotate()
        self.assertEqual(len(d.undo_stack),50)

    def test_nup_counts_text_and_page_coordinates(self):
        import pdfplumber
        for n,count in ((2,10),(4,5),(6,4),(9,3)):
            for order in ('z','column'):
                plan=replace(self.doc.plan,rules=({'op':'nup','n':n,'order':order,'margin':24,'gap':8,'border':True},))
                pages=self.render(plan,f'{n}-{order}.pdf').pages
                self.assertEqual(len(pages),count)
                self.assertTrue(all(f'P{i:03}' in pages[0].extract_text() for i in range(1,n+1)))
        plan=replace(self.doc.plan,rules=({'op':'nup','n':4,'margin':24,'gap':8},))
        self.render(plan,'positions.pdf')
        with pdfplumber.open(self.folder/'positions.pdf') as pdf:
            words={w['text']:w for w in pdf.pages[0].extract_words() if w['text'].startswith('P')}
            self.assertLess(words['P001']['x0'],words['P002']['x0'])
            self.assertLess(words['P001']['top'],words['P003']['top'])

    def test_booklet_order_and_padding(self):
        plan=replace(self.doc.plan,rules=({'op':'booklet','binding':'left','duplex':'short'},))
        pages=self.render(plan).pages
        self.assertEqual(len(pages),10)
        self.assertLess(pages[0].extract_text().index('P020'),pages[0].extract_text().index('P001'))
        padded=replace(plan,pages=plan.pages[:18])
        pages=self.render(padded,'padded.pdf').pages
        self.assertEqual(len(pages),10)
        self.assertEqual(sum(p.extract_text().count('P') for p in pages),18)

    def test_rules_preset_reproducible_and_batch_continues(self):
        plan=replace(self.doc.plan,rules=({'op':'nup','n':4},{'op':'overlay','overlay':{'kind':'page_number'}}))
        preset_path=self.folder/'preset.ruyiplan.json';save_preset(plan,preset_path)
        preset=load_preset(preset_path)
        self.render(plan,'first.pdf');self.render(apply_preset(self.doc.plan,preset),'second.pdf')
        self.assertEqual((self.folder/'first.pdf').read_bytes(),(self.folder/'second.pdf').read_bytes())
        bad=self.folder/'broken.pdf';bad.write_bytes(b'not a PDF')
        results=batch_apply([self.source,bad,self.source],preset,self.folder/'batch')
        self.assertEqual([r['ok'] for r in results],[True,False,True])

    def test_fit_fill_stretch_split_and_source_overwrite_guard(self):
        for mode in ('fit','fill','stretch'):
            plan=replace(self.doc.plan,pages=(replace(self.doc.plan.pages[0],rotate=90,fit_mode=mode),),sheet=Sheet((612,792),0,24))
            page=self.render(plan,f'{mode}.pdf').pages[0]
            self.assertEqual(tuple(page.mediabox),(0,0,612,792))
            self.assertIn('P001',page.extract_text())
        parts=split_pdf(self.doc.plan,self.doc.sources,self.folder/'split',7)
        self.assertEqual([len(PdfReader(p).pages) for p in parts],[7,7,6])
        with self.assertRaises(ValueError):export_pdf(self.doc.plan,self.doc.sources,self.source)

    def test_long_page_physical_height(self):
        plan=replace(self.doc.plan,rules=({'op':'concat','margin':24,'gap':10},))
        pages=self.render(plan).pages
        self.assertEqual(len(pages),1)
        page=pages[0]
        actual=float(page.mediabox.height)*float(page.get('/UserUnit',1))
        expected=20*float(PdfReader(self.source).pages[0].mediabox.height)+48+19*10
        self.assertAlmostEqual(actual,expected,places=3)


if __name__=='__main__':unittest.main(verbosity=2)
