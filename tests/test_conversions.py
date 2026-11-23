"""Smoke-test all PDF reverse formats using only generated content."""
from pathlib import Path
import sys, tempfile, unittest, zipfile
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from reportlab.pdfgen import canvas
from pdf_engine import convert_pdf
from pdf_organizer.model import sha256


class ReverseFormatTests(unittest.TestCase):
    def test_three_reverse_formats_are_complete_packages_and_preserve_input(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder);source=root/'synthetic.pdf'
            c=canvas.Canvas(str(source),pagesize=(400,500),invariant=1)
            c.setFont('Helvetica',14);c.drawString(40,450,'Ruyi public test')
            c.drawString(40,400,'Item');c.drawString(210,400,'Value')
            c.drawString(40,360,'Example');c.drawString(210,360,'42')
            for y in (340,380,420):c.line(30,y,340,y)
            for x in (30,190,340):c.line(x,340,x,420)
            c.save();before=sha256(source)
            for target,marker in (('docx','word/document.xml'),('xlsx','xl/workbook.xml'),('pptx','ppt/presentation.xml')):
                with self.subTest(target=target):
                    output=root/target;output.mkdir()
                    result=convert_pdf(source,output,target)
                    with zipfile.ZipFile(result.output_path) as package:
                        self.assertIsNone(package.testzip())
                        self.assertIn(marker,package.namelist())
                        self.assertFalse(any('thumbnail' in n.lower() for n in package.namelist()))
                    self.assertEqual(before,sha256(source))


if __name__=='__main__':unittest.main(verbosity=2)
