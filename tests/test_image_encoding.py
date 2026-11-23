from pathlib import Path
from io import BytesIO
from unittest.mock import patch
import sys, unittest
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from PIL import Image, ImageDraw
import numpy as np
import pdf_word_images as codec


class EncodingTests(unittest.TestCase):
    def test_alpha_and_palette_transparency_stay_png(self):
        for mode in ('RGBA','LA','P'):
            im=Image.new(mode,(320,240))
            if mode=='P': im.info['transparency']=0
            self.assertEqual(codec.choose_image_encoding(im)[0],'png')
            encoded=codec.encode_word_image(im)
            with Image.open(encoded) as out:
                self.assertEqual(out.size,im.size)
                self.assertIn('A',out.convert('RGBA').getbands())
                self.assertEqual(out.convert('RGBA').tobytes(),im.convert('RGBA').tobytes())
        im=Image.new('RGBA',(400,400),(23,45,67,255))
        self.assertEqual(codec.choose_image_encoding(im)[0],'png')

    def test_small_pixel_or_byte_count_precedes_photographic_rule(self):
        rng=np.random.default_rng(42)
        for size in ((199,600),(600,199)):
            im=Image.fromarray(rng.integers(0,256,(*size[::-1],3),dtype='uint8'))
            self.assertEqual(codec.choose_image_encoding(im)[0],'png')
        im=Image.fromarray(rng.integers(0,256,(300,300,3),dtype='uint8'))
        self.assertEqual(codec.choose_image_encoding(im,codec.SourceImage(b'few bytes'))[0],'png')

    def test_binary_and_antialiasing_are_lossless_gray(self):
        im=Image.new('RGB',(500,500),'white');ImageDraw.Draw(im).rectangle((10,10,100,12),fill='black')
        self.assertEqual(codec.choose_image_encoding(im)[1]['mode'],'1')
        im.putpixel((20,20),(180,180,180))
        self.assertEqual(codec.choose_image_encoding(im)[1]['mode'],'L')
        with Image.open(codec.encode_word_image(im,codec.SourceImage(b'0'*40000))) as out:
            self.assertEqual(out.mode,'L');self.assertEqual(out.convert('RGB').tobytes(),im.tobytes())
        im.putpixel((22,22),(255,0,0))
        self.assertNotIn('mode',codec.choose_image_encoding(im)[1])

    def test_lines_tables_and_flat_colors_are_lossless(self):
        im=Image.new('RGB',(600,400),'#a8cfee');d=ImageDraw.Draw(im)
        for x in range(0,600,12):d.line((x,0,x,400),fill='#2277ff',width=2)
        d.text((20,20),'Formula x=1+2; table 12345',fill='black')
        self.assertEqual(codec.choose_image_encoding(im)[0],'png')
        with Image.open(codec.encode_word_image(im)) as out:self.assertEqual(out.tobytes(),im.tobytes())

    def test_gradient_photo_jpeg_quality_and_no_geometry_change(self):
        rng=np.random.default_rng(6)
        base=np.linspace(0,250,640)[None,:,None]+np.zeros((480,1,3))
        im=Image.fromarray(np.uint8(np.clip(base+rng.normal(0,2,base.shape),0,255)))
        ext,options=codec.choose_image_encoding(im)
        self.assertEqual((ext,options['quality'],options['subsampling']),('jpg',88,2))
        with Image.open(codec.encode_word_image(im)) as out:
            self.assertEqual(out.size,im.size);self.assertEqual(out.format,'JPEG')
            diff=np.abs(np.asarray(out.convert('L'),dtype='int16')-np.asarray(im.convert('L'),dtype='int16'))
            self.assertLessEqual(diff.mean(),2)

    def test_source_reuse_requires_all_guards_and_is_byte_identical(self):
        im=Image.new('RGB',(300,300),'#76895e');stream=BytesIO();im.save(stream,format='JPEG',quality=96)
        raw=stream.getvalue()
        for flags in ((False,True,True),(True,False,True),(True,True,False)):
            self.assertNotIn('source',codec.choose_image_encoding(im,codec.SourceImage(raw,*flags))[1])
        info=codec.SourceImage(raw,True,True,True)
        self.assertEqual(codec.encode_word_image(im,info).getvalue(),raw)
        rgba=im.convert('RGBA')
        self.assertEqual(codec.choose_image_encoding(rgba,info)[0],'png')

    def test_real_pdf_source_jpeg_and_modified_or_scaled_fallback(self):
        import pypdfium2 as pdfium
        from pdf_docx_layout import _PagePictures
        im=Image.new('RGB',(200,200),(128,128,128));jpeg=BytesIO();im.save(jpeg,format='JPEG',quality=95)
        raw=jpeg.getvalue();pdf=pdfium.PdfDocument.new();page=pdf.new_page(100,100)
        obj=pdfium.PdfImage.new(pdf);obj.load_jpeg(BytesIO(raw),inline=True)
        obj.set_matrix(pdfium.PdfMatrix().scale(100,100));page.insert_obj(obj);page.gen_content()
        pics=_PagePictures(page,100,100)
        try:
            out=pics.crop((0,0,100,100));self.assertEqual(out.getvalue(),raw)
            with Image.open(pics.crop((0,0,99,100))) as partial:self.assertEqual(partial.format,'PNG')
            changed=im.copy();changed.putpixel((100,100),(0,0,0))
            self.assertIsNone(codec.matching_source_jpeg(page,(0,0,100,100),changed,100))
            self.assertIsNone(codec.matching_source_jpeg(page,(0,0,100,100),im.resize((400,400)),100))
        finally:pics.close();page.close();pdf.close()

    def test_whole_feature_can_roll_back_to_original_png(self):
        im=Image.new('RGB',(300,300),'white');im.putpixel((2,3),(17,28,33));original=BytesIO();im.save(original,format='PNG')
        with patch.object(codec,'OPTIMIZE_WORD_IMAGES',False):
            self.assertEqual(codec.encode_word_image(im,codec.SourceImage(b'jpeg',True,True,True)).getvalue(),original.getvalue())


if __name__=='__main__':unittest.main(verbosity=2)
