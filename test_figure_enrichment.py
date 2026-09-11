import unittest
import uuid, shutil
from contextlib import contextmanager
from pathlib import Path
from unittest.mock import patch
import pymupdf as f
import native_pdf_to_md as n
from figure_enrichment import image_regions, candidate_page, discover_regions

@contextmanager
def workspace_temp():
    root=Path(__file__).resolve().parent.parent / ('_figure_test_'+uuid.uuid4().hex)
    root.mkdir()
    try:
        yield root
    finally:
        shutil.rmtree(root)


class FigureTests(unittest.TestCase):
    def test_default_enhancement_and_no_upload_precedence(self):
        args = n.build_parser().parse_args(['sample.pdf'])
        self.assertEqual(n.resolve_figure_mode(args.route, args.figure_mode), 'mineru')
        self.assertEqual(n.resolve_figure_mode('native', args.figure_mode), 'local')
        self.assertEqual(n.resolve_figure_mode('auto', 'local'), 'local')
        with self.assertRaises(n.ConversionError):
            n.resolve_figure_mode('native', 'mineru')

    def test_default_conversion_calls_enhancement_and_preserves_output_on_failure(self):
        with workspace_temp() as directory, f.open() as doc:
            path = directory / 'sample.pdf'
            doc.new_page().insert_text((50, 50), 'Figure 1-1: Example diagram')
            doc.save(path)
            destination = directory / 'sample'
            destination.mkdir()
            previous = destination / 'sample.md'
            previous.write_text('previous output')
            with patch.object(n, 'validate_text_layer', return_value=1000), patch('figure_enrichment.discover_regions', side_effect=ValueError('missing token')) as discover:
                with self.assertRaises(n.ConversionError):
                    n.convert_pdf(path, directory, overwrite=True)
                discover.assert_called_once()
            self.assertEqual(previous.read_text(), 'previous output')

    def test_split_charts_restore_full_bitmap(self):
        class Page:
            rect=f.Rect(0,0,612,792)
            rotation=0
            def get_image_info(self): return [{'bbox':[50,80,350,400]}]
        info={'page_size':[612,792],'para_blocks':[
            {'type':'chart','blocks':[{'type':'chart_body','bbox':[55,90,345,220]}]},
            {'type':'chart','blocks':[{'type':'chart_body','bbox':[55,240,345,390]}]}]}
        rects=image_regions(info,Page())
        self.assertEqual(len(rects),1)
        self.assertLessEqual(rects[0].y0,80)
        self.assertGreaterEqual(rects[0].y1,400)

    def test_short_diagram_caption_label_is_recovered(self):
        with f.open() as doc:
            page=doc.new_page(width=612,height=792)
            page.insert_text((60,150),'Argument (not all bits may be used)',fontsize=10)
            info={'page_size':[612,792],'para_blocks':[
                {'type':'image','blocks':[
                    {'type':'image_body','bbox':[50,80,300,135]},
                    {'type':'image_caption','bbox':[60,139,250,153],
                     'lines':[{'spans':[{'content':'Argument (not all bits may be used)'}]}]}]}]}
            self.assertGreater(image_regions(info,page)[0].y1,150)

    def test_chart_and_separate_figure_caption_table(self):
        class Page:
            rect=f.Rect(0,0,612,792)
            rotation=0
            def get_image_info(self): return []
        caption={'type':'title','bbox':[50,50,300,65],
                 'lines':[{'spans':[{'content':'Figure 1-1: Timing'}]}]}
        info={'page_size':[612,792],'para_blocks':[caption,
              {'type':'table','bbox':[50,80,300,180],'blocks':[
                  {'type':'table_body','bbox':[50,80,300,180]}]},
              {'type':'chart','bbox':[50,300,300,400],'blocks':[
                  {'type':'chart_body','bbox':[50,300,300,400]}]}]}
        self.assertEqual(len(image_regions(info,Page())),2)

    def test_caption_below_previous_table_does_not_convert_table(self):
        class Page:
            rect=f.Rect(0,0,612,792)
            rotation=0
            def get_image_info(self): return []
        info={'page_size':[612,792],'para_blocks':[
            {'type':'table','bbox':[50,60,500,180],'blocks':[
                {'type':'table_body','bbox':[50,60,500,180]},
                {'type':'table_caption','bbox':[100,200,400,215],
                 'lines':[{'spans':[{'content':'Figure 1-1: Diagram below'}]}]}]},
            {'type':'image','blocks':[{'type':'image_body','bbox':[50,250,500,400]}]}]}
        rects=image_regions(info,Page())
        self.assertEqual(len(rects),1)
        self.assertGreater(rects[0].y0,200)

    def test_labels_recovered_but_prose_protected_and_nested_body_deduped(self):
        with f.open() as doc:
            page=doc.new_page(width=612,height=792)
            page.insert_text((155,110),'Side label',fontsize=10)
            page.insert_text((60,154),'Ordinary paragraph',fontsize=10)
            info={'page_size':[612,792],'para_blocks':[
                {'type':'image','blocks':[{'type':'image_body','bbox':[50,80,150,135]}]},
                {'type':'image','blocks':[{'type':'image_body','bbox':[60,90,140,120]}]},
                {'type':'text','bbox':[55,142,250,160]}]}
            rects=image_regions(info,page)
            self.assertEqual(len(rects),1)
            self.assertGreater(rects[0].x1,190)
            self.assertLess(rects[0].y1,142)

    def test_wrapped_figure_list_is_not_an_upload_candidate(self):
        class Page:
            def get_text(self, kind):
                return [(0,0,0,0,'Figure 1-1: A long caption',0,0),
                        (0,0,0,0,'continued here................ 1-2',1,0)]
            def get_drawings(self):
                return []
        self.assertFalse(candidate_page(Page()))

    def test_figure_list_entry_without_leader(self):
        class Page:
            def get_text(self, kind):
                text = '\n'.join(['Figure 1-1: Long title 1-2',
                                  'Figure 1-2: A........ 1-3',
                                  'Figure 1-3: B........ 1-4',
                                  'Figure 1-4: C........ 1-5'])
                return [(0,0,0,0,text,0,0)]
            def get_drawings(self):
                return []
        self.assertFalse(candidate_page(Page()))

    def test_regions_use_body_and_preserve_raster_footnotes(self):
        class Page:
            rect=f.Rect(0,0,612,792)
            rotation=0
            def get_image_info(self):
                return [{'bbox':(93,89,569,383)}]
        info={'page_size':[612,792],'para_blocks':[{'type':'image','blocks':[
            {'type':'image_body','bbox':[93,91,569,364]},
            {'type':'image_caption','bbox':[50,600,580,650]}]}]}
        rect=image_regions(info,Page())[0]
        self.assertGreaterEqual(rect.y1,383)
        self.assertLess(rect.y1,400)
        Page.rotation=90
        with self.assertRaises(ValueError): image_regions(info,Page())

    def test_mixed_block_keeps_prose_and_heading(self):
        with workspace_temp() as directory, f.open() as doc:
            page=doc.new_page()
            page.insert_text((50,60),'Chapter Title',fontsize=16)
            page.insert_text((50,180),'Diagram label')
            page.insert_text((50,400),'Body paragraph survives')
            doc.set_toc([[1,'Chapter Title',1]])
            _,bookmarks=n.make_bookmarks(doc)
            report=n.DocumentReport(source='sample',output='sample',pages=1)
            parts=n.render_page(doc,page,Path(directory),bookmarks[1],set(),report,
                                figure_regions=[f.Rect(40,140,300,220)])
            text=n.join_markdown_parts(parts)
            self.assertIn('# Chapter Title',text)
            self.assertIn('Body paragraph survives',text)
            self.assertNotIn('Diagram label',text)
            self.assertIn('enhanced-01.png',text)

    def test_upload_prohibition(self):
        with workspace_temp() as directory, f.open() as doc:
            path=Path(directory)/'sample.pdf'
            doc.new_page().insert_text((50,50),'sample')
            doc.save(path)
            with self.assertRaises(n.ConversionError):
                n.convert_pdf(path,Path(directory),route='native',figure_mode='mineru')

    def test_completed_layout_is_reused_without_network(self):
        import hashlib,json
        with workspace_temp() as directory, f.open() as doc:
            page=doc.new_page()
            page.insert_text((50,50),'Figure 1-1: Example')
            self.assertTrue(candidate_page(page))
            path=Path(directory)/'input.pdf'
            doc.save(path)
            digest=hashlib.sha256(path.read_bytes()).hexdigest()
            key=hashlib.sha256((digest+'vlm'+'en'+'figures-v1').encode()).hexdigest()
            folder=Path(directory)/'cache'/key/'1'
            folder.mkdir(parents=True)
            info={'page_size':[595,842],'para_blocks':[{'type':'image','blocks':[
                {'type':'image_body','bbox':[50,70,300,200]}]}]}
            (folder/'layout.json').write_text(json.dumps({'pdf_info':[info]}))
            with patch('figure_enrichment.MineruApi',side_effect=AssertionError('network forbidden')):
                result=discover_regions(path,doc,Path(directory)/'cache',None)
            self.assertEqual(len(result[1]),1)

if __name__=='__main__': unittest.main()
