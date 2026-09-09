"""Optional MinerU image-region discovery; native text remains authoritative."""
from pathlib import Path
import hashlib
import json
import logging
import math
import re
import pymupdf as fitz
from mineru_pdf_to_md import MineruApi, safe_extract_zip

LOGGER = logging.getLogger(__name__)
CAPTION = re.compile(r'^(?:Figure|Fig\.)\s+[A-Z0-9]+(?:[-.][A-Z0-9]+)*\s*[:.]', re.I)


def candidate_page(page):
    """Include captioned figures and substantial vector groups, omit TOC lines."""
    lines = [line.strip() for block in page.get_text('blocks')
             if len(block) > 6 and block[6] == 0
             for line in block[4].splitlines()]
    figure_list = sum(bool(re.search(r'\.{4,}\s*\d+(?:-\d+)?$', line))
                      for line in lines) >= 3
    for index, line in enumerate(lines):
        if not CAPTION.match(line):
            continue
        # A wrapped list-of-figures entry may put its dot leader on the
        # following line, even in a separate PDF text block.
        entry = [line]
        for following in lines[index+1:index+4]:
            if CAPTION.match(following):
                break
            entry.append(following)
        joined = ' '.join(entry)
        if figure_list and any(re.search(r'\s\d+(?:-\d+)?$', part) for part in entry):
            continue
        if not re.search(r'\.{4,}', joined):
            return True
    paths = page.get_drawings()
    substantial = [fitz.Rect(p['rect']) for p in paths
                   if fitz.Rect(p['rect']).width > 12 and fitz.Rect(p['rect']).height > 12]
    return len(substantial) >= 5


def image_regions(info, page):
    """Map only image_body boxes; captions can contain misclassified prose."""
    size = info.get('page_size', [])
    if len(size) != 2 or any(not math.isfinite(float(v)) or float(v) <= 0 for v in size):
        raise ValueError('MinerU page_size invalid')
    if page.rotation:
        raise ValueError('Rotated-page figure mapping is not yet supported')
    sx, sy = page.rect.width / size[0], page.rect.height / size[1]
    result = []
    blocks = info.get('para_blocks', info.get('preproc_blocks', []))
    previous_figure_table = None
    for block in blocks:
        kind = block.get('type')
        own_text = ' '.join(span.get('content', '') for line in block.get('lines', [])
                            for span in line.get('spans', []))
        if kind in {'text', 'title'} and CAPTION.match(own_text.strip()):
            previous_figure_table = block.get('bbox')
            continue
        captions = [child for child in block.get('blocks', [])
                    if child.get('type', '').endswith('_caption')]
        caption_text = ' '.join(span.get('content', '') for child in captions
                                for line in child.get('lines', [])
                                for span in line.get('spans', []))
        figure_table = (kind == 'table' and bool(CAPTION.match(caption_text.strip()))
                        and any(c.get('bbox', [0,0,0,0])[3] <= block['bbox'][1]
                                for c in captions))
        if (kind == 'table' and not captions and previous_figure_table is not None
                and 0 <= block['bbox'][1] - previous_figure_table[3] <= 50):
            figure_table = True
        previous_figure_table = block.get('bbox') if figure_table else None
        if kind not in {'image', 'chart'} and not figure_table:
            continue
        for body in block.get('blocks', []):
            if body.get('type') != kind + '_body':
                continue
            box = body.get('bbox', [])
            if len(box) != 4 or not all(math.isfinite(float(v)) for v in box):
                raise ValueError('MinerU image bbox invalid')
            rect = fitz.Rect(box[0]*sx, box[1]*sy, box[2]*sx, box[3]*sy) & page.rect
            if rect.is_empty:
                raise ValueError('MinerU image bbox empty')
            for raster in page.get_image_info():
                rr = fitz.Rect(raster['bbox'])
                if (rr & rect).get_area() / max(rr.get_area(), 1) > .65:
                    rect |= rr
            rect = (rect + (-2, -2, 2, 2)) & page.rect
            if not any((old & rect).get_area()/max(rect.get_area(),1) > .95 for old in result):
                result.append(rect)
    return refine_regions(result, blocks, page, sx, sy)


def refine_regions(regions, blocks, page, sx=1, sy=1):
    """Recover clipped native labels without expanding into ordinary prose.

    MinerU can omit diagram-side labels and split legends into separate bodies.
    Native text assigned to prose, headings, code or captions remains protected.
    """
    # A single embedded bitmap may have been split into several chart bodies.
    # Restore its complete extent once their combined coverage is substantial.
    for raster in page.get_image_info():
        rr = fitz.Rect(raster['bbox'])
        hits = [i for i,r in enumerate(regions) if (r & rr).get_area() > 0]
        coverage = sum((regions[i] & rr).get_area() for i in hits)
        if len(hits) > 1 and coverage / max(rr.get_area(),1) > .65:
            combined = fitz.Rect(rr)
            for i in hits:
                combined |= regions[i]
            regions = [r for i,r in enumerate(regions) if i not in hits] + [combined]
    protected = []
    for block in blocks:
        children = block.get('blocks', [])
        protected_parts = ([block] if block.get('type') in {'text','title','code'} else [])
        for child in children:
            if not child.get('type','').endswith('_caption'):
                continue
            text = ' '.join(s.get('content','') for line in child.get('lines',[])
                            for s in line.get('spans',[])).strip()
            # Short unnumbered captions can be diagram labels, e.g. a bit-field
            # brace explanation. Numbered captions and prose stay as native text.
            if re.match(r'^(?:Figure|Fig\.|Table)\s',text,re.I) or len(text)>100:
                protected_parts.append(child)
        for item in protected_parts:
            box = item.get('bbox', [])
            if len(box) == 4:
                protected.append(fitz.Rect(box[0]*sx,box[1]*sy,box[2]*sx,box[3]*sy))
    if hasattr(page, 'get_text'):
        lines = [line for block in page.get_text('dict')['blocks'] if block['type']==0
                 for line in block['lines']]
        for index, original in enumerate(regions):
            rect = fitz.Rect(original)
            for _ in range(3):
                before = tuple(rect)
                for line in lines:
                    box = fitz.Rect(line['bbox'])
                    if any((box & p).get_area()/max(box.get_area(),1) > .5 for p in protected):
                        continue
                    # Only add nearby unassigned labels, not distant columns.
                    if (box & (rect + (-12,-12,12,12))).is_empty:
                        continue
                    rect |= box
                if tuple(rect) == before:
                    break
            regions[index] = (rect + (-1,-1,1,1)) & page.rect
    merged = []
    for rect in regions:
        for index, old in enumerate(merged):
            overlap = (old & rect).get_area()/max(min(old.get_area(),rect.get_area()),1)
            near = not (old & (rect + (-8,-8,8,8))).is_empty
            bridge = old | rect
            crosses_prose = any((p & bridge).get_area()/max(p.get_area(),1) > .5
                                and (p & old).is_empty and (p & rect).is_empty
                                for p in protected)
            if overlap > .9 or (near and not crosses_prose):
                merged[index] = bridge
                break
        else:
            merged.append(rect)
    return sorted(merged, key=lambda r:(r.y0,r.x0))


def discover_regions(source, doc, cache_root, token, model='vlm', language='en'):
    """One-page tasks bound upload size and retain completed tasks across retries."""
    digest = hashlib.sha256()
    with source.open('rb') as handle:
        for chunk in iter(lambda: handle.read(1024*1024), b''):
            digest.update(chunk)
    key = hashlib.sha256((digest.hexdigest()+model+language+'figures-v1').encode()).hexdigest()
    root = cache_root / key
    api = None
    results = {}
    for page in doc:
        if not candidate_page(page):
            continue
        number = page.number + 1
        if page.rotation:
            raise ValueError(f'第 {number} 页旋转，尚不支持图片坐标映射')
        folder = root / str(number)
        folder.mkdir(parents=True, exist_ok=True)
        layout = folder / 'layout.json'
        if not layout.exists():
            if not token:
                raise ValueError('图片增强需要 MINERU_TOKEN')
            if api is None:
                api = MineruApi(token, model_version=model, language=language, is_ocr=True)
            sample = folder / 'page.pdf'
            if not sample.exists():
                with fitz.open() as selected:
                    selected.insert_pdf(doc, from_page=page.number, to_page=page.number)
                    selected.save(sample)
            if sample.stat().st_size >= 190*1024*1024:
                raise ValueError(f'第 {number} 页超过图片增强单页上传限制 190 MiB')
            state = folder / 'task.json'
            LOGGER.warning('图片增强：第 %d/%d 页，上传/查询 MinerU', number, len(doc))
            if state.exists():
                batch = json.loads(state.read_text(encoding='utf-8'))['batch_id']
            else:
                batch, url = api.create_upload_task(sample.name, f'figure-{number}')
                api.upload_file(sample, url)
                state.write_text(json.dumps({'batch_id':batch}), encoding='utf-8')
            entry = api.wait_for_batch(batch, file_name=f'page-{number}', data_id=f'figure-{number}', poll_interval=5, poll_timeout=600)
            archive = folder / 'result.zip'
            api.download_zip(entry['full_zip_url'], archive)
            extracted = folder / 'result'
            safe_extract_zip(archive, extracted)
            matches = list(extracted.rglob('layout.json'))
            if len(matches) != 1:
                raise ValueError(f'第 {number} 页未返回唯一 layout.json')
            data = json.loads(matches[0].read_text(encoding='utf-8'))
            # Validate before caching a normalized copy.
            if len(data.get('pdf_info', [])) != 1:
                raise ValueError('MinerU 单页结果页数不匹配')
            image_regions(data['pdf_info'][0], page)
            layout.write_text(json.dumps(data), encoding='utf-8')
        data = json.loads(layout.read_text(encoding='utf-8'))
        regions = image_regions(data['pdf_info'][0], page)
        if not regions:
            LOGGER.warning('图片增强：第 %d 页未识别到图片，保留本地结果，请核查', number)
        results[number] = regions
    return results
