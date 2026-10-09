# -*- coding: utf-8 -*-
"""
百度OCR工具：支持图片识别、PDF识别（API直传）、PDF逐页渲染OCR（大文件）
用法：
  python baidu_ocr.py image <图片路径> [--output <输出文件>]
  python baidu_ocr.py pdf <PDF路径> [--pages <页码范围>] [--output <输出文件>]
  python baidu_ocr.py pdf_render <PDF路径> [--pages <页码范围>] [--output <输出文件>] [--concurrency 4] [--cache-dir <缓存目录>]
  
页码范围示例：1-3,5,7-10（默认全部页码）

新功能 pdf_render 模式（适用于大容量扫描版PDF）：
  - 用 PyMuPDF 逐页渲染为 PNG 图片
  - 逐页（或并发）调用百度高精度图片OCR API
  - 支持断点续跑（每页结果缓存到 --cache-dir）
  - 支持 --concurrency 控制并发数（默认4）
"""

import sys
import os
import json
import base64
import requests
import time
import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed

# 百度OCR API凭证（优先从环境变量读取，未设置时使用默认值）
API_KEY = os.environ.get('BAIDU_OCR_API_KEY', '<REDACTED_SECRET>')
SECRET_KEY = os.environ.get('BAIDU_OCR_SECRET_KEY', '<REDACTED_SECRET>')

TOKEN_URL = 'https://aip.baidubce.com/oauth/2.0/token'
ACCURATE_URL = 'https://aip.baidubce.com/rest/2.0/ocr/v1/accurate_basic'
PDF_URL = 'https://aip.baidubce.com/rest/2.0/ocr/v1/doc_analysis'

# 全局token缓存
_token = None


def get_access_token():
    """获取百度API access_token"""
    global _token
    if _token:
        return _token
    params = {
        'grant_type': 'client_credentials',
        'client_id': API_KEY,
        'client_secret': SECRET_KEY,
    }
    resp = requests.get(TOKEN_URL, params=params)
    resp.raise_for_status()
    result = resp.json()
    _token = result['access_token']
    return _token


def ocr_image_bytes(img_bytes, max_retry=3):
    """
    高精度图片OCR识别（传入字节流）
    返回: list of dict, 每项 {'words': str, 'location': {...}}
    """
    token = get_access_token()
    url = ACCURATE_URL + '?access_token=' + token
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    
    for attempt in range(max_retry):
        try:
            img_base64 = base64.b64encode(img_bytes).decode('utf-8')
            data = {
                'image': img_base64,
                'detect_direction': 'true',
                'paragraph': 'false',
            }
            resp = requests.post(url, headers=headers, data=data, timeout=30)
            resp.raise_for_status()
            result = resp.json()
            if 'error_code' in result:
                # QPS超限，等待后重试
                if result.get('error_code') == 18:
                    time.sleep(1 * (attempt + 1))
                    continue
                print('[ERROR] Baidu OCR: ' + str(result), file=sys.stderr)
                return []
            return result.get('words_result', [])
        except Exception as e:
            if attempt == max_retry - 1:
                print('[WARN] OCR failed after retry: ' + str(e), file=sys.stderr)
                return []
            time.sleep(1 * (attempt + 1))
    return []


def ocr_image_path(image_path):
    """高精度图片OCR识别（传入文件路径）"""
    if not os.path.exists(image_path):
        print('[ERROR] file not found: ' + image_path, file=sys.stderr)
        sys.exit(1)
    with open(image_path, 'rb') as f:
        img_bytes = f.read()
    return ocr_image_bytes(img_bytes)


def ocr_pdf_api(pdf_path, page_range=None):
    """
    百度PDF OCR API（整文件上传，适合小文件）
    注意：大文件（>20MB 或 >3000页）可能失败，建议用 pdf_render 模式
    """
    if not os.path.exists(pdf_path):
        print('[ERROR] file not found: ' + pdf_path, file=sys.stderr)
        sys.exit(1)

    with open(pdf_path, 'rb') as f:
        pdf_base64 = base64.b64encode(f.read()).decode('utf-8')

    token = get_access_token()
    url = PDF_URL + '?access_token=' + token
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}

    all_results = {}
    batch_size = 50
    start_page = 1

    while True:
        if page_range:
            batch_pages = [p for p in page_range if start_page <= p < start_page + batch_size]
            if not batch_pages:
                break
            pdf_num = str(batch_pages[0])
        else:
            pdf_num = str(start_page)

        data = {
            'pdf_file': pdf_base64,
            'pdf_file_num': pdf_num,
        }

        try:
            resp = requests.post(url, headers=headers, data=data, timeout=120)
            resp.raise_for_status()
            result = resp.json()
        except Exception as e:
            print('[ERROR] Baidu PDF OCR request failed: ' + str(e), file=sys.stderr)
            break

        if 'error_code' in result:
            print('[ERROR] Baidu PDF OCR: ' + str(result), file=sys.stderr)
            break

        results = result.get('words_result', [])
        for item in results:
            page_num = item.get('page_num', start_page)
            all_results[page_num] = item

        if result.get('is_end', True):
            break
        next_start = result.get('next_page', start_page + batch_size)
        if next_start <= start_page:
            break
        start_page = next_start

    return all_results


def render_pages_sequential(pdf_path, page_nums, dpi=200):
    """
    一次性打开PDF，顺序渲染多页为PNG字节流（避免并发打开大文件）
    返回: dict {page_num: png_bytes}
    """
    import fitz
    result = {}
    try:
        doc = fitz.open(pdf_path)
        mat = fitz.Matrix(dpi / 72, dpi / 72)
        for p in page_nums:
            page = doc[p - 1]
            pix = page.get_pixmap(matrix=mat)
            result[p] = pix.tobytes('png')
        doc.close()
    except Exception as e:
        print('[ERROR] Render pages failed: ' + str(e), file=sys.stderr)
    return result


def ocr_pdf_render(pdf_path, page_range=None, concurrency=4, cache_dir=None, dpi=200):
    """
    PDF逐页渲染 + 百度图片OCR（适用于大容量扫描版PDF）
    
    流程：
      1. 单次打开PDF，按批次顺序渲染为PNG（避免并发打开大文件的IO瓶颈）
      2. 并发调用百度高精度图片OCR API
      3. 每页结果缓存到 cache_dir（断点续跑）
    
    返回: dict {page_num: {'words_result': [{'words': str}]}}
    """
    import fitz
    
    if not os.path.exists(pdf_path):
        print('[ERROR] file not found: ' + pdf_path, file=sys.stderr)
        sys.exit(1)

    doc = fitz.open(pdf_path)
    total_pages = doc.page_count
    doc.close()

    # 确定要处理的页码
    if page_range:
        pages_to_process = [p for p in page_range if 1 <= p <= total_pages]
    else:
        pages_to_process = list(range(1, total_pages + 1))

    print('[INFO] PDF total pages: %d, to process: %d' % (total_pages, len(pages_to_process)), file=sys.stderr)
    print('[INFO] DPI: %d, concurrency: %d' % (dpi, concurrency), file=sys.stderr)

    # 缓存目录
    if cache_dir is None:
        cache_dir = pdf_path + '_ocr_cache'
    cache_dir = os.path.abspath(cache_dir)
    if not os.path.exists(cache_dir):
        os.makedirs(cache_dir)
        print('[INFO] Cache dir created: ' + cache_dir, file=sys.stderr)

    # 检查缓存，跳过已完成的页
    def cache_path(page_num):
        return os.path.join(cache_dir, 'page_%04d.json' % page_num)

    all_results = {}
    pages_need = []
    for p in pages_to_process:
        cp = cache_path(p)
        if os.path.exists(cp):
            try:
                with open(cp, 'r', encoding='utf-8') as f:
                    all_results[p] = json.load(f)
                continue
            except Exception:
                pass
        pages_need.append(p)

    print('[INFO] Cache hit: %d pages, need process: %d pages' % (len(all_results), len(pages_need)), file=sys.stderr)

    if not pages_need:
        print('[INFO] All pages cached, skip processing.', file=sys.stderr)
        return all_results

    # 分批处理：先顺序渲染（避免并发打开大文件），再并发OCR
    render_batch_size = concurrency
    total_to_process = len(pages_need)
    completed = 0

    for batch_start in range(0, total_to_process, render_batch_size):
        batch_pages = pages_need[batch_start:batch_start + render_batch_size]
        
        # 步骤1：顺序渲染该批次的页面（单次打开PDF，避免IO瓶颈）
        print('[INFO] Rendering pages %s ...' % str(batch_pages), file=sys.stderr)
        rendered = render_pages_sequential(pdf_path, batch_pages, dpi=dpi)
        
        if len(rendered) != len(batch_pages):
            missing = set(batch_pages) - set(rendered.keys())
            print('[WARN] %d pages failed to render: %s' % (len(missing), missing), file=sys.stderr)
        
        # 步骤2：并发OCR该批次的已渲染页面
        def process_one_page(page_num):
            cp = cache_path(page_num)
            try:
                png_bytes = rendered[page_num]
                words_result = ocr_image_bytes(png_bytes)
                cache_data = {
                    'page_num': page_num,
                    'words_result': words_result,
                }
                with open(cp, 'w', encoding='utf-8') as f:
                    json.dump(cache_data, f, ensure_ascii=False, indent=2)
                return page_num, cache_data, None
            except Exception as e:
                return page_num, None, str(e)

        if rendered:
            with ThreadPoolExecutor(max_workers=concurrency) as executor:
                futures = {executor.submit(process_one_page, p): p for p in rendered.keys()}
                for future in as_completed(futures):
                    page_num, result, error = future.result()
                    completed += 1
                    if result:
                        all_results[page_num] = result
                    else:
                        print('[WARN] Page %d failed: %s' % (page_num, error), file=sys.stderr)
        
        batch_done = batch_start + len(batch_pages)
        print('[INFO] Progress: %d/%d pages done' % (min(batch_done, total_to_process), total_to_process), file=sys.stderr)

    print('[INFO] All done. Total pages processed: %d' % len(all_results), file=sys.stderr)
    return all_results


def format_ocr_image(results):
    """格式化图片OCR结果为可读文本"""
    lines = []
    for item in results:
        text = item.get('words', '').strip()
        if text:
            lines.append(text)
    return '\n'.join(lines)


def format_ocr_pdf(results):
    """格式化PDF OCR结果为可读文本（API直传模式）"""
    all_text = []
    for page_num in sorted(results.keys()):
        page = results[page_num]
        all_text.append('========== 第 %d 页 ==========' % page_num)
        for block in page.get('words_result', []):
            words = block.get('words', '').strip()
            if words:
                all_text.append(words)
        all_text.append('')
    return '\n'.join(all_text)


def format_ocr_pdf_render(results):
    """格式化PDF逐页渲染OCR结果为可读文本"""
    all_text = []
    for page_num in sorted(results.keys()):
        page = results[page_num]
        all_text.append('========== 第 %d 页 ==========' % page_num)
        for item in page.get('words_result', []):
            words = item.get('words', '').strip()
            if words:
                all_text.append(words)
        all_text.append('')
    return '\n'.join(all_text)


def parse_page_range(range_str):
    """解析页码范围字符串，如 '1-3,5,7-10' -> [1,2,3,5,7,8,9,10]"""
    pages = []
    if not range_str:
        return None
    for part in range_str.split(','):
        part = part.strip()
        if '-' in part:
            a, b = part.split('-', 1)
            pages.extend(range(int(a), int(b) + 1))
        else:
            pages.append(int(part))
    return sorted(set(pages))


def main():
    parser = argparse.ArgumentParser(
        description='百度OCR工具：图片识别 / PDF识别 / PDF逐页渲染OCR（大文件）'
    )
    subparsers = parser.add_subparsers(dest='mode', help='识别模式')

    # image 模式
    sp_image = subparsers.add_parser('image', help='图片OCR识别')
    sp_image.add_argument('file', help='图片路径')
    sp_image.add_argument('--output', help='输出文件路径（不指定则打印到stdout）')

    # pdf 模式（API直传，适合小文件）
    sp_pdf = subparsers.add_parser('pdf', help='PDF OCR识别（API直传，适合小文件）')
    sp_pdf.add_argument('file', help='PDF文件路径')
    sp_pdf.add_argument('--pages', help='页码范围，如 1-3,5,7-10')
    sp_pdf.add_argument('--output', help='输出文件路径')

    # pdf_render 模式（逐页渲染，适合大文件）
    sp_render = subparsers.add_parser('pdf_render', help='PDF逐页渲染OCR（适合大容量扫描版PDF）')
    sp_render.add_argument('file', help='PDF文件路径')
    sp_render.add_argument('--pages', help='页码范围，如 1-3,5,7-10')
    sp_render.add_argument('--output', help='输出文件路径')
    sp_render.add_argument('--concurrency', type=int, default=4, help='并发数（默认4）')
    sp_render.add_argument('--cache-dir', help='缓存目录（默认 PDF路径_ocr_cache）')
    sp_render.add_argument('--dpi', type=int, default=200, help='渲染DPI（默认200，越高越清晰但越慢）')

    args = parser.parse_args()

    if not args.mode:
        parser.print_help()
        sys.exit(1)

    if args.mode == 'image':
        print('[INFO] Recognizing image: ' + args.file, file=sys.stderr)
        results = ocr_image_path(args.file)
        text = format_ocr_image(results)
        if args.output:
            with open(args.output, 'w', encoding='utf-8') as f:
                f.write(text)
            print('[OK] Output saved to: ' + args.output, file=sys.stderr)
        else:
            print(text)

    elif args.mode == 'pdf':
        page_range = parse_page_range(args.pages) if args.pages else None
        print('[INFO] Recognizing PDF (API mode): ' + args.file, file=sys.stderr)
        if page_range:
            print('[INFO] Pages: ' + str(page_range), file=sys.stderr)
        results = ocr_pdf_api(args.file, page_range)
        text = format_ocr_pdf(results)
        if args.output:
            with open(args.output, 'w', encoding='utf-8') as f:
                f.write(text)
            print('[OK] Output saved to: ' + args.output, file=sys.stderr)
        else:
            print(text)

    elif args.mode == 'pdf_render':
        page_range = parse_page_range(args.pages) if args.pages else None
        print('[INFO] Recognizing PDF (render mode): ' + args.file, file=sys.stderr)
        if page_range:
            print('[INFO] Pages: ' + str(page_range), file=sys.stderr)
        results = ocr_pdf_render(
            args.file,
            page_range=page_range,
            concurrency=args.concurrency,
            cache_dir=args.cache_dir,
            dpi=args.dpi,
        )
        text = format_ocr_pdf_render(results)
        if args.output:
            with open(args.output, 'w', encoding='utf-8') as f:
                f.write(text)
            print('[OK] Output saved to: ' + args.output, file=sys.stderr)
            # 同时输出JSON（含位置信息）
            json_path = args.output.replace('.txt', '.json').replace('.md', '.json')
            if json_path == args.output:
                json_path = args.output + '.json'
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(results, f, ensure_ascii=False, indent=2)
            print('[OK] JSON output saved to: ' + json_path, file=sys.stderr)
        else:
            print(text)


if __name__ == '__main__':
    main()
