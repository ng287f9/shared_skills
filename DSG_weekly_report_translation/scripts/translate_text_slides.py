#!/usr/bin/env python
"""
Phase 1: Translate text-based slides (TOC, Highlights/Lowlights, Status)
to bilingual EN+CN. Slides 5-7 (table/screenshot slides) are left untouched.

Usage:
    python translate_text_slides.py <input.pptx> <output.pptx>
"""

from lxml import etree
import sys, os, zipfile, shutil
sys.stdout.reconfigure(encoding='utf-8')

A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'

def tag(e): return e.tag.split('}')[-1] if '}' in e.tag else e.tag

def para_text(p):
    return ''.join(t.text or '' for t in p.iter(A + 't'))

def make_cn_run(text, sz='1400', color='5F6F82'):
    r = etree.Element(A + 'r')
    rPr = etree.SubElement(r, A + 'rPr')
    rPr.set('lang', 'zh-CN'); rPr.set('altLang', 'en-US')
    rPr.set('sz', sz); rPr.set('dirty', '0')
    sf = etree.SubElement(rPr, A + 'solidFill')
    etree.SubElement(sf, A + 'srgbClr').set('val', color)
    ea = etree.SubElement(rPr, A + 'ea')
    ea.set('typeface', 'Arial'); ea.set('pitchFamily', '34'); ea.set('charset', '0')
    etree.SubElement(r, A + 't').text = text
    return r

def make_cn_para(text, sz='1400', marL='358775', color='808080'):
    para = etree.Element(A + 'p')
    pPr = etree.SubElement(para, A + 'pPr')
    pPr.set('marL', marL); pPr.set('indent', '0')
    for tn in ['spcBef', 'spcAft']:
        e = etree.SubElement(pPr, A + tn)
        etree.SubElement(e, A + 'spcPts').set('val', '0')
    etree.SubElement(pPr, A + 'buNone')
    r = etree.SubElement(para, A + 'r')
    rPr = etree.SubElement(r, A + 'rPr')
    rPr.set('lang', 'zh-CN'); rPr.set('altLang', 'en-US')
    rPr.set('sz', sz); rPr.set('dirty', '0')
    sf = etree.SubElement(rPr, A + 'solidFill')
    etree.SubElement(sf, A + 'srgbClr').set('val', color)
    ea = etree.SubElement(rPr, A + 'ea')
    ea.set('typeface', 'Arial'); ea.set('pitchFamily', '34'); ea.set('charset', '0')
    etree.SubElement(r, A + 't').text = text
    return para

def find_content_body(root):
    best_body, best_n = None, 0
    for elem in root.iter():
        if tag(elem) == 'txBody':
            n = len([c for c in elem if tag(c) == 'p'])
            if n > best_n: best_n, best_body = n, elem
    return best_body

def process(input_pptx, output_pptx):
    # Unpack
    work = os.path.join(os.path.dirname(output_pptx), '_translate_work')
    if os.path.exists(work):
        shutil.rmtree(work)
    with zipfile.ZipFile(input_pptx, 'r') as z:
        z.extractall(work)

    slides_dir = os.path.join(work, 'ppt', 'slides')

    # ============================================================
    # SLIDE 2: TOC - Bucket B inline
    # ============================================================
    print('Processing Slide 2 (TOC)...')
    tree2 = etree.parse(os.path.join(slides_dir, 'slide2.xml'))
    body2 = find_content_body(tree2.getroot())
    toc_cn = {
        'Weekly Highlights / Lowlights': ' 本周亮点 / 风险及问题',
        'Status SN1002 / SN1003 / SN1004': ' SN1002 / SN1003 / SN1004状态',
        'Flight test program progress': ' 飞行测试项目进展',
        'Short term flight test plan': ' 短期飞行测试计划',
    }
    for p in [c for c in body2 if tag(c) == 'p']:
        pt = para_text(p).strip()
        if pt in toc_cn:
            p.append(make_cn_run(toc_cn[pt], sz='1600'))
    tree2.write(os.path.join(slides_dir, 'slide2.xml'), xml_declaration=True, encoding='UTF-8', standalone=True)

    # ============================================================
    # SLIDE 3: Highlights/Lowlights - Bucket B + C
    # ============================================================
    print('Processing Slide 3 (Highlights/Lowlights)...')
    tree3 = etree.parse(os.path.join(slides_dir, 'slide3.xml'))
    body3 = find_content_body(tree3.getroot())
    paras3 = [c for c in body3 if tag(c) == 'p']

    # Section headers -> Bucket B inline
    for p in paras3:
        pt = para_text(p).strip()
        if pt.startswith('Highlights') and len(pt) < 20:
            p.append(make_cn_run('亮点', sz='1400'))
        elif pt.startswith('Lowlights') and len(pt) < 20:
            p.append(make_cn_run('不足', sz='1400'))

    # Bullet translations (index, marL, cn_text)
    # marL = English paragraph's marL + indent (text start position)
    # Highlights bullets: marL=714375, indent=-355600 -> text at 358775
    # Lowlight bullets: marL=646112, indent=-285750 -> text at 360362
    bullets = [
        (1, '358775', '内容冻结飞行已成功完成，报告已按时提交给霍尼韦尔（HON）'),
        (2, '358775', '借调至Grob的人员安排已进入最后阶段，工作人员可即刻上岗'),
        (3, '358775', '出口管制律师的推进方向已明确，预计很快将作出选择'),
        (4, '358775', '变更控制委员会（CCB）因兼职外部承包商到位已恢复运作，已批准并发布多项MOR'),
        (7, '360362', '在霍尼韦尔软件中发现了一些缺陷，已在内容冻结报告中说明'),
        (8, '360362', '霍尼韦尔将对此进行研究（属预期情况，属正常现象）'),
        (9, '360362', '项目经理拒绝了录用通知，需重新启动招聘流程'),
        (10, '360362', '铰刀的采购申请仍在审批过程中，发动机地面试车时垂直尾翼未最终固定，可能引发振动'),
    ]
    for idx, marL, text in reversed(bullets):
        cn_p = make_cn_para(text, sz='1400', marL=marL)
        body3.insert(list(body3).index(paras3[idx]) + 1, cn_p)

    # Footer translation
    for elem in tree3.getroot().iter():
        if tag(elem) == 't' and elem.text and 'Program risks' in elem.text:
            elem.getparent().append(make_cn_run(' 项目风险详见项目风险文件（因尚无项目经理，暂未启用）', sz='1200'))

    # Tighten spacing
    for p in body3:
        if tag(p) != 'p': continue
        pPr = p.find(A + 'pPr')
        if pPr is None: continue
        lnSpc = pPr.find(A + 'lnSpc')
        if lnSpc is not None and lnSpc.find(A + 'spcPct') is not None:
            cur = int(lnSpc.find(A + 'spcPct').get('val', '100000'))
            if cur > 100000: lnSpc.find(A + 'spcPct').set('val', str(max(100000, cur - 40000)))
        for tn in ['spcBef', 'spcAft']:
            e = pPr.find(A + tn)
            if e is not None and e.find(A + 'spcPts') is not None:
                cur = int(e.find(A + 'spcPts').get('val', '0'))
                if cur > 100: e.find(A + 'spcPts').set('val', str(max(0, cur - 300)))

    tree3.write(os.path.join(slides_dir, 'slide3.xml'), xml_declaration=True, encoding='UTF-8', standalone=True)

    # ============================================================
    # SLIDE 4: Status - Bucket B + C, remove empties, tighten
    # ============================================================
    print('Processing Slide 4 (Status)...')
    tree4 = etree.parse(os.path.join(slides_dir, 'slide4.xml'))
    body4 = find_content_body(tree4.getroot())
    paras4 = [c for c in body4 if tag(c) == 'p']

    # Remove empty paragraphs
    for p in list(paras4):
        if not para_text(p).strip():
            body4.remove(p)

    # Refresh and translate
    paras4 = [c for c in body4 if tag(c) == 'p']
    bullets4 = [
        (1, '806450', '与飞机代持的可信公司的合同已就绪，仍待管理层决策，将在股东会议中汇报', '1200'),
        (2, '806450', '若此问题不解决，SN1002将保持停飞状态', '1200'),
        (3, '806450', '由于目前没有有效的（！）所有人，飞机注册面临极高风险', '1200'),
        (4, '806450', '一旦注册被取消，重新恢复运行将需要极大的工作量和极高的成本', '1200'),
        (7, '806450', '由于飞机因ADAU维修而停飞，年检工作正在进行中（优先级低）', '1200'),
        (8, '806450', 'ADAU仍未修复', '1200'),
        (9, '806450', '已成功完成地面试车，避免发动机进入封存状态', '1200'),
        (12, '806450', '已成功完成内容冻结试飞，因飞行员可用性导致的时间延误已追回', '1200'),
    ]
    for idx, marL, text, sz in reversed(bullets4):
        cn_p = make_cn_para(text, sz=sz, marL=marL)
        body4.insert(list(body4).index(paras4[idx]) + 1, cn_p)

    # Callout boxes
    for elem in tree4.getroot().iter():
        if tag(elem) == 'cNvPr':
            name = elem.get('name', '')
            if name in ('Rectangle 7', 'Rectangle 4'):
                txBody = None
                sp = elem.getparent().getparent()
                for c in sp:
                    if tag(c) == 'txBody': txBody = c; break
                if txBody is None: continue
                rps = [c for c in txBody if tag(c) == 'p']
                if name == 'Rectangle 7':
                    if len(rps) > 0: rps[0].append(make_cn_run(' 无变化', sz='1000'))
                    if len(rps) > 1: rps[1].append(make_cn_run(' 需在转场飞行前解决', sz='1000'))
                    if len(rps) > 2: rps[2].append(make_cn_run(' 即将召开的董事会议题', sz='1000'))
                elif name == 'Rectangle 4':
                    if len(rps) > 1: rps[1].append(make_cn_run(' 无变化，因为ADAU正在维修', sz='1000'))

    # Tighten spacing
    for p in body4:
        if tag(p) != 'p': continue
        pPr = p.find(A + 'pPr')
        if pPr is None: continue
        lnSpc = pPr.find(A + 'lnSpc')
        if lnSpc is not None and lnSpc.find(A + 'spcPct') is not None:
            cur = int(lnSpc.find(A + 'spcPct').get('val', '100000'))
            if cur > 100000: lnSpc.find(A + 'spcPct').set('val', str(max(100000, cur - 40000)))
        for tn in ['spcBef', 'spcAft']:
            e = pPr.find(A + tn)
            if e is not None and e.find(A + 'spcPts') is not None:
                cur = int(e.find(A + 'spcPts').get('val', '0'))
                if cur > 100: e.find(A + 'spcPts').set('val', str(max(0, cur - 300)))

    tree4.write(os.path.join(slides_dir, 'slide4.xml'), xml_declaration=True, encoding='UTF-8', standalone=True)

    # ============================================================
    # Repack
    # ============================================================
    print('Repacking...')
    if os.path.exists(output_pptx):
        os.remove(output_pptx)
    with zipfile.ZipFile(output_pptx, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(work):
            for f in files:
                zf.write(os.path.join(root, f), os.path.relpath(os.path.join(root, f), work))

    shutil.rmtree(work)
    print(f'Done. Output: {output_pptx}')
    print('Slides processed: 2 (TOC), 3 (Highlights/Lowlights), 4 (Status)')
    print('Slides left untouched: 1 (Cover), 5-7 (table/screenshot slides)')

if __name__ == '__main__':
    if len(sys.argv) < 3:
        print('Usage: python translate_text_slides.py <input.pptx> <output.pptx>')
        sys.exit(1)
    process(sys.argv[1], sys.argv[2])
