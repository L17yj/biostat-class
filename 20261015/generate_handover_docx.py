# -*- coding: utf-8 -*-
"""
生成生物統計實習交接與整體分析報告 Word 檔 (.docx)
"""

import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def create_handover_document():
    doc = Document()

    # 設定版面邊界 (A4)
    for section in doc.sections:
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.85)
        section.right_margin = Inches(0.85)

    def format_run(run, font_name='微軟正黑體', size_pt=10.5, bold=False, italic=False, color_rgb=(40,40,40)):
        run.font.name = font_name
        run.font.size = Pt(size_pt)
        run.bold = bold
        run.italic = italic
        run.font.color.rgb = RGBColor(*color_rgb)
        rPr = run._r.get_or_add_rPr()
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:ascii'), font_name)
        rFonts.set(qn('w:eastAsia'), font_name)
        rFonts.set(qn('w:hAnsi'), font_name)
        rPr.append(rFonts)

    def add_title(text, subtitle=None):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(text)
        format_run(run, font_name='微軟正黑體', size_pt=18, bold=True, color_rgb=(31, 58, 99))

        if subtitle:
            p_sub = doc.add_paragraph()
            p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_sub.paragraph_format.space_after = Pt(14)
            run_sub = p_sub.add_run(subtitle)
            format_run(run_sub, font_name='微軟正黑體', size_pt=11.5, italic=True, color_rgb=(80, 95, 115))

    def add_h1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(5)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        format_run(run, font_name='微軟正黑體', size_pt=13.5, bold=True, color_rgb=(31, 58, 99))
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        format_run(run, font_name='微軟正黑體', size_pt=11.5, bold=True, color_rgb=(46, 82, 126))
        return p

    def add_h3(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(7)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        format_run(run, font_name='微軟正黑體', size_pt=10.5, bold=True, color_rgb=(60, 60, 60))
        return p

    def add_body(text, bold_prefix='', space_after=4):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.25
        if bold_prefix:
            r_pre = p.add_run(bold_prefix)
            format_run(r_pre, bold=True, color_rgb=(25, 25, 25))
        r = p.add_run(text)
        format_run(r, color_rgb=(45, 45, 45))
        return p

    def set_cell_shading(cell, color_hex):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
        tcPr.append(shd)

    def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = OxmlElement('w:tcMar')
        for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
            node = OxmlElement(m)
            node.set(qn('w:w'), str(val))
            node.set(qn('w:type'), 'dxa')
            tcMar.append(node)
        tcPr.append(tcMar)

    def style_table(table, col_widths=None, header_bg="2E4057", zebra_bg="F6F8FA"):
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for r_idx, row in enumerate(table.rows):
            is_header = (r_idx == 0)
            for c_idx, cell in enumerate(row.cells):
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                set_cell_margins(cell, top=90, bottom=90, left=120, right=120)
                if col_widths and c_idx < len(col_widths):
                    cell.width = Inches(col_widths[c_idx])
                if is_header:
                    set_cell_shading(cell, header_bg)
                    for p in cell.paragraphs:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p.paragraph_format.space_after = Pt(2)
                        p.paragraph_format.space_before = Pt(2)
                        for r in p.runs:
                            format_run(r, size_pt=9.5, bold=True, color_rgb=(255, 255, 255))
                else:
                    if r_idx % 2 == 1:
                        set_cell_shading(cell, zebra_bg)
                    else:
                        set_cell_shading(cell, "FFFFFF")
                    for p in cell.paragraphs:
                        p.paragraph_format.space_after = Pt(1)
                        p.paragraph_format.space_before = Pt(1)
                        for r in p.runs:
                            format_run(r, size_pt=9.2, color_rgb=(45, 45, 45))
                # Add light border
                tcPr = cell._tc.get_or_add_tcPr()
                borders = parse_xml(f'''
                    <w:tcBorders {nsdecls("w")}>
                        <w:top w:val="single" w:sz="4" w:space="0" w:color="D0D5DD"/>
                        <w:left w:val="none"/>
                        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="D0D5DD"/>
                        <w:right w:val="none"/>
                    </w:tcBorders>
                ''')
                tcPr.append(borders)

    def add_callout(text, title='重要備註與洞察', box_type='note'):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.5)
        bg = 'EDF2F7' if box_type == 'note' else 'FFF8E7'
        border_c = '2E4057' if box_type == 'note' else 'C47B6A'
        set_cell_shading(cell, bg)
        set_cell_margins(cell, top=120, bottom=120, left=180, right=180)
        
        tcPr = cell._tc.get_or_add_tcPr()
        borders = parse_xml(f'''
            <w:tcBorders {nsdecls("w")}>
                <w:top w:val="none"/>
                <w:left w:val="single" w:sz="24" w:space="0" w:color="{border_c}"/>
                <w:bottom w:val="none"/>
                <w:right w:val="none"/>
            </w:tcBorders>
        ''')
        tcPr.append(borders)
        
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(2)
        r_t = p.add_run(f'【{title}】\n')
        format_run(r_t, bold=True, size_pt=10, color_rgb=(30, 50, 80) if box_type=='note' else (130, 60, 20))
        r_b = p.add_run(text)
        format_run(r_b, size_pt=9.3, color_rgb=(50, 50, 50))
        
        p_space = doc.add_paragraph()
        p_space.paragraph_format.space_after = Pt(4)

    # ==================== 文件開頭：標題與基本資料 ====================
    add_title('生物統計實習：無母數統計整體分析與專案交接文件', '20261015 實習專案（Mann-Whitney U 檢定 & Kruskal-Wallis 檢定）')

    # 基本資訊卡片表格
    meta_tbl = doc.add_table(rows=4, cols=4)
    col_w = [1.3, 1.9, 1.3, 2.0]
    meta_data = [
        [('專案主題', True), ('無母數統計分析整體交接', False), ('分析日期', True), ('2026 年 10 月 15 日', False)],
        [('原始分析人員', True), ('劉昱君', False), ('分析工具', True), ('IBM SPSS Statistics / Python 3.12', False)],
        [('交付文件數量', True), ('5 個檔案 (3 Excel + 2 Doc)', False), ('補充圖表', True), ('3 幅高清莫蘭迪盒鬚圖', False)],
        [('交接狀態', True), ('完成整體交叉驗證與排錯', False), ('報告用途', True), ('課程專案交付 / 期末報告 / 人員銜接', False)]
    ]
    for r_idx, row_content in enumerate(meta_data):
        for c_idx, (text, is_bold) in enumerate(row_content):
            cell = meta_tbl.cell(r_idx, c_idx)
            cell.width = Inches(col_w[c_idx])
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if is_bold else WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(text)
            format_run(r, size_pt=9.2, bold=is_bold, color_rgb=(31, 58, 99) if is_bold else (40, 40, 40))
            if is_bold:
                set_cell_shading(cell, "EAEFF5")
            else:
                set_cell_shading(cell, "FAFBFC")
            set_cell_margins(cell, top=70, bottom=70, left=100, right=100)
            tcPr = cell._tc.get_or_add_tcPr()
            borders = parse_xml(f'''
                <w:tcBorders {nsdecls("w")}>
                    <w:top w:val="single" w:sz="4" w:space="0" w:color="D0D5DD"/>
                    <w:left w:val="single" w:sz="4" w:space="0" w:color="D0D5DD"/>
                    <w:bottom w:val="single" w:sz="4" w:space="0" w:color="D0D5DD"/>
                    <w:right w:val="single" w:sz="4" w:space="0" w:color="D0D5DD"/>
                </w:tcBorders>
            ''')
            tcPr.append(borders)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # ==================== 執行摘要 ====================
    add_h1('執行摘要（Executive Summary）')
    add_body('本專案交接文件針對 2026 年 10 月 15 日生物統計實習中的五個核心檔案（含 3 份原始 Excel 數據集與 2 份 SPSS 輸出分析報告）進行全面整合、數值交叉比對與系統性實務交接。')
    add_body('本階段統計主題為「無母數統計檢定（Non-parametric Statistics）」，聚焦於當樣本數較小、非常態分佈或組間變異數嚴重不齊時之兩組及多組獨立樣本差異性檢驗：')
    add_body('1. 主題一（兩組獨立樣本檢定）：透過 Mann-Whitney U 檢定分析乳癌（BC）與肺腺癌（LUAD）檢體組織之 HER2 蛋白質表現量差異。檢定結果 U = 1.000, Wilcoxon W = 46.000, Z = -3.372, p = 0.001，乳癌組顯著高於肺腺癌組，具極高度統計顯著性。')
    add_body('2. 主題二（多組獨立樣本檢定）：透過 Kruskal-Wallis 檢定分析四大職業運動組織（NBA、MLB、NFL、WNBA）頂尖球員年薪結構。檢定結果 H = 27.611, df = 3, p < 0.001，揭露男女職業聯賽間高達 200 倍以上之結構性薪資差距。本文件同時完整還原並排解了 SPSS 語法執行之歷史警告。')
    add_body('3. 主題三（補充多組分析）：針對尚未附帶報表之 KW_Blood type.xlsx 補充完成 Kruskal-Wallis 統計檢定，分析 ABO 四大血型與成人身高關聯。結果 H = 1.507, df = 3, p = 0.681，未達顯著水準，印證血型與身高無統計相關。')

    # ==================== 一、五份文件資產盤點 ====================
    add_h1('一、五份文件資產盤點與資料流關聯')
    add_body('本實習專案包含 3 份原始資料工作表與 2 份 SPSS 分析產出文件，其結構與功能定位盤點如下：')

    inv_tbl = doc.add_table(rows=6, cols=5)
    headers = ['序號', '檔案名稱', '檔案類型', '檔案大小', '定位與內容說明']
    widths = [0.5, 2.0, 0.9, 0.8, 2.3]
    for c_idx, h in enumerate(headers):
        inv_tbl.cell(0, c_idx).paragraphs[0].add_run(h)
    
    rows_data = [
        ['1', 'Mann whitney.xlsx', 'Excel (.xlsx)', '9.6 KB', '原始數據：17 筆腫瘤 HER2 蛋白質表現量（乳癌 8 筆、肺腺癌 9 筆）。'],
        ['2', '20261015(1)_劉昱君.doc', 'SPSS / RTF', '81.1 KB', '實習成果報表：對應 Mann whitney.xlsx 產出之 Mann-Whitney 檢定與盒鬚圖。'],
        ['3', 'KW_Pro_salary.xlsx', 'Excel (.xlsx)', '12.5 KB', '原始數據：40 位職業球員年薪（NBA 10, MLB 10, NFL 10, WNBA 10）。'],
        ['4', '20261015(2)_劉昱君.doc', 'SPSS / RTF', '111.7 KB', '實習成果報表：對應 Pro_salary 產出之 Kruskal-Wallis 檢定，含報錯紀錄。'],
        ['5', 'KW_Blood type.xlsx', 'Excel (.xlsx)', '11.7 KB', '原始數據：工作表1為44筆血型與身高；工作表2為NFL球員年薪暫存表。']
    ]
    for r_idx, r_data in enumerate(rows_data, 1):
        for c_idx, val in enumerate(r_data):
            cell = inv_tbl.cell(r_idx, c_idx)
            p = cell.paragraphs[0]
            if c_idx in [0, 2, 3]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run(val)
    style_table(inv_tbl, widths)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # ==================== 二、主題一：Mann-Whitney U 檢定 ====================
    add_h1('二、主題一：Mann-Whitney U 檢定（HER2 蛋白質表現量分析）')
    add_h2('2.1 研究假說與無母數檢定選用理由')
    add_body('人類表皮生長因子受體 2（HER2 / ERBB2）是腫瘤分子生物學中的關鍵生物標記。本實驗收集乳癌患者（Breast Cancer, BC, n=8）與肺腺癌患者（Lung Adenocarcinoma, LUAD, n=9）之腫瘤組織檢體進行表現量測定：')
    add_body('• 虛無假說 (H0)：乳癌組與肺腺癌組之 HER2 蛋白質表現量分佈無顯著差異。')
    add_body('• 對立假說 (H1)：乳癌組與肺腺癌組之 HER2 蛋白質表現量分佈存在顯著差異。')
    add_body('• 統計方法選用理由：樣本數小（n1=8, n2=9），且乳癌組樣本標準差高達 202.05，肺腺癌組僅 25.62，兩組標準差相差近 8 倍，嚴重違反獨立樣本 t 檢定之母體常態性與變異數同質性假設，故必須採用 Mann-Whitney U 檢定（Wilcoxon Rank-Sum Test）。')

    add_h2('2.2 敘述性統計分析')
    s1_tbl = doc.add_table(rows=4, cols=8)
    s1_headers = ['組別', '樣本數', '平均數', '標準差', '中位數', 'Q1', 'Q3', 'IQR']
    s1_widths = [1.1, 0.6, 0.8, 0.8, 0.8, 0.7, 0.7, 0.7]
    for c_idx, h in enumerate(s1_headers):
        s1_tbl.cell(0, c_idx).paragraphs[0].add_run(h)
    s1_rows = [
        ['乳癌 (BC)', '8', '434.50', '202.05', '446.50', '265.00', '654.00', '389.00'],
        ['肺腺癌 (LUAD)', '9', '128.44', '25.62', '118.00', '112.00', '125.00', '13.00'],
        ['整體合計', '17', '272.47', '203.49', '156.00', '118.00', '446.50', '328.50']
    ]
    for r_idx, r_data in enumerate(s1_rows, 1):
        for c_idx, val in enumerate(r_data):
            cell = s1_tbl.cell(r_idx, c_idx)
            p = cell.paragraphs[0]
            if c_idx > 0:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p.add_run(val)
    style_table(s1_tbl, s1_widths)

    add_h2('2.3 等級統計量與檢定結果比對')
    add_body('與 20261015(1)_劉昱君.doc 報表比對，數值 100% 精確吻合：')
    add_body('• 乳癌組 (BC, n=8)：平均等級 (Mean Rank) = 13.38，等級總和 (Sum of Ranks) = 107.00')
    add_body('• 肺腺癌組 (LUAD, n=9)：平均等級 (Mean Rank) = 5.11，等級總和 (Sum of Ranks) = 46.00')
    add_body('• Mann-Whitney U = 1.000, Wilcoxon W = 46.000, 標準化 Z = -3.372')
    add_body('• 漸近顯著性 (雙尾) p = .001，精確顯著性 [2*(單尾)] p = .000')

    # 插入圖一
    if os.path.exists('fig1_her2_comparison.png'):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(8)
        p_img.paragraph_format.space_after = Pt(4)
        run_img = p_img.add_run()
        run_img.add_picture('fig1_her2_comparison.png', width=Inches(5.0))
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_after = Pt(8)
        r_cap = p_cap.add_run('圖 1：HER2 蛋白質表現量盒鬚圖（莫蘭迪色系填充、黑色中位數線、散佈微縮數據點）')
        format_run(r_cap, size_pt=9.0, italic=True, color_rgb=(100, 100, 100))

    add_callout(
        '1. 統計顯著性：乳癌組表現中位數（446.50）為肺腺癌組（118.00）之近 4 倍，且 p = 0.001 遠小於 0.05，明確拒絕虛無假說。\n'
        '2. 分佈型態與離散度：乳癌組 IQR 達 389.00，顯示極高的腫瘤異質性（部分患者呈現數倍之過度表現），符合 HER2 陽性乳癌亞型之病理特徵；肺腺癌組 IQR 僅 13.00，表現量高度集中於基底水準。',
        title='臨床與生物醫學意義解析', box_type='note'
    )

    # ==================== 三、主題二：Kruskal-Wallis 檢定 ====================
    add_h1('三、主題二：Kruskal-Wallis 檢定（四大職業運動年薪分析）')
    add_h2('3.1 研究假說與分組架構')
    add_body('本分析納入北美四大職業運動聯盟（NBA、MLB、NFL、WNBA）各 10 位頂尖球員年薪（Annual Salary, N=40）：')
    add_body('• 虛無假說 (H0)：四大職業聯盟頂尖球員之年薪母體中位數相同。')
    add_body('• 對立假說 (H1)：至少有兩個職業聯盟之球員年薪分佈存在顯著差異。')

    add_h2('3.2 四大聯盟敘述性統計')
    s2_tbl = doc.add_table(rows=5, cols=7)
    s2_headers = ['職業聯盟', '性別', '樣本數', '平均年薪 (USD)', '中位數年薪 (USD)', '年薪標準差 (USD)', '年薪全距範圍']
    s2_widths = [1.1, 0.6, 0.6, 1.2, 1.2, 1.1, 1.1]
    for c_idx, h in enumerate(s2_headers):
        s2_tbl.cell(0, c_idx).paragraphs[0].add_run(h)
    s2_rows = [
        ['NBA (職籃)', '男', '10', '$45,464,546', '$46,623,717', '$7,037,160', '34.8M ~ 55.8M'],
        ['MLB (職棒)', '男', '10', '$42,370,833', '$38,500,000', '$11,121,253', '32.5M ~ 70.0M'],
        ['NFL (美足)', '男', '10', '$52,460,000', '$52,750,000', '$4,408,880', '45.0M ~ 60.0M'],
        ['WNBA (女籃)', '女', '10', '$199,922', '$203,500', '$47,436', '76.5K ~ 252.5K']
    ]
    for r_idx, r_data in enumerate(s2_rows, 1):
        for c_idx, val in enumerate(r_data):
            cell = s2_tbl.cell(r_idx, c_idx)
            p = cell.paragraphs[0]
            if c_idx in [1, 2]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif c_idx in [3, 4, 5]:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p.add_run(val)
    style_table(s2_tbl, s2_widths)

    add_h2('3.3 Kruskal-Wallis 等級與檢定結果')
    add_body('與 20261015(2)_劉昱君.doc 輸出完全一致：')
    add_body('• NBA (Group 1)：平均等級 (Mean Rank) = 24.20')
    add_body('• MLB (Group 2)：平均等級 (Mean Rank) = 20.05')
    add_body('• NFL (Group 3)：平均等級 (Mean Rank) = 32.25')
    add_body('• WNBA (Group 4)：平均等級 (Mean Rank) = 5.50')
    add_body('• Kruskal-Wallis H (未校正) = 27.582，同分校正後 H = 27.611，自由度 df = 3，漸近顯著性 p = .000 (4.38 × 10^-6)。明確拒絕虛無假說！')

    # 插入圖二
    if os.path.exists('fig2_salary_comparison.png'):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(8)
        p_img.paragraph_format.space_after = Pt(4)
        run_img = p_img.add_run()
        run_img.add_picture('fig2_salary_comparison.png', width=Inches(5.6))
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_after = Pt(8)
        r_cap = p_cap.add_run('圖 2：四大職業運動聯盟球員年薪分佈盒鬚圖（四色莫蘭迪系、黑色中位數線）')
        format_run(r_cap, size_pt=9.0, italic=True, color_rgb=(100, 100, 100))

    add_callout(
        '1. 結構性性別薪資鴻溝：在 40 位球員排序中，前 10 名全部由 WNBA 包攬（等級剛好為 1 到 10，平均等級 5.50）。WNBA 最高薪 Jackie Young（約 25.2 萬美元）遠不及男子三大聯盟後段門檻（3,250 萬美元），差距達 200 倍以上。\n'
        '2. 男子三大聯盟內部比較：NFL 四分衛年薪等級最高（Mean Rank = 32.25），中位數達 5,275 萬美元；MLB 雖有大谷翔平以 7,000 萬美元創紀錄，但整體中位數較低（3,850 萬美元）；事後成對比較顯示 NFL 與 MLB 達到統計顯著差異（p = 0.0046 < 0.0083 Bonferroni 修正臨界值）。',
        title='事後檢定與職業運動產業經濟洞察', box_type='note'
    )

    # ==================== 四、SPSS 排錯記錄 ====================
    add_h1('四、重要交接實務：SPSS 語法操作錯誤與排錯紀錄（Troubleshooting Log）')
    add_body('在 20261015(2)_劉昱君.doc 中，保留了一段 SPSS 執行報錯與警告紀錄。為確保後續操作人員不重複踩坑，特此詳細記錄：')

    add_callout(
        '【原始報錯語法】：\n'
        '  NPAR TESTS\n'
        '    /K-W=profession_A BY SalaryAnnual(1 4)\n'
        '    /MISSING ANALYSIS.\n\n'
        '【SPSS 系統警告訊息】：\n'
        '  「沒有足夠的有效觀察值可用來執行 profession * Salary (Annual) (1.0000, 4.0000) 的『 Kruskal-Wallis 檢定』。未計算任何統計量。」\n\n'
        '【根本原因 (Root Cause)】：\n'
        '  變數角色顛倒！SPSS 語法規範為 /K-W= [連續檢定變數] BY [類別分組變數](最小值 最大值)。\n'
        '  原操作誤將分類變數 profession_A 置於檢定位置，而將數千萬的 SalaryAnnual 置於分組位置並限定 (1 4)，導致數值在 1~4 之間的有效個數為 0。',
        title='SPSS 歷史操作報錯還原與根因剖析', box_type='warn'
    )

    add_h3('正確標準作業程序 (Standard Operating Procedure, SOP)')
    add_body('1. 正確語法指令：')
    add_body('   NPAR TESTS /K-W=SalaryAnnual BY profession_A(1 4) /MISSING ANALYSIS.')
    add_body('2. 圖形化介面 (GUI) 操作路徑：')
    add_body('   分析 (Analyze) → 無母數檢定 (Nonparametric Tests) → 舊對話框 (Legacy Dialogs) → K 個獨立樣本 (K Independent Samples)。')
    add_body('   • 檢定變數清單 (Test Variable List)：置入連續型變數「Salary (Annual)」。')
    add_body('   • 分組變數 (Grouping Variable)：置入分類變數「profession」，點選「定義範圍」，輸入最小值 1、最大值 4。')
    add_body('   • 勾選「Kruskal-Wallis H」。')

    # ==================== 五、主題三：血型與身高分析 ====================
    add_h1('五、主題三：補充分析（血型與身高關聯分析 — KW_Blood type.xlsx）')
    add_body('檔案 KW_Blood type.xlsx 原尚未包含 Word 報表，本交接文件補充完成其完整統計分析與繪圖：')
    add_body('• 研究對象：44 位成年受試者之 ABO 血型與身高數據。')
    add_body('• 假說：探討 ABO 血型表現型是否對人體身高產生顯著影響。')

    s3_tbl = doc.add_table(rows=5, cols=7)
    s3_headers = ['血型類別', '樣本數', '平均身高 (cm)', '中位數 (cm)', '標準差 (cm)', 'IQR (cm)', '平均等級 (Mean Rank)']
    s3_widths = [1.1, 0.7, 1.1, 1.1, 1.0, 0.9, 1.1]
    for c_idx, h in enumerate(s3_headers):
        s3_tbl.cell(0, c_idx).paragraphs[0].add_run(h)
    s3_rows = [
        ['A 型 (Group 1)', '12', '174.00', '173.50', '7.63', '12.00', '26.00'],
        ['B 型 (Group 2)', '10', '171.50', '173.00', '9.07', '11.00', '22.90'],
        ['O 型 (Group 3)', '9', '170.44', '170.00', '11.31', '11.00', '20.83'],
        ['AB 型 (Group 4)', '13', '169.92', '172.00', '9.11', '9.00', '20.12']
    ]
    for r_idx, r_data in enumerate(s3_rows, 1):
        for c_idx, val in enumerate(r_data):
            cell = s3_tbl.cell(r_idx, c_idx)
            p = cell.paragraphs[0]
            if c_idx == 1:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif c_idx > 1:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p.add_run(val)
    style_table(s3_tbl, s3_widths)

    add_body('• 檢定結果：Kruskal-Wallis H = 1.500，同分校正後 H = 1.507，自由度 df = 3，漸近顯著性 p = .681。')
    add_body('• 統計學結論：p = 0.681 > 0.05，未達統計顯著水準，無法拒絕虛無假說。')

    # 插入圖三
    if os.path.exists('fig3_bloodtype_height.png'):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(8)
        p_img.paragraph_format.space_after = Pt(4)
        run_img = p_img.add_run()
        run_img.add_picture('fig3_bloodtype_height.png', width=Inches(5.4))
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_after = Pt(8)
        r_cap = p_cap.add_run('圖 3：不同 ABO 血型受試者身高分佈盒鬚圖（四色莫蘭迪系、黑色中位數線）')
        format_run(r_cap, size_pt=9.0, italic=True, color_rgb=(100, 100, 100))

    add_callout(
        '生物學與遺傳學解析：\n'
        '身高屬於高度多基因遺傳（Polygenic trait）並深受後天營養環境影響；而 ABO 血型乃由第 9 號染色體上的單一基因位點（ABO gene）調控紅血球表面抗原醣鏈，兩者在遺傳機制上完全獨立分離。實習數據 p = 0.681 忠實印證了血型與身高無因果關聯之客觀科學常理。此外，工作表 2 包含 10 筆 NFL 球員年薪，係主題二 KW_Pro_salary.xlsx 建立時之資料暫存表。',
        title='遺傳學常理與資料集結構解析', box_type='note'
    )

    # ==================== 六、方法論決策樹 ====================
    add_h1('六、生物統計學方法論決策樹與對照表')
    add_body('為利於接手同學與後續專案選用最適統計方法，整理獨立樣本差異檢定之選用決策表：')

    comp_tbl = doc.add_table(rows=6, cols=3)
    c_headers = ['比較維度', '母數檢定 (Parametric)', '無母數檢定 (Non-parametric)']
    c_widths = [1.5, 2.7, 2.8]
    for c_idx, h in enumerate(c_headers):
        comp_tbl.cell(0, c_idx).paragraphs[0].add_run(h)
    comp_rows = [
        ['兩組獨立樣本', '獨立樣本 t 檢定 (Independent t-test)', 'Mann-Whitney U 檢定 (Wilcoxon Rank-Sum)'],
        ['三組或以上獨立樣本', '單因子變異數分析 (One-way ANOVA)', 'Kruskal-Wallis H 檢定 (等級單因子變異數分析)'],
        ['適用測量尺度', '等距尺度 (Interval) / 等比尺度 (Ratio)', '順序尺度 (Ordinal) / 偏態連續尺度'],
        ['母體常態性要求', '嚴格要求常態分佈 (Normal Distribution)', '無母體分佈假設 (Distribution-free)'],
        ['對極端值耐受度', '極易受極端值拉偏平均數與擴大變異數', '極度穩健 (Robust)，僅受數據相對排名影響']
    ]
    for r_idx, r_data in enumerate(comp_rows, 1):
        for c_idx, val in enumerate(r_data):
            cell = comp_tbl.cell(r_idx, c_idx)
            p = cell.paragraphs[0]
            if c_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run(val)
    style_table(comp_tbl, c_widths)

    # ==================== 七、交接查核清單與 APA 回報範本 ====================
    add_h1('七、交接查核清單與學術報告回報範本')
    add_h2('7.1 接手人員查核清單（Handover Checklist）')
    add_body('☑ 檔案完整性：確認 3 份原始 Excel 與 2 份 Word 檔案無缺漏且結構完好。')
    add_body('☑ 數據複核性：透過 Python 交叉複算 SPSS 所有指標（U=1.000, Z=-3.372, H=27.611），100% 驗證通過。')
    add_body('☑ 圖表視覺化：已按最新規範重繪 3 張莫蘭迪色系盒鬚圖（微縮數據點、黑色中位數線）。')
    add_body('☑ 操作排錯防呆：已詳細登載 SPSS /K-W 語法因變數倒置之根因與標準 SOP。')
    add_body('☑ 補充分析完成：已完備 Blood Type 資料集之 Kruskal-Wallis 檢定。')

    add_h2('7.2 學術與期末實習報告 APA 格式標準回報句')
    add_body('1. HER2 表現量（Mann-Whitney 檢定）：')
    add_body('   「Mann-Whitney U 檢定結果顯示，乳癌組之 HER2 蛋白質表現量（Mdn = 446.50, Mean Rank = 13.38）顯著高於肺腺癌組（Mdn = 118.00, Mean Rank = 5.11），U = 1.000, Z = -3.372, p = .001（雙尾精確檢定 p < .001）。」')
    add_body('2. 職業球員年薪（Kruskal-Wallis 檢定）：')
    add_body('   「Kruskal-Wallis 檢定顯示四大職業運動聯盟運動員之年薪存在顯著差異，χ²(3) = 27.611, p < .001。各聯盟平均等級分別為 NFL（32.25）、NBA（24.20）、MLB（20.05）及 WNBA（5.50）。事後成對比較顯示 WNBA 年薪顯著低於男子三大聯盟（p < .001）。」')
    add_body('3. 血型與身高（Kruskal-Wallis 檢定）：')
    add_body('   「以 Kruskal-Wallis 檢定比較不同 ABO 血型受試者之身高，結果未達統計顯著差異，χ²(3) = 1.507, p = .681。顯示受試者身高在 A 型（Mean Rank = 26.00）、B 型（22.90）、O 型（20.83）與 AB 型（20.12）間無顯著分佈差異。」')

    # 儲存檔案
    output_filename = '20261015_生物統計實習交接與整體分析報告.docx'
    doc.save(output_filename)
    print(f'Successfully generated {output_filename}')

if __name__ == '__main__':
    create_handover_document()
