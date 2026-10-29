import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def build_docx():
    doc = docx.Document()
    
    # Page setup - A4 with 1 inch (2.54cm) margins
    for section in doc.sections:
        section.page_width = Inches(8.27)  # A4
        section.page_height = Inches(11.69)
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        
        # Header & Footer
        header = section.header
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hrun = hp.add_run("橘郡異常事件流行病學調查報告【公衛內部通報·極機密】 | 2026-10-22")
        hrun.font.name = "Microsoft JhengHei"
        hrun.font.size = Pt(8.5)
        hrun.font.color.rgb = RGBColor(120, 120, 120)
        
        footer = section.footer
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        frun = fp.add_run("衛生調查局流行病學調查小組 · 專案調查員實地回報報告書")
        frun.font.name = "Microsoft JhengHei"
        frun.font.size = Pt(8.5)
        frun.font.color.rgb = RGBColor(120, 120, 120)

    # Styles setup
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Microsoft JhengHei'
    style_normal.font.size = Pt(10.5)
    style_normal.font.color.rgb = RGBColor(40, 40, 40)
    style_normal.paragraph_format.line_spacing = 1.25
    style_normal.paragraph_format.space_after = Pt(4)

    # Colors
    NAVY = RGBColor(27, 54, 93)      # #1B365D - Primary Header
    DARK_BLUE = RGBColor(41, 128, 185) # #2980B9 - Secondary
    CRIMSON = RGBColor(192, 57, 43)  # #C0392B - Alert / Risk
    CHARCOAL = RGBColor(44, 62, 80)  # #2C3E50 - Body title
    GRAY = RGBColor(100, 100, 100)

    # Helper functions
    def add_title(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(text)
        run.font.name = 'Microsoft JhengHei'
        run.font.size = Pt(20)
        run.font.bold = True
        run.font.color.rgb = NAVY
        return p

    def add_subtitle(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(14)
        run = p.add_run(text)
        run.font.name = 'Microsoft JhengHei'
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = DARK_BLUE
        return p

    def add_h1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Microsoft JhengHei'
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.color.rgb = NAVY
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Microsoft JhengHei'
        run.font.size = Pt(11.5)
        run.font.bold = True
        run.font.color.rgb = CHARCOAL
        return p

    def add_bullet(p, bold_prefix, text):
        run_b = p.add_run(bold_prefix)
        run_b.font.bold = True
        run_b.font.color.rgb = CHARCOAL
        run_t = p.add_run(text)
        run_t.font.color.rgb = RGBColor(50, 50, 50)

    def set_cell_shading(cell, color_hex):
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
        cell._tc.get_or_add_tcPr().append(shd)

    def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
        tcPr.append(tcMar)

    def format_table(table, col_widths, align_right_cols=[]):
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        # Borders: APA style 3 lines
        tblPr = table._tbl.tblPr
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'<w:top w:val="single" w:sz="12" w:space="0" w:color="1B365D"/>'
            f'<w:bottom w:val="single" w:sz="12" w:space="0" w:color="1B365D"/>'
            f'<w:left w:val="none"/>'
            f'<w:right w:val="none"/>'
            f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="D5D8DC"/>'
            f'<w:insideV w:val="none"/>'
            f'</w:tblBorders>'
        )
        tblPr.append(borders)
        
        # Header formatting
        for cell in table.rows[0].cells:
            set_cell_shading(cell, "1B365D")
            set_cell_margins(cell, top=140, bottom=140, left=150, right=150)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in p.runs:
                    run.font.name = 'Microsoft JhengHei'
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(255, 255, 255)
                    run.font.size = Pt(9.5)
                    
        # Body rows
        for r_idx, row in enumerate(table.rows[1:], start=1):
            bg = "F9FBFC" if r_idx % 2 == 1 else "FFFFFF"
            for c_idx, cell in enumerate(row.cells):
                set_cell_shading(cell, bg)
                set_cell_margins(cell, top=100, bottom=100, left=150, right=150)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                for p in cell.paragraphs:
                    p.paragraph_format.space_after = Pt(0)
                    p.paragraph_format.space_before = Pt(0)
                    if c_idx in align_right_cols:
                        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    elif c_idx == 0:
                        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    else:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in p.runs:
                        run.font.name = 'Microsoft JhengHei'
                        run.font.size = Pt(9.5)
                        
        # Column widths
        for row in table.rows:
            for idx, width in enumerate(col_widths):
                row.cells[idx].width = width

    # ==========================================
    # Document Header & Meta Information Box
    # ==========================================
    add_title("橘郡異常事件現場流行病學調查報告書")
    add_subtitle("Epidemiological Investigation Report on Orange County Outbreak Event")

    # Metadata table (Official Dispatch memo)
    memo_table = doc.add_table(rows=4, cols=4)
    memo_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    memo_data = [
        [("通報辦公室", True), ("流行病學調查辦公室 / 疾管指揮中心", False), ("調查日期", True), ("2026年10月22日", False)],
        [("主辦調查員", True), ("專案流行病學調查組 特約調查員", False), ("事件代號", True), ("OC-20261022-EPID", False)],
        [("調查地點", True), ("橘郡（Orange County）受災社區", False), ("案件密等", True), ("極機密 / 公衛緊急事件", False)],
        [("受查總人數", True), ("669 人（全數完成流行病學個案追蹤）", False), ("調查狀態", True), ("第一階段完成，致病因鎖定並呈報", False)]
    ]
    for r_idx, row in enumerate(memo_table.rows):
        for c_idx, cell in enumerate(row.cells):
            label, is_bold = memo_data[r_idx][c_idx]
            cell.text = label
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            if c_idx % 2 == 0:
                set_cell_shading(cell, "EAECEE")
            else:
                set_cell_shading(cell, "F8F9FA")
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx % 2 == 1 else WD_ALIGN_PARAGRAPH.RIGHT
            p.paragraph_format.space_after = Pt(0)
            run = p.runs[0]
            run.font.name = 'Microsoft JhengHei'
            run.font.size = Pt(9)
            if is_bold:
                run.font.bold = True
                run.font.color.rgb = NAVY

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # Executive Summary / Memo to office
    p_memo = doc.add_paragraph()
    p_memo.paragraph_format.space_after = Pt(8)
    p_memo.paragraph_format.line_spacing = 1.3
    run_memo_title = p_memo.add_run("【調查員回報辦公室受文簡述】\n")
    run_memo_title.font.bold = True
    run_memo_title.font.size = Pt(11)
    run_memo_title.font.color.rgb = NAVY
    p_memo.add_run(
        "報告長官：\n"
        "職受命前往橘郡（Orange County）展開不明致死性疾病異常事件之現場流行病學調查。本專案共計追蹤調查全區受影響居民 669 名，"
        "現已全數完成流行病學個案訪談、水源暴露史調查、疾病嚴重程度判定及生存結局確認。經資料庫建立與生物統計學推論分析，"
        "目前已完全掌握疫情之關鍵致命根源。以下謹將 5 大核心數據指標，以及包含卡方檢定（Chi-Square Test）、風險比（Risk Ratio, RR）"
        "與勝算比（Odds Ratio, OR）之因果推論實證報告呈報辦公室，以供緊急公衛決策與源頭斷源作為之依據。"
    )

    # Callout Box: 5 Key Indicators Summary
    p_callout = doc.add_paragraph()
    p_callout.paragraph_format.space_before = Pt(4)
    p_callout.paragraph_format.space_after = Pt(10)
    box_table = doc.add_table(rows=1, cols=1)
    box_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_box = box_table.rows[0].cells[0]
    c_box.width = Inches(6.27)
    set_cell_shading(c_box, "F2F4F4")
    set_cell_margins(c_box, top=140, bottom=140, left=180, right=180)
    p_b = c_box.paragraphs[0]
    p_b.paragraph_format.space_after = Pt(0)
    r_bt = p_b.add_run("★ 調查員緊急回報辦公室：5 大核心數據速查摘要\n")
    r_bt.font.bold = True
    r_bt.font.size = Pt(11)
    r_bt.font.color.rgb = CRIMSON
    
    add_bullet(p_b, "1. 患病比例：", "76.98%（總計 515 / 669 人患病；其中輕度患病佔 30.94%，重度患病佔 46.04%）。\n")
    add_bullet(p_b, "2. 飲用水來源種類比例：", "深井水（Deep well）佔 75.04%（502人）、淺井水（shallow well）佔 17.19%（115人）、自來水（Top water）佔 7.77%（52人）。\n")
    add_bullet(p_b, "3. 患病嚴重程度分布：", "健康無病（Healthy）佔 23.02%（154人）、輕度病況（minor）佔 30.94%（207人）、重度病況（severe）佔 46.04%（308人）。\n")
    add_bullet(p_b, "4. 疾病死亡率：", "全體居民總粗死亡率為 76.98%（515 / 669 人死亡）；患病者致死率（Case Fatality Rate）高達 100.00%（515 / 515 人，只要患病全數死亡！）。\n")
    add_bullet(p_b, "5. 性別比例：", "女性佔 61.14%（409人）、男性佔 38.86%（260人）；男女比例為 0.636 : 1（即女男比例 1.573 : 1）。")

    # ==========================================
    # SECTION 1
    # ==========================================
    add_h1("第一部分：橘郡居民人口學與健康狀況描述性統計分析")
    
    p_s1 = doc.add_paragraph()
    p_s1.add_run(
        "本階段調查以橘郡異常事件爆發區域之全部 669 位居民為分析母群體（Cohort）。調查指標涵蓋性別、年齡、飲用水來源、"
        "疾病臨床發病嚴重程度以及最終生命結局（存活或死亡）。各項指標分佈經嚴格核算，彙整如【表 1】所示："
    )

    # Table 1: Descriptive Statistics Table
    t1 = doc.add_table(rows=17, cols=5)
    t1_widths = [Inches(1.8), Inches(2.0), Inches(1.1), Inches(1.1), Inches(1.2)]
    t1_data = [
        ["分析變項", "分組類別", "人數 (N)", "佔總體比例 (%)", "備註說明"],
        ["患病狀態", "健康無病 (Healthy)", "154", "23.02%", "全數存活"],
        ["", "患病 (輕度 minor + 重度 severe)", "515", "76.98%", "全體罹患率 (Attack Rate)"],
        ["飲用水來源種類", "深井水 (Deep well)", "502", "75.04%", "主要供水來源 (高危水源)"],
        ["", "淺井水 (shallow well)", "115", "17.19%", "次要供水來源 (低危水源)"],
        ["", "自來水 (Top water)", "52", "7.77%", "安全乾淨水源 (零患病)"],
        ["患病嚴重程度分佈", "健康無病 (Healthy)", "154", "23.02%", "臨床無症狀"],
        ["", "輕度患病 (minor)", "207", "30.94%", "佔患病者之 40.19%"],
        ["", "重度患病 (severe)", "308", "46.04%", "佔患病者之 59.81%"],
        ["生存結局與死亡率", "存活 (alive)", "154", "23.02%", "存活者全數為健康人+1深井"],
        ["", "死亡 (Dead)", "515", "76.98%", "全體粗死亡率 (Mortality)"],
        ["", "患病者病死率 (Case Fatality)", "515 / 515", "100.00%", "患病個案致死率極端兇猛"],
        ["性別分佈", "女性 (Female, F)", "409", "61.14%", "女男比 1.573 : 1"],
        ["", "男性 (Male, M)", "260", "38.86%", "男女比 0.636 : 1"],
        ["年齡分組分佈", "年輕族群 (<= 40 歲)", "349", "52.17%", "平均年齡 39.67 歲"],
        ["", "中高齡族群 (41+ 歲)", "320", "47.83%", "年齡範圍 8 ~ 90 歲"],
        ["總計母群體", "橘郡受查總人口", "669", "100.00%", "完全調查無失聯"]
    ]
    for r_idx, row_content in enumerate(t1_data):
        row = t1.rows[r_idx]
        for c_idx, val in enumerate(row_content):
            row.cells[c_idx].text = val
    format_table(t1, t1_widths, align_right_cols=[2, 3])

    p_t1_note = doc.add_paragraph()
    p_t1_note.paragraph_format.space_after = Pt(8)
    r_n = p_t1_note.add_run("註：資料來源為現場調查員個案登記簿（N=669）。百分比採四捨五入計算至小數點後兩位。")
    r_n.font.size = Pt(8.5)
    r_n.font.italic = True
    r_n.font.color.rgb = GRAY

    # Subsections for Section 1
    add_h2("1.1 患病比例（Disease Morbidity / Attack Rate）分析")
    p = doc.add_paragraph()
    p.add_run(
        "橘郡全體受查居民 669 人中，確診患病者高達 515 人，整體患病比例（罹病率/侵襲率 Attack Rate）為 76.98%（約 77.0%）。"
        "未受感染或未發病之健康人口僅 154 人，佔 23.02%（約 23.0%）。此一極高的罹病率顯示該致病因子具有極強的穿透力與普遍暴露性，"
        "符合典型「共同媒介物（Common Source/Vehicle）」爆發之特徵。"
    )

    add_h2("1.2 飲用水來源種類比例（Drinking Water Sources Distribution）")
    p = doc.add_paragraph()
    p.add_run(
        "調查發現該社區飲用水來源存在三種主要形式：\n"
        "（1）深井水（Deep well）：共 502 人飲用，佔總人口 75.04%，為該郡最主要的民生自備水源；\n"
        "（2）淺井水（shallow well）：共 115 人飲用，佔總人口 17.19%；\n"
        "（3）自來水（Top water）：共 52 人飲用，佔總人口 7.77%，多為鄰近主幹道管網之住戶。\n"
        "三者合規劃分為兩大暴露類群：深井水組（water_group = 2，佔 75.04%）與非深井水組（water_group = 1，包含淺井水與自來水，共 167 人，佔 24.96%）。"
    )

    add_h2("1.3 患病嚴重程度分布情形（Disease Severity Distribution）")
    p = doc.add_paragraph()
    p.add_run(
        "依據臨床病徵與病程進展，全體居民健康狀態被嚴格歸納為三級：\n"
        "• 健康狀態（Healthy）：154 人（23.02%），受查期間完全無任何臨床症狀，生命徵象穩定。\n"
        "• 輕度患病（minor）：207 人（30.94%），出現急性中毒或感染初期症狀（如嚴重噁心、腹絞痛、電解質失衡與中樞神經輕度抑制）。\n"
        "• 重度患病（severe）：308 人（46.04%），迅速併發急性多重器官衰竭、劇烈抽搐、昏迷與不可逆性休克。\n"
        "若以患病人口（515人）作為分母計算內部嚴重度比例，輕度患者佔 40.19%（207/515），重度患者則高達 59.81%（308/515），顯示重症比例極高。"
    )

    add_h2("1.4 疾病死亡率與致死率（Mortality Rate & Case Fatality Rate）")
    p = doc.add_paragraph()
    p.add_run(
        "本事件之死亡指標呈现極其駭人之結果：\n"
        "• 全體居民總粗死亡率（Overall Crude Mortality Rate）：全體 669 人中死亡 515 人，粗死亡率高達 76.98%。\n"
        "• 患病者致死率（Case Fatality Rate, CFR）：在 515 名發病個案中，死亡人數亦為 515 人，致死率達到恐怖的 100.00%！\n"
        "• 健康者死亡率：在 154 名健康者中，死亡人數為 0 人（0.00%），全數存活。\n"
        "調查員特別警示辦公室：本疾病一旦發病（無論初期判定為 minor 或 severe），病程演變皆為致命性結局，致死率為 100%，"
        "凸顯該病原體或化學毒物對人體具有不可逆的毀滅性殺傷力。"
    )

    add_h2("1.5 性別比例與人口學特徵（Gender & Demographic Distribution）")
    p = doc.add_paragraph()
    p.add_run(
        "在受查人口中，女性居民共有 409 人，佔比 61.14%；男性居民共有 260 人，佔比 38.86%。\n"
        "男女比例（Male : Female）為 0.636 : 1（即每 100 名女性對應約 63.6 名男性）；女男比例（Female : Male）為 1.573 : 1。\n"
        "在年齡方面，全體居民平均年齡為 39.67 ± 18.72 歲（範圍 8 至 90 歲）。40 歲以下年輕族群佔 52.17%（349人），41 歲以上中高齡族群佔 47.83%（320人），"
        "人口年齡結構分佈均勻。"
    )

    # Insert Chart 1
    if os.path.exists('chart1_water_survival.png'):
        p_img1 = doc.add_paragraph()
        p_img1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img1.paragraph_format.space_before = Pt(8)
        p_img1.paragraph_format.space_after = Pt(2)
        run_img1 = p_img1.add_run()
        run_img1.add_picture('chart1_water_survival.png', width=Inches(5.8))
        
        p_cap1 = doc.add_paragraph()
        p_cap1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap1.paragraph_format.space_after = Pt(12)
        r_c1 = p_cap1.add_run("圖 1：橘郡居民不同飲用水來源之存活與死亡分佈對比圖（含各類水源死亡率趨勢線）")
        r_c1.font.size = Pt(9)
        r_c1.font.bold = True
        r_c1.font.color.rgb = DARK_BLUE

    # ==========================================
    # SECTION 2
    # ==========================================
    add_h1("第二部分：居民患病及死亡之根本原因調查與推論性統計檢定")
    
    p_s2_intro = doc.add_paragraph()
    p_s2_intro.add_run(
        "為徹底追查造成橘郡 515 名居民集體發病且全數死亡的根本原因，調查小組針對流行病學三大經典假說進行了假設檢定：\n"
        "【假說一】飲用水源暴露假說（深井水污染）；\n"
        "【假說二】性別易感性假說（特定性別生理弱勢）；\n"
        "【假說三】年齡易感性假說（老年或幼童抵抗力脆弱）。\n"
        "利用 Pearson 卡方檢定（Chi-Square Test）、連續性校正（Yates' Continuity Correction）、費雪精確檢定（Fisher's Exact Test），"
        "並計算風險比（Relative Risk / Risk Ratio, RR）與勝算比（Odds Ratio, OR）及其 95% 信賴區間（CI），檢定結果獲致突破性結論。"
    )

    add_h2("2.1 致病與致死根本原因：深井地下水源（Deep well）遭受致命污染")
    p = doc.add_paragraph()
    p.add_run(
        "統計結果明確指出：飲用深井水（Deep well）是造成居民集體患病與死亡的單一絕對主因。\n"
        "在飲用深井水的 502 名居民中，高達 501 人死亡，死亡率高達 99.80%，僅 1 人倖免於難；"
        "反觀飲用非深井水（淺井水與自來水）的 167 名居民中，僅 14 人發病死亡，死亡率僅 8.38%，存活率高達 91.62%！"
        "以下為 2x2 列聯表及相關檢定數據："
    )

    # Table 2: 2x2 Contingency Table for Deep Well vs Mortality
    t2 = doc.add_table(rows=5, cols=6)
    t2_widths = [Inches(1.8), Inches(1.1), Inches(1.1), Inches(1.1), Inches(1.2), Inches(1.2)]
    t2_data = [
        ["水源分組類別", "死亡 (Dead)", "存活 (Alive)", "合計人數", "死亡率 (Risk)", "死亡勝算 (Odds)"],
        ["暴露組：深井水 (Deep well)", "501", "1", "502", "99.80%", "501.000"],
        ["非暴露組：非深井水 (淺井+自來)", "14", "153", "167", "8.38%", "0.0915"],
        ["合計人數", "515", "154", "669", "76.98%", "3.3442"],
        ["統計檢定指標", "Pearson 卡方值: 590.987 (df=1, P < 0.001)\n校正卡方值: 585.839 (P < 0.001)\nFisher 精確 P: 2.09 × 10⁻¹³³", "", "", "風險比 (RR):\n11.90 (7.21-19.66)", "勝算比 (OR):\n5475.21 (714.18-41975.49)"]
    ]
    # merge bottom row cells for stats
    for r_idx, row_content in enumerate(t2_data):
        row = t2.rows[r_idx]
        for c_idx, val in enumerate(row_content):
            row.cells[c_idx].text = val
    # merge bottom row cells 1, 2, 3
    t2.rows[4].cells[0].merge(t2.rows[4].cells[1]).merge(t2.rows[4].cells[2]).merge(t2.rows[4].cells[3])
    format_table(t2, t2_widths, align_right_cols=[1, 2, 3, 4, 5])

    p_t2_desc = doc.add_paragraph()
    p_t2_desc.paragraph_format.space_before = Pt(4)
    p_t2_desc.paragraph_format.space_after = Pt(6)
    add_bullet(p_t2_desc, "★ 關鍵統計檢定數據解讀：\n", "")
    add_bullet(p_t2_desc, "1. 卡方檢定統計量（Chi-Square）：", 
               "Pearson χ² = 590.987（自由度 df = 1，P = 1.53 × 10⁻¹³⁰ < 0.001）；Yates 連續性校正卡方值 χ² = 585.839（P < 0.001）；"
               "Fisher 費雪精確檢定雙尾 P = 2.09 × 10⁻¹³³。P 值遠低於 0.001 之極顯著水準，強烈拒絕虛無假說，證實深井水暴露與居民發病/死亡存在極端顯著之統計關聯。\n")
    add_bullet(p_t2_desc, "2. 風險比（Risk Ratio / Relative Risk, RR）：", 
               "RR = (501 / 502) / (14 / 167) = 0.9980 / 0.0838 = 11.9048（95% 信賴區間 CI: 7.2105 ~ 19.6554）。"
               "這意味著：飲用深井水居民之死亡風險是非深井水居民的整整 11.90 倍！95% 信賴區間下限高達 7.21 倍，具有無可辯駁的極高度危險性。\n")
    add_bullet(p_t2_desc, "3. 勝算比（Odds Ratio, OR）：", 
               "OR = (501 × 153) / (1 × 14) = 76,653 / 14 = 5475.2143（95% 信賴區間 CI: 714.1780 ~ 41,975.4927）。"
               "在病例對照的勝算維度上，飲用深井水之死亡勝算是非深井水組的 5475.21 倍！\n")
    add_bullet(p_t2_desc, "4. 患病關聯同質性：", 
               "因所有患病者皆死亡（致死率 100%），深井水暴露與「患病（Diseased）」的關聯檢定與上述「死亡」之列聯表及數值完全一致（χ² = 590.987, RR = 11.90, OR = 5475.21）。")

    # Insert Chart 3 (Forest Plot)
    if os.path.exists('chart3_forest_rr.png'):
        p_img3 = doc.add_paragraph()
        p_img3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img3.paragraph_format.space_before = Pt(8)
        p_img3.paragraph_format.space_after = Pt(2)
        run_img3 = p_img3.add_run()
        run_img3.add_picture('chart3_forest_rr.png', width=Inches(6.0))
        
        p_cap3 = doc.add_paragraph()
        p_cap3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap3.paragraph_format.space_after = Pt(12)
        r_c3 = p_cap3.add_run("圖 3：可疑危險因子之風險比 (RR) 與 95% 信賴區間森林圖（對數尺度展示深井水之顯著性）")
        r_c3.font.size = Pt(9)
        r_c3.font.bold = True
        r_c3.font.color.rgb = DARK_BLUE

    add_h2("2.2 飲用水來源之劑量-反應效應（Dose-Response Effect）分析")
    p = doc.add_paragraph()
    p.add_run(
        "在流行病學病因推論中，「生物梯度（Biological Gradient）」或「劑量-反應關係」是確立因果關聯的最有力證據之一。"
        "本調查將三種飲用水來源與疾病發病及嚴重程度進行細緻交叉分析，結果呈現教科書級別的水源污染梯度："
    )

    # Table 3: 3x3 table of water source vs severity
    t3 = doc.add_table(rows=5, cols=6)
    t3_widths = [Inches(1.8), Inches(1.1), Inches(1.1), Inches(1.1), Inches(1.1), Inches(1.1)]
    t3_data = [
        ["飲用水來源", "健康無病", "輕度患病 (minor)", "重度患病 (severe)", "合計人數", "重症比率 (%)"],
        ["自來水 (Top water)", "52 (100.0%)", "0 (0.0%)", "0 (0.0%)", "52", "0.00%"],
        ["淺井水 (shallow well)", "101 (87.8%)", "14 (12.2%)", "0 (0.0%)", "115", "0.00%"],
        ["深井水 (Deep well)", "1 (0.2%)", "193 (38.4%)", "308 (61.4%)", "502", "61.35%"],
        ["卡方檢定統計量", "Pearson χ² = 598.030, 自由度 df = 4, P = 4.14 × 10⁻¹²⁸ (P < 0.001 極顯著)", "", "", "", ""]
    ]
    for r_idx, row_content in enumerate(t3_data):
        row = t3.rows[r_idx]
        for c_idx, val in enumerate(row_content):
            row.cells[c_idx].text = val
    t3.rows[4].cells[0].merge(t3.rows[4].cells[1]).merge(t3.rows[4].cells[2]).merge(t3.rows[4].cells[3]).merge(t3.rows[4].cells[4]).merge(t3.rows[4].cells[5])
    format_table(t3, t3_widths, align_right_cols=[1, 2, 3, 4, 5])

    p_t3_desc = doc.add_paragraph()
    p_t3_desc.paragraph_format.space_before = Pt(4)
    p_t3_desc.paragraph_format.space_after = Pt(6)
    p_t3_desc.add_run(
        "由【表 3】可清晰觀察到劑量梯級關係：\n"
        "1. 自來水組（Top water, N=52）：52 人全部為健康存活（100%），患病率 0%，重症率 0%，死亡率 0%。證實市政自來水系統具備完善防護，完全未受污染。\n"
        "2. 淺井水組（shallow well, N=115）：87.8% 居民健康，僅 12.2%（14人）發病且均屬輕度（minor），重症率為 0%。這顯示淺井水僅受到極輕微之滲透或周邊交叉微量污染。\n"
        "3. 深井水組（Deep well, N=502）：僅 1 人健康（0.2%），高達 99.8% 居民發病，且其中 61.4%（308人）直接陷入器官衰竭之重度（severe）狀態，全數死亡！\n"
        "交叉表卡方值高達 χ² = 598.030（df = 4, P < 0.001），證實污染源深度集中於深層地下水層（Deep aquifer），水質毒性濃度與疾病嚴重度成高度正向梯度相關。"
    )

    # Insert Chart 2
    if os.path.exists('chart2_water_severity.png'):
        p_img2 = doc.add_paragraph()
        p_img2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img2.paragraph_format.space_before = Pt(8)
        p_img2.paragraph_format.space_after = Pt(2)
        run_img2 = p_img2.add_run()
        run_img2.add_picture('chart2_water_severity.png', width=Inches(5.8))
        
        p_cap2 = doc.add_paragraph()
        p_cap2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap2.paragraph_format.space_after = Pt(12)
        r_c2 = p_cap2.add_run("圖 2：不同飲用水來源之疾病嚴重程度百分比堆疊圖（展示鮮明之生物毒性梯度效應）")
        r_c2.font.size = Pt(9)
        r_c2.font.bold = True
        r_c2.font.color.rgb = DARK_BLUE

    add_h2("2.3 排除其他潛在干擾因素：性別與年齡易感性檢驗")
    p = doc.add_paragraph()
    p.add_run(
        "為防止混淆變項（Confounding factors）干擾調查結論，調查員對「性別」與「年齡分組」進行嚴密的獨立性卡方檢定與風險評估，"
        "實證結果證實性別與年齡均非致病或死亡之危險因子："
    )

    # Table 4: Confounder Analysis (Gender and Age)
    t4 = doc.add_table(rows=5, cols=7)
    t4_widths = [Inches(1.2), Inches(1.3), Inches(0.9), Inches(0.9), Inches(1.0), Inches(1.1), Inches(1.2)]
    t4_data = [
        ["分析變項", "分組對比", "死亡人數", "存活人數", "死亡率 (%)", "卡方值 (P值)", "風險比 RR (95% CI)"],
        ["性別因素", "男性 (Male, N=260)", "200", "60", "76.92%", "χ² = 0.001\n(P = 0.978)", "0.999 (0.917 - 1.088)\n[勝算比 OR = 0.995]"],
        ["", "女性 (Female, N=409)", "315", "94", "77.02%", "(無顯著差異)", "基準對照 (Ref)"],
        ["年齡因素", "中高齡 (41+ 歲, N=320)", "247", "73", "77.19%", "χ² = 0.015\n(P = 0.903)", "1.005 (0.925 - 1.092)\n[勝算比 OR = 1.023]"],
        ["", "年輕組 (<= 40 歲, N=349)", "268", "81", "76.79%", "(無顯著差異)", "基準對照 (Ref)"]
    ]
    for r_idx, row_content in enumerate(t4_data):
        row = t4.rows[r_idx]
        for c_idx, val in enumerate(row_content):
            row.cells[c_idx].text = val
    format_table(t4, t4_widths, align_right_cols=[2, 3, 4, 5, 6])

    p_t4_desc = doc.add_paragraph()
    p_t4_desc.paragraph_format.space_before = Pt(4)
    p_t4_desc.paragraph_format.space_after = Pt(6)
    p_t4_desc.add_run(
        "檢定結果解讀：\n"
        "• 性別檢定：男性死亡率 76.92% vs. 女性死亡率 77.02%，Pearson χ² = 0.001（df = 1, P = 0.978），RR = 0.999（95% CI: 0.917-1.088），"
        "OR = 0.995（95% CI: 0.688-1.439）。P 值遠大於 0.05 且 RR 信賴區間包含 1.0，證實男女罹病及死亡風險完全相同，性別無統計顯著影響。\n"
        "• 年齡檢定：41 歲以上死亡率 77.19% vs. 40 歲以下死亡率 76.79%，Pearson χ² = 0.015（df = 1, P = 0.903），RR = 1.005（95% CI: 0.925-1.092），"
        "OR = 1.023（95% CI: 0.713-1.466）。P 值同樣遠大於 0.05，年齡分組與發病死亡無任何統計關聯。\n"
        "結論：完全排除性別與年齡之混淆作用，進一步坐實深井水是唯一的暴露病因。"
    )

    add_h2("2.4 布萊德福·希爾（Bradford Hill）因果關係準則全面檢驗")
    p = doc.add_paragraph()
    p.add_run(
        "為使本調查報告具備無懈可擊之科學說服力，調查員採用現代流行病學黃金標準 Bradford Hill 因果準則進行綜合檢驗：\n"
        "1. 關聯強度（Strength）：深井水暴露之 RR 高達 11.90 倍，OR 高達 5475.21 倍，χ² = 590.99，關聯強度處於極端顯著之罕見水平。\n"
        "2. 劑量反應關係（Biological Gradient）：呈現「自來水 (0%) → 淺井水 (12.2% 輕症) → 深井水 (99.8% 且 61.4% 重症)」之完美階梯式劑量效應。\n"
        "3. 時序性（Temporality）：所有死亡居民均先有長期或近日大量飲用深井水之暴露史，隨後出現中毒/發病症狀，最終器官衰竭死亡，時序完全契合。\n"
        "4. 特異性（Specificity）：疾病暴發高度集中於深井水管網用戶，非深井水用戶（尤其是自來水戶）完全具備保護性免疫屏障（0% 發病）。\n"
        "5. 生物合理性與一致性（Plausibility & Consistency）：深層地下水層易受地質重金屬（如高濃度砷、鎘、放射性氡氣）或工業違法深井灌注劇毒化學物（如氰化物、有機磷、有機溶劑）污染，"
        "其高濃度攝入能在短時間內造成致命性中樞神經崩潰與急性腎肝衰竭，致死率 100% 之臨床表現完全具備毒理學合理性。"
    )

    # ==========================================
    # SECTION 3
    # ==========================================
    add_h1("第三部分：調查員實地防制建議與緊急公衛干預策略")
    
    p_s3 = doc.add_paragraph()
    p_s3.add_run(
        "基於上述嚴密之流行病學調查與生物統計檢定結果，調查員正式向辦公室及主管機關提出以下四項「黃金處置作為」，建請即刻批准執行："
    )

    rec_items = [
        ("一、 立即實體封閉橘郡境內所有深井水源，切斷毒性傳播鏈",
         "動員警政與水利單位，即刻查封橘郡全區所有深層地下水抽水馬達與取水井口，拉設封鎖線並張貼緊急危險告示；"
         "全面截斷深井水管網閥門，嚴禁居民直接飲用、烹飪或沐浴，以杜絕新增個案。"),
        ("二、 啟動應急自來水延伸與民生安全包裝飲用水全面配送",
         "調派國軍工兵部隊與自來水公司水車進駐橘郡各里，每日配發高規格瓶裝水；同時緊急搶修與延伸既有市政自來水（Top water）主管線，"
         "利用已知安全之自來水管網全面取代地下水系統。"),
        ("三、 專案化驗小組進駐：進行深層水體毒物學、重金屬與生化戰劑全譜篩檢",
         "採集深井水（水頭、水底沉澱物）、淺井水對照組及自來水樣本，送交國家級毒物檢驗所進行 GC-MS、ICP-MS 質譜分析，"
         "重點清查砷化物、有機磷、氰化物、重金屬及地底特殊細菌毒素（如肉毒桿菌毒素），並回溯深層水文地質周邊是否有工廠非法暗管偷排。"),
        ("四、 設立前進醫療指揮所與存活個案（N=154）強化醫學監測",
         "針對唯一步入深井暴露卻奇蹟存活之 1 名個案進行全基因組定序與毒物代謝分析，探討其存活機轉；"
         "對淺井水已發病之 14 名輕症（minor）轉院個案實施重症加護治療與血液透析，防止其轉變為死亡結局；對全體居民提供創傷後心理輔導。")
    ]

    for title, desc in rec_items:
        p_r = doc.add_paragraph()
        p_r.paragraph_format.space_before = Pt(4)
        p_r.paragraph_format.space_after = Pt(4)
        add_bullet(p_r, f"{title}\n", desc)

    # Closing & Signature Block
    p_sign = doc.add_paragraph()
    p_sign.paragraph_format.space_before = Pt(16)
    p_sign.paragraph_format.keep_with_next = True
    
    sign_table = doc.add_table(rows=2, cols=2)
    sign_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sign_widths = [Inches(3.1), Inches(3.1)]
    
    sign_data = [
        ["報告呈報調查員：", "辦公室長官核閱簽章："],
        ["專案流行病學調查員  劉昱君  (簽署)\n日期：2026 年 10 月 22 日", "流行病學調查指揮中心 辦公室主任\n核定日期：2026 年 ___ 月 ___ 日"]
    ]
    for r_idx, row_content in enumerate(sign_data):
        row = sign_table.rows[r_idx]
        for c_idx, val in enumerate(row_content):
            row.cells[c_idx].text = val
            set_cell_margins(row.cells[c_idx], top=80, bottom=80, left=100, right=100)
            if r_idx == 0:
                set_cell_shading(row.cells[c_idx], "EAECEE")
                row.cells[c_idx].paragraphs[0].runs[0].font.bold = True
                row.cells[c_idx].paragraphs[0].runs[0].font.color.rgb = NAVY
            else:
                set_cell_shading(row.cells[c_idx], "FAFAFA")
    format_table(sign_table, sign_widths)

    # Save document
    filename = "橘郡異常事件調查報告_流行病學調查員回報.docx"
    doc.save(filename)
    print(f"Report successfully saved as: {filename}")

if __name__ == '__main__':
    build_docx()
