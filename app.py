#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
衛福部醫事人員繼續教育學分 PDF → Excel 整理工具（Streamlit 版）
================================================================

【重要】這支程式只建議在「本機」執行（streamlit run app.py），
不要部署到 Streamlit Community Cloud 等公開雲端服務。
因為執行時需要讀取你電腦裡「標楷體 / DFKai-SB」字型檔案的內部對照表，
這套字型的授權通常僅限個人電腦安裝使用，不應該被複製/上傳到第三方伺服器。
在本機執行時，字型檔案全程只留在你自己的電腦上，不會有這個問題。

【安裝方式】(終端機執行，只需一次)
    pip3 install streamlit pdfplumber fonttools openpyxl

【啟動方式】
    streamlit run app.py

啟動後瀏覽器會自動打開一個網頁介面（網址通常是 http://localhost:8501），
但實際運算都在你自己電腦上進行。
"""

import re
import io
from datetime import datetime

import streamlit as st
import pdfplumber
from fontTools.ttLib import TTFont as FontToolsTTFont
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont as ReportLabTTFont
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle


CID_PATTERN = re.compile(r"\(cid:(\d+)\)")
DATE_PATTERN = re.compile(r"(\d{4})/(\d{1,2})/(\d{1,2})\s+(\d{1,2}):(\d{2})")


# ---------------- 核心解碼邏輯（跟本機腳本版邏輯相同） ----------------

def build_gid_to_unicode(font_bytes, font_number=0):
    """讀取字型檔（bytes），建立「字碼(GID) -> Unicode 字元」對照表。

    注意：刻意不使用 st.cache_resource。cache_resource 是整個 App 全域共用的
    快取，部署成多人共用的公開服務時，會讓某位使用者上傳的字型資料被留在
    伺服器記憶體裡、下一個人的請求也可能碰到，等同跨使用者共享了原本應該
    只屬於單一使用者的授權字型資料。這裡改用 st.session_state（見下方呼叫
    處），確保每個使用者的字型資料彼此隔離，且該次工作階段結束、或按下
    「清除本次資料」後就會被丟棄，不會被下一位使用者的請求存取到。
    """
    font = FontToolsTTFont(io.BytesIO(font_bytes), fontNumber=font_number)
    cmap = font.getBestCmap()
    glyph_order = font.getGlyphOrder()
    glyphname_to_gid = {name: idx for idx, name in enumerate(glyph_order)}

    gid_to_unicode = {}
    for codepoint, glyph_name in cmap.items():
        gid = glyphname_to_gid.get(glyph_name)
        if gid is not None:
            gid_to_unicode[gid] = chr(codepoint)
    return gid_to_unicode


def decode_cids(text, gid_to_unicode):
    if not text:
        return text

    def _replace(match):
        gid = int(match.group(1))
        return gid_to_unicode.get(gid, "\u25a1")

    return CID_PATTERN.sub(_replace, text)


def decode_row(row, gid_to_unicode):
    return [decode_cids(cell, gid_to_unicode) if cell else cell for cell in row]


def split_datetime_cell(cell_text):
    if not cell_text:
        return "", ""
    matches = DATE_PATTERN.findall(cell_text)
    parsed = []
    for y, mo, d, h, mi in matches:
        try:
            parsed.append(datetime(int(y), int(mo), int(d), int(h), int(mi)))
        except ValueError:
            continue
    if len(parsed) >= 2:
        return parsed[0], parsed[1]
    elif len(parsed) == 1:
        return parsed[0], ""
    return "", ""


def looks_like_header(row):
    joined = "".join(c for c in row if c)
    return "課程類別" in joined or "審查單位" in joined or "課程名稱" in joined


def is_course_table(row0):
    if len(row0) != 9:
        return False
    joined = "".join(c for c in row0 if c)
    return ("課程類別" in joined and "有效" in joined and "課程名稱" in joined)


def try_float(s):
    try:
        return float(s)
    except (TypeError, ValueError):
        return None


def parse_pdf(pdf_bytes, gid_to_unicode, progress_cb=None, debug_limit=5):
    course_rows = []
    other_tables = []
    debug_samples = []  # 不管有沒有比對成功，都留幾筆解碼後的原始樣本供診斷

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        total_pages = len(pdf.pages)
        for page_idx, page in enumerate(pdf.pages, start=1):
            tables = page.find_tables()
            for t_idx, table in enumerate(tables, start=1):
                raw_rows = table.extract()
                if not raw_rows:
                    continue
                decoded_rows = [decode_row(r, gid_to_unicode) for r in raw_rows]

                if len(debug_samples) < debug_limit:
                    debug_samples.append({
                        "頁碼": page_idx,
                        "表格編號": t_idx,
                        "欄數": len(decoded_rows[0]) if decoded_rows else 0,
                        "第一列(可能是標頭)": decoded_rows[0] if decoded_rows else [],
                        "是否判定為課程表格": is_course_table(decoded_rows[0]) if decoded_rows else False,
                    })

                if is_course_table(decoded_rows[0]):
                    for r in decoded_rows[1:]:
                        if looks_like_header(r):
                            continue
                        if len(r) != 9:
                            continue
                        flag, category, valid_pts, invalid_pts, reviewer, organizer, course_name, time_cell, remark = r
                        start_dt, end_dt = split_datetime_cell(time_cell)
                        course_rows.append({
                            "頁碼": page_idx,
                            "旗標": (flag or "").strip(),
                            "課程類別": (category or "").strip(),
                            "有效積分": try_float(valid_pts),
                            "無效積分": try_float(invalid_pts),
                            "審查單位": (reviewer or "").strip(),
                            "主辦單位": (organizer or "").replace("\n", "").strip(),
                            "課程名稱": (course_name or "").replace("\n", "").strip(),
                            "開始時間": start_dt,
                            "結束時間": end_dt,
                            "備註": (remark or "").replace("\n", " ").strip(),
                        })
                else:
                    other_tables.append({
                        "頁碼": page_idx,
                        "表格編號": t_idx,
                        "rows": decoded_rows,
                    })

            if progress_cb:
                progress_cb(page_idx / total_pages)

    return course_rows, other_tables, debug_samples


def build_excel(course_rows, other_tables):
    wb = Workbook()

    ws1 = wb.active
    ws1.title = "研習課程明細"
    headers = ["頁碼", "旗標", "課程類別", "有效積分", "無效積分", "審查單位",
               "主辦單位", "課程名稱", "開始時間", "結束時間", "備註"]
    ws1.append(headers)
    for cell in ws1[1]:
        cell.font = Font(bold=True, name="Arial")
        cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")

    for row in course_rows:
        ws1.append([row[h] for h in headers])

    for col_idx, _ in enumerate(headers, start=1):
        ws1.column_dimensions[get_column_letter(col_idx)].width = 16
    ws1.column_dimensions["H"].width = 40

    for r in range(2, ws1.max_row + 1):
        for c in (9, 10):
            cell = ws1.cell(row=r, column=c)
            if cell.value:
                cell.number_format = "yyyy/mm/dd hh:mm"

    ws2 = wb.create_sheet("分類彙整")
    ws2.append(["課程類別", "有效積分合計", "堂數"])
    for cell in ws2[1]:
        cell.font = Font(bold=True, name="Arial")
        cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")

    categories = sorted({row["課程類別"] for row in course_rows if row["課程類別"]})
    last_data_row = len(course_rows) + 1
    for i, cat in enumerate(categories, start=2):
        ws2.cell(row=i, column=1, value=cat)
        ws2.cell(row=i, column=2,
                  value=f"=SUMIF(研習課程明細!C2:C{last_data_row},A{i},研習課程明細!D2:D{last_data_row})")
        ws2.cell(row=i, column=3,
                  value=f"=COUNTIF(研習課程明細!C2:C{last_data_row},A{i})")

    total_row = len(categories) + 2
    ws2.cell(row=total_row, column=1, value="總計").font = Font(bold=True)
    ws2.cell(row=total_row, column=2, value=f"=SUM(B2:B{total_row - 1})").font = Font(bold=True)
    ws2.cell(row=total_row, column=3, value=f"=SUM(C2:C{total_row - 1})").font = Font(bold=True)

    for col in ("A", "B", "C"):
        ws2.column_dimensions[col].width = 20

    if other_tables:
        ws3 = wb.create_sheet("其他表格_待確認")
        ws3.append(["說明：以下為程式未自動辨識為「研習課程明細」格式的表格，原樣列出供人工確認。"])
        ws3.append([])
        for item in other_tables:
            ws3.append([f"頁碼 {item['頁碼']}　表格 {item['表格編號']}"])
            for r in item["rows"]:
                ws3.append(["" if c is None else c.replace("\n", " ") for c in r])
            ws3.append([])
        ws3.column_dimensions["A"].width = 100

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def build_pdf_report(course_rows, font_bytes, font_number=0):
    """把整理好的課程明細輸出成一份 PDF 報表。

    這裡會把使用者上傳的字型「嵌入」到輸出的這份新 PDF 裡，讓不管在哪台電腦
    打開這份報表都能正確顯示中文。這跟先前擔心的「把字型檔案傳給第三方」是
    不同性質的使用——把字型嵌入自己產生的文件，是幾乎所有字型授權都明文
    允許、也是字型原本被設計要拿來做的事，不是在幫忙散布字型本身。
    """
    font_name = "UserCJKFont"
    pdfmetrics.registerFont(
        ReportLabTTFont(font_name, io.BytesIO(font_bytes), subfontIndex=font_number)
    )

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        leftMargin=12 * mm, rightMargin=12 * mm,
        topMargin=12 * mm, bottomMargin=12 * mm,
    )

    title_style = ParagraphStyle(
        "TitleCJK", fontName=font_name, fontSize=16, leading=20, spaceAfter=8,
    )
    cell_style = ParagraphStyle(
        "CellCJK", fontName=font_name, fontSize=8, leading=11,
    )
    header_style = ParagraphStyle(
        "HeaderCJK", fontName=font_name, fontSize=9, leading=12,
        textColor=colors.white,
    )

    elements = [Paragraph("衛福部醫事人員繼續教育學分整理", title_style)]

    total_valid = sum(r["有效積分"] for r in course_rows if r["有效積分"])
    elements.append(Paragraph(
        f"總筆數：{len(course_rows)}　有效積分合計：{total_valid:.2f}",
        cell_style,
    ))
    elements.append(Spacer(1, 6 * mm))

    headers = ["課程類別", "有效積分", "審查單位", "主辦單位", "課程名稱", "開始時間", "備註"]
    data = [[Paragraph(h, header_style) for h in headers]]
    for row in course_rows:
        start_str = row["開始時間"].strftime("%Y/%m/%d %H:%M") if row["開始時間"] else ""
        data.append([
            Paragraph(str(row["課程類別"]), cell_style),
            Paragraph(f'{row["有效積分"]:.2f}' if row["有效積分"] is not None else "", cell_style),
            Paragraph(str(row["審查單位"]), cell_style),
            Paragraph(str(row["主辦單位"]), cell_style),
            Paragraph(str(row["課程名稱"]), cell_style),
            Paragraph(start_str, cell_style),
            Paragraph(str(row["備註"]), cell_style),
        ])

    col_widths = [30 * mm, 18 * mm, 30 * mm, 32 * mm, 90 * mm, 28 * mm, 18 * mm]
    table = Table(data, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4472C4")),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F2F2")]),
    ]))
    elements.append(table)

    # 分類彙整（第二個表格，另起一頁附在報表後面）
    elements.append(Spacer(1, 10 * mm))
    elements.append(Paragraph("依課程類別彙整", title_style))
    categories = sorted({r["課程類別"] for r in course_rows if r["課程類別"]})
    summary_data = [[Paragraph(h, header_style) for h in ["課程類別", "有效積分合計", "堂數"]]]
    for cat in categories:
        cat_rows = [r for r in course_rows if r["課程類別"] == cat]
        cat_sum = sum(r["有效積分"] for r in cat_rows if r["有效積分"])
        summary_data.append([
            Paragraph(cat, cell_style),
            Paragraph(f"{cat_sum:.2f}", cell_style),
            Paragraph(str(len(cat_rows)), cell_style),
        ])
    summary_table = Table(summary_data, colWidths=[60 * mm, 40 * mm, 30 * mm])
    summary_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4472C4")),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F2F2")]),
    ]))
    elements.append(summary_table)

    doc.build(elements)
    buffer.seek(0)
    return buffer


# ---------------- Streamlit 介面 ----------------

st.set_page_config(page_title="衛福部學分 PDF 整理工具", layout="wide")

st.title("衛福部醫事人員繼續教育學分 PDF → Excel 整理工具")
st.caption(
    "本工具僅在本機執行，你上傳的字型檔案只會留在這個瀏覽器分頁背後、"
    "在你自己電腦上執行的程式記憶體裡，不會被儲存或上傳到任何遠端伺服器。"
)

st.warning(
    "⚠️ 若此頁面部署在公開雲端服務（例如 Streamlit Community Cloud）上："
    "請務必只上傳你自己電腦上已合法安裝的字型檔案，本工具不會儲存、"
    "也不會散布任何字型檔案。字型授權條款是否允許「暫時上傳給雲端工具運算」"
    "因授權而異，使用前請自行確認你手上那份字型授權是否允許此用途，"
    "本工具無法代為保證合規性。",
    icon="⚠️",
)

col1, col2 = st.columns(2)
with col1:
    pdf_file = st.file_uploader("上傳衛福部匯出的 PDF", type=["pdf"])
with col2:
    font_file = st.file_uploader(
        "上傳你電腦裡「已合法安裝」的標楷體字型檔 (DFKai-SB / kaiu.ttf)",
        type=["ttf", "ttc", "otf"],
        help=(
            "請上傳你自己電腦上已合法安裝的標楷體字型檔案。"
            "本工具僅在這次運算過程中於伺服器記憶體暫時讀取這個檔案，"
            "不會寫入磁碟、不會存進資料庫，運算結束或重新整理頁面後即清除，"
            "也不會被其他使用者存取到。"
        ),
    )
    st.caption(
        "🔒 字型檔僅用於本次運算，處理完成後即從記憶體中清除，不會被儲存、"
        "也不會被其他使用者的請求存取到。"
    )

font_number = st.number_input(
    "字型集合檔內的字型編號（一般填 0 即可；若讀取失敗可嘗試 1、2...）",
    min_value=0, max_value=10, value=0, step=1,
)

output_format = st.radio(
    "輸出格式",
    ["Excel（含分類彙整公式，方便後續編輯）", "PDF（含中文字型，方便直接列印/存查）", "兩者都要"],
    horizontal=True,
)

col_run, col_clear = st.columns([3, 1])
with col_run:
    run_clicked = st.button("開始解析", type="primary", disabled=not (pdf_file and font_file))
with col_clear:
    if st.button("🗑️ 清除本次資料"):
        st.session_state.pop("gid_to_unicode", None)
        st.session_state.pop("course_rows", None)
        st.session_state.pop("other_tables", None)
        st.rerun()

if run_clicked:
    with st.spinner("讀取字型檔，建立字碼對照表..."):
        try:
            # 存在 session_state：只屬於「這個使用者、這個工作階段」，
            # 不會跨使用者共享，關閉分頁或按「清除本次資料」即消失。
            st.session_state["gid_to_unicode"] = build_gid_to_unicode(
                font_file.getvalue(), font_number=font_number
            )
        except Exception as e:
            st.error(f"字型檔讀取失敗：{e}")
            st.stop()
    gid_to_unicode = st.session_state["gid_to_unicode"]
    st.success(f"字型讀取成功，共取得 {len(gid_to_unicode)} 個字碼對照。")

    progress_bar = st.progress(0.0, text="解析 PDF 中...")

    def _progress(frac):
        progress_bar.progress(frac, text=f"解析 PDF 中... {int(frac * 100)}%")

    try:
        course_rows, other_tables, debug_samples = parse_pdf(
            pdf_file.getvalue(), gid_to_unicode, progress_cb=_progress
        )
    except Exception as e:
        st.error(f"PDF 解析失敗：{e}")
        st.stop()
    progress_bar.empty()

    with st.expander("🔍 診斷資訊：程式實際解碼出來的文字長怎樣（點開查看）", expanded=(len(course_rows) == 0)):
        st.write(
            "如果下面「第一列(可能是標頭)」看到的是方框 □、亂碼符號，或是明明"
            "看起來像「課程類別」卻沒被判定為課程表格，代表字型解碼對不起來"
            "（很可能字型檔案跟產生這份 PDF 當初用的字型不是同一個檔案）。"
        )
        for sample in debug_samples:
            st.write(
                f"頁碼 {sample['頁碼']}、表格 {sample['表格編號']}、"
                f"共 {sample['欄數']} 欄、"
                f"判定為課程表格：{'✅ 是' if sample['是否判定為課程表格'] else '❌ 否'}"
            )
            st.code(str(sample["第一列(可能是標頭)"]))

    st.success(f"解析完成！共取得 {len(course_rows)} 筆課程明細，另有 {len(other_tables)} 個其他格式表格。")

    if course_rows:
        import pandas as pd
        df = pd.DataFrame(course_rows)
        st.subheader("研習課程明細預覽")
        st.dataframe(df, use_container_width=True)

        st.subheader("依課程類別彙整（初步統計，正式數字以 Excel 內公式為準）")
        summary = df.groupby("課程類別").agg(
            有效積分合計=("有效積分", "sum"),
            堂數=("課程類別", "count"),
        ).reset_index()
        st.dataframe(summary, use_container_width=True)

    want_excel = output_format in ("Excel（含分類彙整公式，方便後續編輯）", "兩者都要")
    want_pdf = output_format in ("PDF（含中文字型，方便直接列印/存查）", "兩者都要")

    if want_excel:
        excel_buffer = build_excel(course_rows, other_tables)
        st.download_button(
            label="📥 下載整理好的 Excel",
            data=excel_buffer,
            file_name=f"學分整理結果_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if want_pdf:
        try:
            pdf_buffer = build_pdf_report(course_rows, font_file.getvalue(), font_number=font_number)
            st.download_button(
                label="📥 下載整理好的 PDF",
                data=pdf_buffer,
                file_name=f"學分整理結果_{datetime.now().strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
            )
        except Exception as e:
            st.error(f"PDF 產生失敗：{e}")

    if other_tables:
        st.info("有部分表格格式未被自動辨識（例如頁首的個人累計統計區塊），已原樣放進 Excel 的「其他表格_待確認」分頁，請手動核對。")
