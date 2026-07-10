#!/usr/bin/env python3
"""Convert the Chinese underperformance Markdown report to a professional PDF.

Workflow:
1. Read the CN Markdown report.
2. Convert common Traditional Chinese characters/phrases to Simplified Chinese.
3. Render Markdown to a McKinsey-style HTML report.
4. Use LibreOffice headless to convert HTML to PDF.
5. Validate the generated PDF with pypdf.
"""
from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import markdown
from pypdf import PdfReader
from bs4 import BeautifulSoup, NavigableString

try:
    import fitz  # PyMuPDF
except ImportError:  # pragma: no cover - handled at runtime
    fitz = None

ROOT = Path(__file__).resolve().parent
DEFAULT_INPUT = ROOT / "PnL_OCI_2026_A_HK_Underperformance_Analysis_CN.md"
DEFAULT_OUTPUT = ROOT / "PnL_OCI_2026_A_HK_Underperformance_Analysis_CN.pdf"
DEFAULT_HTML = ROOT / "PnL_OCI_2026_A_HK_Underperformance_Analysis_CN_print.html"
SOFFICE = Path("C:/Program Files/LibreOffice/program/soffice.exe")
EDGE = Path("C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe")

# A compact fallback map for common Traditional characters and finance/reporting words.
# If opencc/zhconv is installed later, the script will prefer it automatically.
TRAD_TO_SIMP_PHRASES = {
    "??": "??",
    "??": "??",
    "??": "??",
    "??": "??",
    "??": "??",
    "??": "??",
    "??": "??",
    "??": "??",
    "??": "??",
    "??": "??",
    "??": "??",
    "??": "??",
    "??": "??",
    "??": "??",
    "??": "??",

    "麥肯錫": "麦肯锡",
    "臺": "台",
    "台灣": "台湾",
    "港幣": "港币",
    "人民幣": "人民币",
    "匯率": "汇率",
    "風險": "风险",
    "風格": "风格",
    "收益": "收益",
    "虧損": "亏损",
    "實現": "实现",
    "未實現": "未实现",
    "股息": "股息",
    "對沖": "对冲",
    "報告": "报告",
    "數據": "数据",
    "歸因": "归因",
    "組合": "组合",
    "市場": "市场",
    "銀行": "银行",
    "高質量": "高质量",
    "持倉": "持仓",
    "權重": "权重",
    "凈值": "净值",
    "淨值": "净值",
    "凈損益": "净损益",
    "淨損益": "净损益",
    "買": "买",
    "賣": "卖",
    "價": "价",
    "獲利": "获利",
    "動量": "动量",
    "長期": "长期",
    "轉": "转",
    "業": "业",
    "壓": "压",
    "穩": "稳",
    "體": "体",
    "現": "现",
    "觀": "观",
    "與": "与",
    "為": "为",
    "於": "于",
    "內": "内",
    "這": "这",
    "個": "个",
    "來": "来",
    "時": "时",
    "後": "后",
    "將": "将",
    "應": "应",
    "讓": "让",
    "屬": "属",
    "單": "单",
    "雙": "双",
    "開": "开",
    "關": "关",
    "東": "东",
    "龍": "龙",
    "藥": "药",
    "鐵": "铁",
    "塔": "塔",
    "滬": "沪",
    "甬": "甬",
    "聯": "联",
    "通": "通",
    "國": "国",
    "動": "动",
    "產": "产",
    "質": "质",
    "務": "务",
    "優": "优",
    "輸": "输",
    "贏": "赢",
    "較": "较",
    "擇": "择",
    "擴": "扩",
    "趨": "趋",
    "線": "线",
    "週": "周",
    "萬": "万",
}

TRAD_TO_SIMP_CHARS = str.maketrans({
    "麥": "麦", "錫": "锡", "與": "与", "為": "为", "匯": "汇", "幣": "币", "風": "风",
    "險": "险", "虧": "亏", "實": "实", "現": "现", "對": "对", "沖": "冲", "報": "报",
    "數": "数", "據": "据", "歸": "归", "組": "组", "閤": "合", "場": "场", "銀": "银",
    "質": "质", "倉": "仓", "權": "权", "淨": "净", "凈": "净", "損": "损", "買": "买",
    "賣": "卖", "價": "价", "獲": "获", "動": "动", "長": "长", "轉": "转", "業": "业",
    "壓": "压", "穩": "稳", "體": "体", "觀": "观", "於": "于", "內": "内", "這": "这",
    "個": "个", "來": "来", "時": "时", "後": "后", "將": "将", "應": "应", "讓": "让",
    "屬": "属", "單": "单", "雙": "双", "開": "开", "關": "关", "東": "东", "龍": "龙",
    "藥": "药", "鐵": "铁", "滬": "沪", "聯": "联", "國": "国", "產": "产", "務": "务",
    "優": "优", "輸": "输", "贏": "赢", "較": "较", "擇": "择", "擴": "扩", "趨": "趋",
    "線": "线", "週": "周", "萬": "万",
})

HTML_TEMPLATE = """<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>{title}</title>
<style>
@page {{ size: A4; margin: 13mm 12mm 15mm 12mm; }}
* {{ box-sizing: border-box; }}
body {{
  margin: 0;
  color: #111827;
  background: #ffffff;
  font-family: "Microsoft YaHei", "Noto Sans CJK SC", "Noto Sans SC", "SimHei", Arial, sans-serif;
  font-size: 9.2pt;
  line-height: 1.5;
}}
.report-shell {{ max-width: 1060px; margin: 0 auto; }}
.cover {{
  position: relative;
  border-top: 8px solid #061a33;
  border-bottom: 1px solid #d8dee9;
  padding: 18px 0 15px 0;
  margin-bottom: 14px;
}}
.cover:after {{
  content: "";
  display: block;
  width: 94px;
  height: 5px;
  margin-top: 12px;
  background: #2f6fed;
}}
.kicker {{
  color: #6b7280;
  text-transform: uppercase;
  letter-spacing: .16em;
  font-size: 7.8pt;
  font-weight: 800;
}}
h1 {{
  margin: 8px 0 6px 0;
  color: #061a33;
  font-size: 21pt;
  line-height: 1.14;
  font-weight: 850;
}}
.meta {{ color: #4b5563; font-size: 8.6pt; }}
h2 {{
  margin: 18px 0 9px;
  padding: 8px 11px;
  color: #ffffff;
  background: #061a33;
  border-left: 6px solid #2f6fed;
  font-size: 12.8pt;
  line-height: 1.25;
  font-weight: 800;
  page-break-after: avoid;
}}
h3 {{
  margin: 14px 0 7px;
  color: #0f2e5d;
  border-left: 4px solid #2f6fed;
  padding-left: 8px;
  font-size: 11.2pt;
  font-weight: 800;
  page-break-after: avoid;
}}
p {{ margin: 6px 0; }}
ul, ol {{ margin: 6px 0 9px 18px; padding: 0; }}
li {{ margin: 3px 0; }}
strong {{ color: #071d39; font-weight: 850; }}
code {{
  font-family: Consolas, "Microsoft YaHei", monospace;
  color: #0f2e5d;
  background: #eef4ff;
  padding: 1px 4px;
  border-radius: 3px;
}}
/* Top-headline block styled like a consulting executive takeaways page. */
#?? + ol,
#?????? + ol {{
  margin: 8px 0 16px 0;
  padding: 0;
  list-style: none;
  counter-reset: headline;
}}
#?? + ol > li,
#?????? + ol > li {{
  counter-increment: headline;
  position: relative;
  margin: 8px 0;
  padding: 9px 11px 9px 42px;
  border: 1px solid #d6dce8;
  border-left: 5px solid #2f6fed;
  background: #f7f9fc;
  page-break-inside: avoid;
}}
#?? + ol > li:before,
#?????? + ol > li:before {{
  content: counter(headline);
  position: absolute;
  left: 10px;
  top: 10px;
  width: 21px;
  height: 21px;
  line-height: 21px;
  text-align: center;
  border-radius: 50%;
  color: #ffffff;
  background: #061a33;
  font-weight: 850;
  font-size: 8pt;
}}
#?? + ol > li:nth-child(1),
#?????? + ol > li:nth-child(1) {{ border-left-color: #c00000; }}
#?? + ol > li:nth-child(2),
#?????? + ol > li:nth-child(2) {{ border-left-color: #c77800; }}
#?? + ol > li:nth-child(3),
#?????? + ol > li:nth-child(3) {{ border-left-color: #00843d; }}
#?? + ol > li:nth-child(1):before,
#?????? + ol > li:nth-child(1):before {{ background: #c00000; }}
#?? + ol > li:nth-child(2):before,
#?????? + ol > li:nth-child(2):before {{ background: #c77800; }}
#?? + ol > li:nth-child(3):before,
#?????? + ol > li:nth-child(3):before {{ background: #00843d; }}
table {{
  width: 100%;
  border-collapse: collapse;
  margin: 8px 0 13px;
  font-size: 7.55pt;
  table-layout: auto;
  page-break-inside: auto;
}}
thead {{ display: table-header-group; }}
tr {{ page-break-inside: avoid; page-break-after: auto; }}
th {{
  background: #061a33;
  color: white;
  font-weight: 800;
  text-align: left;
  border: 1px solid #061a33;
  padding: 4px 5px;
  vertical-align: top;
}}
td {{
  border: 1px solid #d6dce8;
  padding: 3px 5px;
  vertical-align: top;
  word-break: keep-all;
}}
tbody tr:nth-child(even) td {{ background: #f7f9fc; }}
img {{
  display: block;
  max-width: 100%;
  height: auto;
  margin: 8px auto 14px;
  page-break-inside: avoid;
}}
img.leadership-page {{
  max-width: 100%;
  max-height: 210mm;
  width: auto;
  height: auto;
  object-fit: contain;
  page-break-before: always;
  page-break-after: always;
}}
hr {{ border: 0; border-top: 1px solid #d8dee9; margin: 14px 0; }}
blockquote {{
  margin: 8px 0;
  border-left: 4px solid #2f6fed;
  padding: 4px 10px;
  color: #374151;
  background: #f5f8ff;
}}
.num-pos {{ color: #c00000 !important; font-weight: 800; }}
.num-neg {{ color: #00843d !important; font-weight: 800; }}
</style>
</head>
<body>
<div class="report-shell">
  <section class="cover">
    <h1>{title}</h1>
  </section>
  {body}
</div>
</body>
</html>
"""


def simplify_chinese(text: str) -> str:
    try:
        from opencc import OpenCC  # type: ignore
        return OpenCC("t2s").convert(text)
    except Exception:
        pass
    try:
        from zhconv import convert  # type: ignore
        return convert(text, "zh-cn")
    except Exception:
        pass
    for trad, simp in sorted(TRAD_TO_SIMP_PHRASES.items(), key=lambda item: len(item[0]), reverse=True):
        text = text.replace(trad, simp)
    text = text.translate(TRAD_TO_SIMP_CHARS)
    # Unicode-safe final pass for common characters that can survive if the
    # source file or shell mangles literal CJK mappings.
    final_map = {
        "沒": "没",  # ? -> ?
        "採": "采",  # ? -> ?
        "衝": "冲",  # ? -> ?
        "對": "对",  # ? -> ?
        "風": "风",  # ? -> ?
        "險": "险",  # ? -> ?
        "體": "体",  # ? -> ?
        "價": "价",  # ? -> ?
        "選": "选",  # ? -> ?
    }
    for trad, simp in final_map.items():
        text = text.replace(trad, simp)
    return text


def normalize_markdown(text: str) -> str:
    # Avoid extremely long unbreakable inline code/table content where possible.
    text = text.replace("\u2014", "\u2014")
    # Localise standalone half-year labels only. Do not touch futures contract
    # codes such as IH2603 / IH2609.
    text = re.sub(r"(?<![A-Za-z0-9])H1(?![A-Za-z0-9])", "\u4e0a\u534a\u5e74", text)
    text = re.sub(r"(?<![A-Za-z0-9])H2(?![A-Za-z0-9])", "\u4e0b\u534a\u5e74", text)
    return text


NUMBER_PATTERN = re.compile(
    r"(?<![\w/])([+-]?\d[\d,]*(?:\.\d+)?%?|[+-]?\.\d+%?)(?![\w/])"
)


def _number_sign(token: str) -> int:
    raw = token.replace(",", "").replace("%", "")
    try:
        value = float(raw)
    except ValueError:
        return 0
    if value > 0:
        return 1
    if value < 0:
        return -1
    return 0


def colorize_numbers(html: str) -> str:
    """Color numeric values red if positive and green if negative.

    Chinese finance convention commonly uses red for gains/positive values and
    green for losses/negative values. This operates on rendered HTML so the
    Markdown source stays clean.
    """
    soup = BeautifulSoup(html, "html.parser")
    skip_tags = {"style", "script", "code", "pre"}

    def replace_text_node(node: NavigableString) -> None:
        parent = node.parent
        if parent is None or parent.name in skip_tags:
            return
        text = str(node)
        if not NUMBER_PATTERN.search(text):
            return
        pieces = []
        last = 0
        for match in NUMBER_PATTERN.finditer(text):
            token = match.group(1)
            sign = _number_sign(token)
            if sign == 0:
                continue
            if match.start() > last:
                pieces.append(NavigableString(text[last:match.start()]))
            span = soup.new_tag("span")
            if sign > 0:
                span["class"] = "num-pos"
                span["style"] = "color:#c00000 !important; font-weight:800;"
                font = soup.new_tag("font", color="#c00000")
            else:
                span["class"] = "num-neg"
                span["style"] = "color:#00843d !important; font-weight:800;"
                font = soup.new_tag("font", color="#00843d")
            bold = soup.new_tag("b")
            bold.string = token
            font.append(bold)
            span.append(font)
            pieces.append(span)
            last = match.end()
        if not pieces:
            return
        if last < len(text):
            pieces.append(NavigableString(text[last:]))
        node.replace_with(*pieces)

    for node in list(soup.find_all(string=True)):
        replace_text_node(node)
    return str(soup)




def style_top_headlines(html: str) -> str:
    """Apply inline card styles to the executive headline lists.

    The source Markdown sometimes renders the three headlines as three separate
    one-item ordered lists because each headline has a paragraph underneath.
    Style the first three OL blocks after the first H2 for robust PDF output.
    """
    soup = BeautifulSoup(html, "html.parser")
    first_h2 = soup.find("h2")
    if first_h2 is None:
        return str(soup)
    colors = ["#c00000", "#c77800", "#00843d"]
    styled = 0
    for sibling in first_h2.find_next_siblings():
        if getattr(sibling, "name", None) == "h2":
            break
        if getattr(sibling, "name", None) != "ol":
            continue
        color = colors[min(styled, len(colors) - 1)]
        sibling["style"] = "margin:8px 0 8px 0; padding:0; list-style:none;"
        for li in sibling.find_all("li", recursive=False):
            li["style"] = (
                "list-style:none; margin:8px 0; padding:9px 11px; "
                f"border:1px solid #d6dce8; border-left:6px solid {color}; "
                "background-color:#f7f9fc; page-break-inside:avoid;"
            )
            badge = soup.new_tag("span")
            badge["style"] = (
                f"display:inline-block; background-color:{color}; color:#ffffff; "
                "font-weight:800; width:20px; height:20px; line-height:20px; "
                "text-align:center; margin-right:8px; border-radius:10px;"
            )
            badge.string = str(styled + 1)
            li.insert(0, badge)
        styled += 1
        if styled >= 3:
            break
    return str(soup)


def style_report_images(html: str) -> str:
    """Tag large appendix images so they fit vertically on A4 pages."""
    soup = BeautifulSoup(html, "html.parser")
    for img in soup.find_all("img"):
        src = str(img.get("src", ""))
        if "leadership_page_12_fitz.png" in src or "leadership_page_13_fitz.png" in src:
            existing = img.get("class", [])
            if isinstance(existing, str):
                existing = existing.split()
            img["class"] = list(dict.fromkeys([*existing, "leadership-page"]))
    return str(soup)


def render_html(markdown_text: str) -> str:
    md = markdown.Markdown(extensions=["extra", "tables", "sane_lists", "toc"])
    title_match = re.search(r"^#\s+(.+)$", markdown_text, flags=re.MULTILINE)
    title = title_match.group(1).strip() if title_match else "PnL OCI 2026 ????"
    # The professional cover already renders the report title. Remove the first
    # Markdown H1 from the body to avoid a duplicated title in the PDF.
    body_markdown = re.sub(r"^#\s+.+(?:\r?\n)+", "", markdown_text, count=1, flags=re.MULTILINE)
    body = style_report_images(style_top_headlines(colorize_numbers(md.convert(body_markdown))))
    return HTML_TEMPLATE.format(title=title, body=body)


def convert_html_to_pdf(html_path: Path, output_pdf: Path) -> None:
    """Convert HTML to PDF.

    Prefer Microsoft Edge headless because it preserves modern HTML/CSS much
    better than LibreOffice's HTML importer. Keep LibreOffice as a fallback for
    machines without Edge.
    """
    if output_pdf.exists():
        output_pdf.unlink()
    file_url = "file:///" + str(html_path.resolve()).replace("\\", "/")

    if EDGE.exists():
        cmd = [
            str(EDGE),
            "--headless",
            "--disable-gpu",
            "--no-first-run",
            "--disable-extensions",
            "--no-pdf-header-footer",
            f"--print-to-pdf={output_pdf}",
            file_url,
        ]
        proc = subprocess.run(cmd, text=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, check=False)
        if output_pdf.exists() and output_pdf.stat().st_size > 10_000:
            return
        # Continue to fallback only if Edge failed to produce a usable PDF.
        edge_output = proc.stdout
    else:
        edge_output = "Edge executable not found"

    if not SOFFICE.exists():
        raise FileNotFoundError(f"No PDF engine worked. Edge: {EDGE}; LibreOffice: {SOFFICE}. Edge output:\n{edge_output}")
    outdir = output_pdf.parent
    cmd = [str(SOFFICE), "--headless", "--convert-to", "pdf", "--outdir", str(outdir), str(html_path)]
    proc = subprocess.run(cmd, text=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, check=False)
    if proc.returncode != 0:
        raise RuntimeError(f"LibreOffice PDF conversion failed ({proc.returncode}):\n{proc.stdout}\nEdge output:\n{edge_output}")
    generated = outdir / f"{html_path.stem}.pdf"
    if generated.resolve() != output_pdf.resolve():
        if output_pdf.exists():
            output_pdf.unlink()
        shutil.move(str(generated), str(output_pdf))


def validate_pdf(pdf_path: Path) -> None:
    if not pdf_path.exists() or pdf_path.stat().st_size < 10_000:
        raise RuntimeError(f"PDF missing or too small: {pdf_path}")
    reader = PdfReader(str(pdf_path))
    pages = len(reader.pages)
    if pages < 1:
        raise RuntimeError("PDF has no pages")
    # LibreOffice PDFs with CJK fonts can extract as mojibake in pypdf, so validate
    # PDF structure here and validate simplified/source text before conversion.
    print(f"PDF validation passed: {pdf_path} ({pages} pages, {pdf_path.stat().st_size:,} bytes)")


def add_page_numbers(pdf_path: Path) -> None:
    """Stamp page numbers onto the generated PDF.

    Edge's command-line PDF printing does not expose reliable header/footer
    templates, so stamp the final PDF directly. Coordinates use PDF points.
    """
    if fitz is None:
        raise RuntimeError("PyMuPDF is required to add page numbers but is not installed")

    doc = fitz.open(pdf_path)
    total_pages = len(doc)
    for index, page in enumerate(doc, start=1):
        rect = page.rect
        text = f"{index} / {total_pages}"
        page.insert_textbox(
            fitz.Rect(rect.width - 95, rect.height - 30, rect.width - 30, rect.height - 15),
            text,
            fontsize=7.5,
            fontname="helv",
            color=(0.35, 0.39, 0.45),
            align=fitz.TEXT_ALIGN_RIGHT,
        )

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf", dir=pdf_path.parent) as tmp:
        temp_path = Path(tmp.name)
    try:
        doc.save(temp_path, garbage=4, deflate=True, clean=True)
        doc.close()
        temp_path.replace(pdf_path)
    finally:
        if temp_path.exists():
            temp_path.unlink()


def main() -> int:
    parser = argparse.ArgumentParser(description="Convert CN underperformance Markdown to a professional PDF.")
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--html", type=Path, default=DEFAULT_HTML)
    args = parser.parse_args()

    md_text = args.input.read_text(encoding="utf-8")
    simplified = simplify_chinese(normalize_markdown(md_text))
    traditional_guard = "".join(chr(cp) for cp in [27794, 25505, 34909, 33274, 28771, 39636, 20729, 29694, 23526, 23565, 33287, 28858, 21295, 24163, 39080, 38570, 34407, 25613, 27512, 32068, 37504, 36074, 20489, 27402, 28136, 36023, 36067, 29554, 21205, 38263, 36681, 26989, 22739, 31337, 35264, 26044, 20839, 36889, 20491, 20358, 26178, 24460, 23559, 25033, 35731, 23660, 21934, 38617, 38283, 38364, 26481, 40845, 34277, 37941, 28396, 32879, 22283, 29986, 21209, 20778, 36664, 36111, 36611, 25799, 25844, 36264, 32218, 36913, 33836, 33288, 33775, 24291, 20497, 40636, 26371, 40670, 32147, 30332, 38989, 36984, 27161, 30435, 35413, 35442, 37002, 35023, 35041])
    residual_traditional = [ch for ch in sorted(set(simplified)) if ch in traditional_guard]
    if residual_traditional:
        raise RuntimeError(f"Simplified Chinese validation failed; residual chars: {residual_traditional}")
    args.input.write_text(simplified, encoding="utf-8")

    html = render_html(simplified)
    args.html.write_text(html, encoding="utf-8")
    convert_html_to_pdf(args.html, args.output)
    add_page_numbers(args.output)
    validate_pdf(args.output)
    print(f"HTML written: {args.html}")
    print(f"PDF written:  {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())