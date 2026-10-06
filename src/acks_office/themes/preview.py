"""主题样张（theme preview）：一页 HTML 样张（色板、对比度、字体、字号、各格式的版式示意），
外加用同一份示例内容生成的 Word、PDF、PPT、Excel。

HTML 不依赖任何外部资源，任何 Agent 都能直接展示；示例文件可以转成图片核对实际效果。
"""

import os
from html import escape
from pathlib import Path
from typing import Any, Dict, List, Sequence

from .loader import load_theme
from .model import Theme
from .validate import validate_theme

FORMATS = ("html", "docx", "pdf", "pptx", "xlsx")

SAMPLE_REPORT = """---
title: 2026 年第三季度\\n经营回顾
short_title: 经营回顾
kicker: 经营分析报告
subtitle: 营收、门店与会员的季度表现
publication: 季度刊
issue: "07"
season: Autumn 2026
lede: 营收增长 24%，36 家新店开业，61% 的会员回来续杯——这是一份关于耐心的季报。
date: 2026 年 10 月 8 日
author: 经营分析部
version: v1.0
---

# 经营概览 {label="生意 · THE BUSINESS"}

> 营收增长 24%，毛利率回落 1.4 个百分点：增长和成本都在提醒我们慢一点。

第三季度营收 **1.28 亿元**，同比增长 24.0%。增长主要来自华东和西南：华东的新店带动营收增长 25.8%，西南体量最小，增速却达到 28.0%。Revenue grew 24% year over year.

## 分区域营收

表：分区域营收（单位：万元）

| 区域 | 营收 | 同比 | 占比 |
|---|--:|--:|--:|
| 华东 | 5,888 | +25.8% | 46.0% |
| 华南 | 3,456 | +22.0% | 27.0% |
| 华北 | 2,304 | +20.5% | 18.0% |
| 西南 | 1,152 | +28.0% | 9.0% |
| 合计 | 12,800 | +24.0% | 100.0% |

> [!TIP]
> 毛利率回落主要来自咖啡豆采购价上涨，第四季度起执行年度采购合同，预计毛利率回到 59% 以上。

# 门店与会员 {label="门店 · STORES"}

本季度新开门店 36 家，其中 21 家位于华东。会员复购率为 61%，提高 3.2 个百分点。

- 华东加密：杭州、苏州优先
- 稳住成本：执行年度采购合同
- 会员升级：上线积分商城

> 增长不靠打折，靠的是有人愿意再来一次。

```text
营收 = 门店数 × 店均营收
```
"""

SAMPLE_DECK: List[Dict[str, Any]] = [
    {"layout": "title", "kicker": "2026 年第三季度", "title": "经营回顾", "subtitle": "营收、门店与会员：一个季度的关键变化",
     "footer": "经营分析部 · 2026 年 10 月",
     "kpis": [{"label": "营收", "value": "1.28", "unit": "亿元"}, {"label": "同比增长", "value": "+24.0%", "highlight": True},
              {"label": "新开门店", "value": "36", "unit": "家"}, {"label": "会员复购率", "value": "61%"}]},
    {"layout": "section", "number": "01", "kicker": "生意", "title": "一个季度的账本", "title_en": "The Business"},
    {"layout": "data", "title": "营收增长 24%，华东贡献近一半",
     "subtitle": "第三季度营收 1.28 亿元；毛利率 58.2%，较去年同期回落 1.4 个百分点。",
     "kpis": [{"label": "营收", "value": "1.28", "unit": "亿元", "delta": "同比 +24.0%"},
              {"label": "新开门店", "value": "36", "unit": "家", "delta": "同比多开 12 家"},
              {"label": "会员复购率", "value": "61", "unit": "%", "delta": "提高 3.2 个百分点"}],
     "chart": {"title": "分区域营收占比", "unit": "单位：%", "data": [["华东", 46.0], ["华南", 27.0], ["华北", 18.0], ["西南", 9.0]],
               "highlight": 0}},
    {"layout": "content", "title": "会员运营的三个抓手", "subtitle": "复购率从 61% 提升到 65%",
     "bullets": ["积分商城：12 月上线，积分可兑换新品试饮", "会员日：每月第一个周三，全场第二杯半价",
                 "门店记住你：店员可在收银系统里看到常点饮品"]},
    {"layout": "table", "title": "第四季度三件事",
     "table": {"header": ["序号", "事项", "目标", "负责"],
               "rows": [["01", "华东加密", "新开 20 家门店，集中在杭州、苏州", "拓展部"],
                        ["02", "稳住成本", "执行年度采购合同，毛利率回到 59% 以上", "供应链部"],
                        ["03", "会员升级", "上线积分商城，复购率提升到 65%", "会员运营部"]]}},
    {"layout": "number", "label": "会员复购率", "value": "61", "unit": "%", "delta_label": "同比", "delta": "+3.2 pt",
     "text": "每十位会员里，有六位在一个月内回来过。"},
    {"layout": "quote", "quote": "增长不靠打折，靠的是有人愿意再来一次。", "author": "门店店长", "role": "杭州湖滨店"},
]

SAMPLE_SHEETS: List[Dict[str, Any]] = [
    {"name": "Q3 分区域", "title": "Q3 分区域营收（万元）",
     "header": ["区域", "7 月", "8 月", "9 月", "Q3 合计", "同比", "毛利率变化"],
     "rows": [["华东", 1846, 1962, 2080, 5888, 0.258, -1.3], ["华南", 1092, 1148, 1216, 3456, 0.220, -1.9],
              ["华北", 742, 768, 794, 2304, 0.205, -1.6], ["西南", 356, 384, 412, 1152, 0.280, -0.2],
              ["合计", 4036, 4262, 4502, 12800, 0.240, -1.4]],
     "formats": {"同比": "+0.0%;-0.0%", "毛利率变化": "#,##0.0;(#,##0.0)"}, "highlight": [[0, 4]],
     "note": "注：毛利率变化单位为百分点。"},
    {"name": "月度明细", "title": "月度营收明细（万元）", "header": ["月份", "营收", "成本", "毛利"],
     "rows": [["7 月", 4036, 1654.2, 2381.8], ["8 月", 4262, 1790.0, 2472.0], ["9 月", 4502, 1885.1, 2616.9]],
     "total": True, "chart": {"values": [1, 3]}},
]


def build_preview(ref: Any, out_dir: str, formats: Sequence[str] = FORMATS) -> Dict[str, Any]:
    """生成样张到 out_dir：HTML 样张和各格式的示例文件。返回生成的文件、校验结果与提醒。"""
    theme = load_theme(ref)
    unknown = [f for f in formats if f not in FORMATS]
    if unknown:
        raise ValueError(f"不支持的样张格式：{', '.join(unknown)}；可选：{', '.join(FORMATS)}")
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    report = validate_theme(theme.path if theme.source == "path" else theme.name)
    files: Dict[str, str] = {}
    warnings: List[Dict[str, str]] = []
    stem = f"{theme.name}-sample"
    if "docx" in formats:
        from ..render.word import render_word
        files["docx"] = str(out / f"{stem}.docx")
        warnings += render_word(None, SAMPLE_REPORT, files["docx"], theme=theme).get("warnings", [])
    if "pdf" in formats:
        from ..render.pdf import render_pdf
        files["pdf"] = str(out / f"{stem}.pdf")
        warnings += render_pdf(None, SAMPLE_REPORT, files["pdf"], theme=theme).get("warnings", [])
    if "pptx" in formats:
        from ..render.slides import render_slides
        files["pptx"] = str(out / f"{stem}.pptx")
        warnings += render_slides("经营回顾", SAMPLE_DECK, files["pptx"], theme=theme,
                                  brand="栖木咖啡", publication="季度刊", issue="07").get("warnings", [])
    if "xlsx" in formats:
        from ..render.sheet import render_sheets
        files["xlsx"] = str(out / f"{stem}.xlsx")
        warnings += render_sheets("经营数据", files["xlsx"], sheets=SAMPLE_SHEETS, theme=theme).get("warnings", [])
    if "html" in formats:
        path = out / f"{theme.name}-specimen.html"
        samples = {kind: os.path.basename(p) for kind, p in files.items()}
        path.write_text(specimen_html(theme, report, samples), encoding="utf-8")
        files = {"html": str(path), **files}
    seen, unique = set(), []
    for item in warnings:
        key = (item.get("code"), item.get("message"))
        if key not in seen:
            seen.add(key)
            unique.append(item)
    return {"theme": theme.name, "dir": str(out), "files": files,
            "validation": {"ok": report["ok"], "errors": report["errors"], "warnings": report["warnings"]},
            "warnings": unique}


# ---------------------------------------------------------------- HTML 样张

def _px(pt: float) -> str:
    return f"{float(pt) * 4 / 3:.1f}px"


def _stack(theme: Theme, slot: str) -> str:
    """CSS 字体栈：西文在前（中文字形落到后面的中文字体上），各自带备选。"""
    names = theme.family_candidates(slot, "en") + theme.family_candidates(slot, "cn")
    generic = "serif" if slot in ("serif", "quote") or "serif" in " ".join(names).lower() else "sans-serif"
    if slot == "mono":
        generic = "monospace"
    # 用单引号：这些字体栈也会写进双引号包着的 style 属性
    return ", ".join("'" + escape(n.replace("'", ""), quote=False) + "'" for n in dict.fromkeys(names)) + f", {generic}"


def _role_css(theme: Theme, role: str) -> str:
    spec = theme.get(f"font.roles.{role}")
    return f"font-family: {_stack(theme, spec['family'])}; font-weight: {int(spec.get('weight', 400))};"


def specimen_html(theme: Theme, report: Dict[str, Any], samples: Dict[str, str]) -> str:
    c = theme.color
    doc, slide, sheet = theme.get("layout.doc"), theme.get("layout.slide"), theme.get("layout.sheet")
    t = lambda key: theme.get(f"type.doc.{key}")  # noqa: E731
    colors = {k: v for k, v in theme.get("color").items() if k != "chart" and isinstance(v, str)}
    swatches = "".join(
        f'<div class="swatch"><div class="chip" style="background:{escape(v)}"></div>'
        f'<b>{escape(k)}</b><code>{escape(v.upper())}</code></div>' for k, v in colors.items())
    chart = "".join(f'<span style="background:{escape(v)}"></span>' for v in theme.chart_colors())
    contrast = "".join(
        f'<tr><td>{escape(r["use"])}</td><td><span class="dot" style="background:{c(r["foreground"])}"></span>'
        f'{escape(r["foreground"])} / {escape(r["background"])}</td>'
        f'<td class="num">{r["ratio"]:.2f}</td><td class="num">{r["minimum"]:g}</td>'
        f'<td>{"通过" if r["ok"] else "<b class=bad>不足</b>"}</td></tr>' for r in report.get("contrast", []))
    font_rows = "".join(
        f'<tr><td>{escape(f["slot"])} · {"中文" if f["script"] == "cn" else "西文"}</td><td>{escape(f["family"])}</td>'
        f'<td>{"已安装" if f.get("installed") else ("改用 " + escape(f["substitute"]) if f.get("substitute") else "<b class=bad>未安装</b>")}</td>'
        f'<td style="font-family:{_stack(theme, f["slot"])}">文字设计 Typography 0123</td></tr>'
        for f in report.get("fonts", []))
    scale = "".join(
        f'<div class="scale" style="{_role_css(theme, role)} font-size:{_px(t(key))}; line-height:1.3">'
        f'<span class="meta">{escape(key)} · {t(key):g} pt</span>{escape(text)}</div>'
        for key, role, text in (("title", "title", "季度经营回顾"), ("h1", "heading", "一级标题 Heading"),
                                ("h2", "heading", "二级标题 Heading"), ("h3", "heading", "三级标题"),
                                ("body", "body", "正文：第三季度营收 1.28 亿元，同比增长 24.0%。"),
                                ("caption", "label", "表 1　分区域营收（单位：万元）")))
    issues = "".join(f'<li class="bad">{escape(e["path"])}：{escape(e["message"])}</li>' for e in report.get("errors", []))
    issues += "".join(f'<li>{escape(w["message"])}</li>' for w in report.get("warnings", []))
    links = "".join(f'<a href="{escape(name)}">{escape(kind.upper())}</a>' for kind, name in samples.items())
    parents = [name for name in theme.chain[1:] if not name.startswith("_")]  # 内部的 _base 不展示

    head_fill = doc["table"]["head"] == "fill"
    bar = doc["table"]["total"] == "bar"
    total_css = (f"border-top:2px solid {c(doc['table']['total_rule'])}" if bar else
                 f"border-top:1px solid {c(doc['table']['total_rule'])}; border-bottom:1px solid {c(doc['table']['total_rule'])}")
    head_css = (f"background:{c('surface')}; border-top:1px solid {c(doc['table']['head_rule'])}" if head_fill else
                f"border-bottom:2px solid {c(doc['table']['head_rule'])}")
    numbers = doc["heading_numbers"] and doc["chapter"] == "inline"
    callout = (f"background:{c('surface')}; padding:12px 14px" if doc["callout"] == "tint" else
               f"border-top:1px solid {c('rule_strong')}; border-bottom:1px solid {c('rule_strong')}; padding:12px 0")
    quote = (f'<div class="pull" style="border-color:{c("rule_strong")}"><span style="color:{c("accent")}; {_role_css(theme, "display")}">“</span>'
             f'<p style="{_role_css(theme, "quote")}">增长不靠打折，靠的是有人愿意再来一次。</p></div>'
             if doc["quote"] == "pull" else
             f'<blockquote style="border-left:3px solid {c("rule_mid")}; color:{c("muted")}; {_role_css(theme, "quote")}">'
             f'增长不靠打折，靠的是有人愿意再来一次。</blockquote>')
    cover = _cover_mock(theme)
    slides = _slide_mocks(theme)
    sheet_head = (f"background:{c('surface')}; font-weight:700" if sheet["head"] == "fill" else f"color:{c('muted')}")
    sheet_total = (f"border-top:2px solid {c(sheet['total_rule'])}" if sheet["total"] == "bar" else
                   f"border-top:1px solid {c(sheet['total_rule'])}; border-bottom:1px solid {c(sheet['total_rule'])}")
    return f"""<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>{escape(theme.title)} 主题样张</title>
<style>
:root {{ --ink:{c('ink')}; --muted:{c('muted')}; --rule:{c('rule')}; --accent:{c('accent')}; --surface:{c('surface')}; }}
* {{ box-sizing: border-box; }}
body {{ margin:0; background:#EDEDEA; color:var(--ink); {_role_css(theme, 'body')} font-size:15px; line-height:1.6; }}
.wrap {{ max-width:1080px; margin:0 auto; padding:40px 16px 64px; }}
section {{ background:#FFFFFF; border-radius:10px; padding:28px; margin-top:20px; overflow-x:auto; }}
h1 {{ {_role_css(theme, 'title')} font-size:34px; margin:0 0 6px; line-height:1.2; }}
h2 {{ {_role_css(theme, 'heading')} font-size:20px; margin:0 0 16px; }}
.lead {{ color:var(--muted); margin:0; }} .meta {{ display:block; font-size:12px; color:var(--muted); font-weight:400; }}
.badge {{ display:inline-block; margin-top:12px; padding:3px 10px; border-radius:999px; font-size:13px;
         background:{'#E7F4EC' if report.get('ok') else '#FBE9E7'}; color:{'#1E6B3A' if report.get('ok') else '#B42318'}; }}
.swatches {{ display:grid; grid-template-columns:repeat(auto-fill,minmax(150px,1fr)); gap:12px; }}
.swatch .chip {{ height:56px; border-radius:6px; border:1px solid rgba(0,0,0,.08); margin-bottom:6px; }}
.swatch b {{ display:block; font-size:13px; }} code {{ font-size:12px; color:var(--muted); }}
.chart {{ display:flex; height:16px; margin-top:16px; border-radius:4px; overflow:hidden; }} .chart span {{ flex:1; }}
table.info {{ width:100%; border-collapse:collapse; font-size:14px; }}
table.info td, table.info th {{ text-align:left; padding:7px 10px; border-bottom:1px solid #E5E7EB; vertical-align:top; }}
.num {{ text-align:right; font-variant-numeric:tabular-nums; }} .bad {{ color:#B42318; }}
.dot {{ display:inline-block; width:10px; height:10px; border-radius:50%; margin-right:6px; }}
.scale {{ padding:8px 0; border-bottom:1px solid #F0F0F0; }}
.doc {{ display:grid; grid-template-columns:minmax(0,260px) minmax(0,1fr); gap:28px; }}
.page {{ aspect-ratio:210/297; background:{c('paper')}; border:1px solid #DDD; position:relative; overflow:hidden; font-size:9px; }}
.excerpt h3 {{ {_role_css(theme, 'heading')} font-size:{_px(t('h1'))}; margin:0 0 8px; }}
.excerpt h4 {{ {_role_css(theme, 'heading')} font-size:{_px(t('h2'))}; margin:16px 0 6px; }}
.excerpt .n {{ color:var(--accent); margin-right:8px; }}
.excerpt p {{ margin:0 0 10px; font-size:{_px(t('body'))}; }}
.tbl {{ width:100%; border-collapse:collapse; font-size:{_px(t('table'))}; margin:6px 0 14px; }}
.tbl th {{ {_role_css(theme, 'label')} font-weight:400; color:var(--muted); font-size:{_px(t('table_head'))}; text-align:left; padding:6px 8px; {head_css}; }}
.tbl td {{ padding:6px 8px; border-bottom:1px solid {c(doc['table']['row_rule'])}; }}
.tbl td.r, .tbl th.r {{ text-align:right; }} .tbl tr.total td {{ font-weight:700; border-bottom:0; {total_css}; }}
.cap {{ {_role_css(theme, 'label')} font-weight:400; color:var(--muted); font-size:{_px(t('caption'))}; }}
.cap b {{ color:{c('accent') if doc.get('caption') == 'accent' else c('muted')}; font-weight:400; }}
.callout {{ {callout}; margin:6px 0 14px; }} .callout b {{ color:var(--accent); font-size:12px; display:block; }}
blockquote {{ margin:10px 0; padding:2px 0 2px 14px; }}
.pull {{ display:flex; gap:12px; border-top:1px solid; border-bottom:1px solid; padding:10px 0; margin:10px 0; }}
.pull span {{ font-size:44px; line-height:1; }} .pull p {{ margin:0; font-size:{_px(t('pull_quote'))}; }}
.slides {{ display:grid; grid-template-columns:repeat(auto-fill,minmax(280px,1fr)); gap:16px; }}
.slide {{ aspect-ratio:16/9; position:relative; overflow:hidden; border:1px solid #DDD; }}
.xl {{ border-collapse:collapse; font-size:13px; }} .xl td, .xl th {{ padding:6px 12px; text-align:right; }}
.xl td:first-child, .xl th:first-child {{ text-align:left; }}
.xl th {{ {sheet_head}; border-bottom:2px solid {c(sheet['head_rule'])}; font-weight:{'700' if sheet['head'] == 'fill' else '400'}; }}
.xl td {{ border-bottom:1px solid {c(sheet['row_rule'])}; }} .xl tr.total td {{ font-weight:700; border-bottom:0; {sheet_total}; }}
.xl .neg {{ color:{c(sheet['negative'])}; }} .xl .hl {{ outline:2px solid {c(sheet['highlight'])}; outline-offset:-2px; }}
.links a {{ display:inline-block; margin:0 10px 6px 0; padding:6px 14px; border-radius:6px; background:var(--surface);
           color:var(--ink); text-decoration:none; border:1px solid var(--rule); }}
ul.issues {{ margin:0; padding-left:20px; font-size:14px; }}
@media (max-width:720px) {{ .doc {{ grid-template-columns:1fr; }} section {{ padding:18px; }} }}
</style></head>
<body><div class="wrap">
<header>
<h1>{escape(theme.title)}</h1>
<p class="lead">{escape(theme.name)}{(" · 继承 " + escape(" → ".join(parents))) if parents else ""}</p>
<p class="lead">{escape(str(theme.meta.get("description", "")))}</p>
<span class="badge">{"校验通过" if report.get("ok") else f"{len(report.get('errors', []))} 处错误"} · {len(report.get("warnings", []))} 条提醒</span>
</header>
<section><h2>色板</h2><div class="swatches">{swatches}</div><div class="chart">{chart}</div></section>
<section><h2>对比度（WCAG AA）</h2><table class="info"><tr><th>用途</th><th>前景 / 背景</th><th class="num">比值</th><th class="num">最低</th><th></th></tr>{contrast}</table></section>
<section><h2>字体</h2><table class="info"><tr><th>槽位</th><th>主题字体</th><th>本机</th><th>示例</th></tr>{font_rows}</table></section>
<section><h2>字号（Word / PDF）</h2>{scale}</section>
<section><h2>文档版式</h2><div class="doc"><div>{cover}<p class="cap">封面：{escape(doc["cover"])} · 章节：{escape(doc["chapter"])} · {int(doc["columns"])} 栏</p></div>
<div class="excerpt">
<h3>{'<span class="n">1</span>' if numbers else ''}经营概览</h3>
<p>第三季度营收 <b>1.28 亿元</b>，同比增长 24.0%。增长主要来自华东和西南，西南增速达到 28.0%。</p>
<h4>{'<span class="n">1.1</span>' if numbers else ''}分区域营收</h4>
<div class="cap"><b>{escape(theme.label("table"))} 1</b>　分区域营收（单位：万元）</div>
<table class="tbl"><tr><th>区域</th><th class="r">营收</th><th class="r">同比</th></tr>
<tr><td>华东</td><td class="r">5,888</td><td class="r">+25.8%</td></tr>
<tr><td>华南</td><td class="r">3,456</td><td class="r">+22.0%</td></tr>
<tr class="total"><td>合计</td><td class="r">12,800</td><td class="r">+24.0%</td></tr></table>
<div class="callout"><b>{escape(theme.label("tip"))}</b>第四季度起执行年度采购合同，预计毛利率回到 59% 以上。</div>
{quote}
</div></div></section>
<section><h2>幻灯片版式</h2><div class="slides">{slides}</div>
<p class="cap">封面：{escape(slide["cover"])} · 章节页：{escape(slide["section"])} · 内容页：{escape(slide["content"])}</p></section>
<section><h2>Excel</h2><table class="xl">
<tr><th>区域</th><th>7 月</th><th>8 月</th><th>Q3 合计</th><th>毛利率变化</th></tr>
<tr><td>华东</td><td>1,846</td><td>1,962</td><td class="hl">5,888</td><td class="neg">(1.3)</td></tr>
<tr><td>华南</td><td>1,092</td><td>1,148</td><td>3,456</td><td class="neg">(1.9)</td></tr>
<tr class="total"><td>合计</td><td>2,938</td><td>3,110</td><td>9,344</td><td class="neg">(1.6)</td></tr></table></section>
{f'<section><h2>示例文件</h2><div class="links">{links}</div><p class="cap">同一份内容按这个主题生成的文件，可以直接打开或转成图片核对。</p></section>' if links else ''}
{f'<section><h2>校验提醒</h2><ul class="issues">{issues}</ul></section>' if issues else ''}
</div></body></html>
"""


def _cover_mock(theme: Theme) -> str:
    c, variant = theme.color, theme.get("layout.doc.cover")
    title = f'<div style="{_role_css(theme, "title")} font-size:20px; line-height:1.15; color:{c("ink")}">季度<br>经营回顾</div>'
    if variant == "issue":
        body = (f'<div style="position:absolute; top:8%; left:10%; right:10%; display:flex; justify-content:space-between; '
                f'color:{c("muted")}">季度刊<span>AUTUMN 2026</span></div>'
                f'<div style="position:absolute; top:18%; right:10%; color:{c("accent")}; {_role_css(theme, "display")} '
                f'font-size:64px; line-height:1">07</div>'
                f'<div style="position:absolute; top:45%; left:10%; right:10%">{title}'
                f'<div style="margin-top:8px; color:{c("muted")}">一份关于耐心的季报。</div>'
                f'<div style="margin-top:10px; width:45%; height:1.5px; background:{c("accent")}"></div></div>'
                f'<div style="position:absolute; bottom:7%; left:10%; right:10%; border-top:1px solid {c("rule")}; '
                f'color:{c("muted")}; padding-top:4px">本期目录 · 01 · 02 · 03</div>')
    elif variant == "standard":
        body = (f'<div style="position:absolute; top:8%; left:10%; right:10%; display:flex; justify-content:space-between">'
                f'<b>品牌</b><span style="color:{c("muted")}">内部资料</span></div>'
                f'<div style="position:absolute; top:38%; left:10%; right:10%"><div style="width:22px; height:3px; '
                f'background:{c("accent")}; margin-bottom:8px"></div><div style="color:{c("accent")}">经营分析报告</div>'
                f'{title}</div><div style="position:absolute; bottom:7%; left:10%; right:10%; border-top:1px solid '
                f'{c("rule_strong")}; color:{c("muted")}; padding-top:4px">日期 · 编制 · 版本</div>')
    else:
        body = (f'<div style="position:absolute; top:7%; left:10%; right:10%"><div style="color:{c("accent")}">经营分析报告</div>'
                f'{title}<div style="color:{c("muted")}; margin-top:4px">2026 年 10 月 8 日 · 经营分析部</div>'
                f'<div style="border-top:1px solid {c("rule_strong")}; margin-top:6px"></div>'
                f'<div style="margin-top:10px; color:{c("muted")}">正文从第一页开始……</div></div>')
    return f'<div class="page">{body}</div>'


def _slide_mocks(theme: Theme) -> str:
    c, layout = theme.color, theme.get("layout.slide")
    paper, dark = c("slide_paper"), c("dark")
    title = f'{_role_css(theme, "title")} color:{c("ink")}'
    cover_variant = layout["cover"]
    if cover_variant == "split":
        cover = (f'<div class="slide" style="background:{paper}"><div style="position:absolute; left:6%; top:30%; {title}; '
                 f'font-size:26px">经营回顾</div><div style="position:absolute; right:0; top:0; bottom:0; width:38%; '
                 f'background:{c("surface")}; padding:6% 4%; font-size:12px; color:{c("muted")}">营收 <b style="display:block; '
                 f'font-size:22px; color:{c("ink")}">1.28 亿</b>同比 <b style="display:block; font-size:22px; '
                 f'color:{c("accent")}">+24%</b></div></div>')
    elif cover_variant == "issue":
        cover = (f'<div class="slide" style="background:{paper}"><div style="position:absolute; left:6%; top:8%; font-size:10px; '
                 f'color:{c("muted")}">ISSUE · 2026 Q3</div><div style="position:absolute; left:6%; bottom:14%; '
                 f'border-left:2px solid {c("accent")}; padding-left:8px; {title}; font-size:24px; line-height:1.15">'
                 f'慢一点，<br>也走得远</div></div>')
    else:
        cover = (f'<div class="slide" style="background:{paper}"><div style="position:absolute; left:6%; top:34%; '
                 f'color:{c("accent")}; font-size:11px">2026 年第三季度</div><div style="position:absolute; left:6%; top:44%; '
                 f'{title}; font-size:26px">经营回顾</div></div>')
    if layout["section"] == "dark":
        section = (f'<div class="slide" style="background:{dark}"><div style="position:absolute; left:6%; top:12%; '
                   f'color:{c("accent_on_dark")}; {_role_css(theme, "display")} font-size:54px; line-height:1">01</div>'
                   f'<div style="position:absolute; left:6%; bottom:16%; color:{c("on_dark")}; {_role_css(theme, "title")} '
                   f'font-size:22px">一个季度的账本</div></div>')
    else:
        section = (f'<div class="slide" style="background:{paper}"><div style="position:absolute; left:6%; top:24%; '
                   f'color:{c("accent")}; {_role_css(theme, "display")} font-size:48px; line-height:1">01</div>'
                   f'<div style="position:absolute; left:6%; top:58%; {title}; font-size:22px">一个季度的账本</div></div>')
    bullets = "".join(f'<li>{escape(b)}</li>' for b in ("积分商城 12 月上线", "每月会员日", "门店记住你的口味"))
    if layout["content"] == "columns":
        content = (f'<div class="slide" style="background:{paper}"><div style="position:absolute; left:6%; top:12%; width:30%; '
                   f'font-size:10px; color:{c("accent")}">会员<div style="{title}; font-size:16px; margin-top:6px">'
                   f'三个抓手</div></div><ul style="position:absolute; left:42%; top:10%; right:6%; font-size:12px; '
                   f'padding-left:16px">{bullets}</ul></div>')
    else:
        content = (f'<div class="slide" style="background:{paper}"><div style="position:absolute; left:6%; top:10%; {title}; '
                   f'font-size:18px">会员运营的三个抓手</div><ul style="position:absolute; left:6%; top:34%; right:6%; '
                   f'font-size:12px; padding-left:16px">{bullets}</ul></div>')
    return cover + section + content
