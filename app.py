from io import BytesIO
from html import escape
from pathlib import Path

import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="Listing Creative Lab",
    page_icon="✳",
    layout="wide",
    initial_sidebar_state="collapsed",
)

DEMO_WARNING = "DEMO MODE · NOT CONNECTED TO AMAZON · Results are examples only"
EXPECTED = ["Product", "Search Data", "Customer Questions", "Competitors", "Test Results"]
WORKSPACE_PAGES = ["Home", "Shopper signals", "Brief builder", "Image studio", "Weekly archive", "Experiment review"]
HISTORY_SHEETS = ["Search Data", "Customer Questions", "Competitors", "Test Results"]


def clean(value, fallback="Not provided"):
    if pd.isna(value) or str(value).strip() == "":
        return fallback
    return str(value).strip()


def read_workbook(upload):
    book = pd.ExcelFile(upload)
    missing = [sheet for sheet in EXPECTED if sheet not in book.sheet_names]
    if missing:
        raise ValueError("Missing workbook tabs: " + ", ".join(missing))
    return {sheet: pd.read_excel(book, sheet_name=sheet) for sheet in EXPECTED}


def add_week_dates(data, default_week):
    dated = {name: frame.copy() for name, frame in data.items()}
    default_week = pd.Timestamp(default_week).normalize()
    for name in HISTORY_SHEETS:
        frame = dated[name]
        if "week_start" not in frame.columns:
            frame.insert(0, "week_start", default_week)
        else:
            parsed = pd.to_datetime(frame["week_start"], errors="coerce").dt.normalize()
            dated[name]["week_start"] = parsed.fillna(default_week)
    return dated


def merge_weekly_data(existing, incoming, default_week):
    incoming = add_week_dates(incoming, default_week)
    merged = {name: frame.copy() for name, frame in existing.items()}
    if not incoming["Product"].empty:
        merged["Product"] = incoming["Product"].copy()
    for name in HISTORY_SHEETS:
        old = add_week_dates(merged, default_week)[name]
        new = incoming[name]
        replaced_weeks = set(new["week_start"].dropna())
        old = old[~old["week_start"].isin(replaced_weeks)]
        combined = pd.concat([old, new], ignore_index=True, sort=False)
        merged[name] = combined.sort_values("week_start", kind="stable").reset_index(drop=True)
    return merged


def archive_bytes(data):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for name in EXPECTED:
            frame = data[name].copy()
            frame.to_excel(writer, sheet_name=name, index=False)
            ws = writer.book[name]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
    return buffer.getvalue()


def save_archive_file(data, path):
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name in EXPECTED:
            frame = data[name].copy()
            frame.to_excel(writer, sheet_name=name, index=False)
            ws = writer.book[name]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions


def get_product(data):
    if data["Product"].empty:
        return {}
    return data["Product"].iloc[0].to_dict()


def make_hypotheses(data, limit=5):
    """Rank and condense workbook signals into a small, varied test shortlist."""
    search = data["Search Data"].copy()
    questions = data["Customer Questions"].copy()
    competitors = data["Competitors"].copy()
    by_category = {"search": [], "question": [], "gap": []}

    # Combine repeated keywords before ranking. High exposure with relatively
    # low CTR is treated as a useful image opportunity to investigate.
    if not search.empty and "keyword" in search.columns:
        for col in ["impressions", "clicks", "orders"]:
            if col not in search.columns:
                search[col] = 0
            search[col] = pd.to_numeric(search[col], errors="coerce").fillna(0)
        search["_key"] = search["keyword"].map(lambda v: clean(v).casefold())
        search = search[search["_key"] != "not provided"]
        grouped = search.groupby("_key", sort=False).agg(
            keyword=("keyword", "first"), impressions=("impressions", "sum"),
            clicks=("clicks", "sum"), orders=("orders", "sum"), rows=("keyword", "size"),
        ).reset_index(drop=True)
        grouped["ctr"] = grouped.apply(lambda r: r["clicks"] / r["impressions"] if r["impressions"] else 0, axis=1)
        grouped["score"] = grouped["impressions"] * (1 - grouped["ctr"])
        for _, row in grouped.sort_values("score", ascending=False).iterrows():
            term = clean(row["keyword"])
            row_note = f" It occurs in {int(row['rows'])} workbook rows." if row["rows"] > 1 else ""
            by_category["search"].append({
                "title": f"Improve the search image for “{term}”",
                "area": "Main image · click-through",
                "why": f"“{term}” has {int(row['impressions']):,} combined impressions and a {row['ctr']:.1%} click rate.{row_note} High exposure with room to improve clicks makes it a candidate to investigate.",
                "test": "Compare the current image with one clearer product-focused alternative. Change one main visual idea.",
                "metric": "Click-through rate (clicks ÷ impressions)",
                "concept": f"Make the product easy to recognize at thumbnail size for shoppers searching “{term}”.",
                "score": float(row["score"]),
                "category": "search",
            })

    # Group similar question rows by topic so repetition, not row count alone,
    # determines priority. Keep example comments as traceable evidence.
    if not questions.empty and "comment" in questions.columns:
        if "topic" not in questions.columns:
            questions["topic"] = questions["comment"]
        if "frequency" not in questions.columns:
            questions["frequency"] = 1
        questions["frequency"] = pd.to_numeric(questions["frequency"], errors="coerce").fillna(1)
        questions["_key"] = questions["topic"].map(lambda v: clean(v, "product use").casefold())
        questions = questions[questions["_key"] != "not provided"]
        grouped = questions.groupby("_key", sort=False).agg(
            topic=("topic", "first"), frequency=("frequency", "sum"), rows=("comment", "size"),
            comments=("comment", lambda s: list(dict.fromkeys(clean(v) for v in s if clean(v) != "Not provided"))[:2]),
        ).reset_index(drop=True)
        for _, row in grouped.sort_values(["frequency", "rows"], ascending=False).iterrows():
            topic = clean(row["topic"], "product use")
            comments = " · ".join(f"“{comment}”" for comment in row["comments"])
            frequency = int(row["frequency"])
            by_category["question"].append({
                "title": f"Clarify the recurring question about {topic}",
                "area": "Secondary image · product understanding",
                "why": f"This topic appears in {int(row['rows'])} question row(s), with {frequency} combined mentions. Examples: {comments}.",
                "test": "Compare an explanatory secondary image with the current-style image.",
                "metric": "Conversion rate or sales per visitor",
                "concept": f"Use one simple visual to answer the recurring question about {topic}, supported by approved product facts.",
                "score": float(frequency) + int(row["rows"]) * 0.01,
                "category": "question",
            })

    # Collapse identical competitor gaps and prioritize the gaps seen most often.
    if not competitors.empty and "image_gap" in competitors.columns:
        if "competitor_name" not in competitors.columns:
            competitors["competitor_name"] = "Competitor listing"
        competitors["_key"] = competitors["image_gap"].map(lambda v: clean(v).casefold())
        competitors = competitors[competitors["_key"] != "not provided"]
        grouped = competitors.groupby("_key", sort=False).agg(
            gap=("image_gap", "first"), rows=("image_gap", "size"),
            brands=("competitor_name", lambda s: list(dict.fromkeys(clean(v) for v in s))[:4]),
        ).reset_index(drop=True)
        for _, row in grouped.sort_values("rows", ascending=False).iterrows():
            gap = clean(row["gap"])
            brands = ", ".join(row["brands"])
            by_category["gap"].append({
                "title": f"Explore the recurring competitor gap: {gap}",
                "area": "Gallery or A+ content · competitive opportunity",
                "why": f"This gap appears across {int(row['rows'])} competitor observation(s): {brands}. Treat it as an opportunity to explore, then validate it with shoppers.",
                "test": "Create one gallery or A+ concept that addresses this gap and compare it with the current layout.",
                "metric": "Conversion rate or sales per visitor",
                "concept": f"Make {gap.lower()} clearer in one focused gallery image or A+ module, using approved product facts.",
                "score": float(row["rows"]),
                "category": "gap",
            })

    # Ensure the shortlist reflects different evidence types before adding
    # second-tier items from whichever category has the strongest next signals.
    product = get_product(data)
    product_benefits = clean(product.get("main_benefits"), "the product's approved benefits")
    usage = clean(product.get("how_to_use"), "how to use the product")
    aplus = {
        "title": "Build an A+ story around product benefits and use",
        "area": "A+ content · product understanding",
        "why": "The uploaded product facts provide a basis for an A+ layout that explains the product and its use.",
        "test": "Compare a benefit-led layout with a routine and how-to layout.",
        "metric": "Conversion rate or sales per visitor",
        "concept": f"Organize A+ content around these approved benefits: {product_benefits}. Explain how to use it: {usage}.",
        "score": 0.0,
        "category": "aplus",
    }
    selected = []
    # One leading item per available evidence source, plus the A+ idea.
    for category in ("search", "question", "gap"):
        if by_category[category]:
            selected.append(by_category[category][0])
    selected.append(aplus)
    # Fill extra slots in ranked rounds so large search counts do not crowd
    # out recurring question and competitor signals just because their units
    # (mentions) are smaller than impressions.
    rank = 1
    while len(selected) < limit:
        added = False
        for category in ("search", "question", "gap"):
            if len(by_category[category]) > rank:
                selected.append(by_category[category][rank])
                added = True
                if len(selected) >= limit:
                    break
        if not added:
            break
        rank += 1
    return selected[:limit]


def make_brief(idea, product, format_name):
    name = clean(product.get("product_name"), "Your product")
    facts = clean(product.get("main_benefits"))
    claims = clean(product.get("claims_allowed"))
    avoid = clean(product.get("claims_to_avoid"))
    colors = clean(product.get("brand_colors"))
    audience = clean(product.get("target_customer"))
    use = clean(product.get("how_to_use"))
    if format_name == "Main image":
        purpose = "Help a shopper recognize the product clearly in Amazon search results. Use a clean white background and show the product alone. Do not add badges, benefit text, decorative props, or invented packaging details. Check current category image rules before use."
        placement = "Search result thumbnail"
    elif format_name == "Secondary image":
        purpose = f"Explain one useful product detail at a glance. Concept: {idea['concept']} Use short, legible copy only where supported by the product facts."
        placement = "Product image gallery"
    else:
        purpose = f"Create one A+ module concept based on this idea: {idea['concept']} Use a clear hierarchy, simple visuals, and concise copy."
        placement = "A+ content module"
    return f"""CREATIVE BRIEF · {format_name}

Product: {name}
Placement: {placement}
Audience: {audience}
Hypothesis: {idea['title']}
Evidence: {idea['why']}
Success measure: {idea['metric']}

GOAL
{purpose}

PRODUCT FACTS TO PRESERVE
- Benefits/features: {facts}
- How to use: {use}
- Approved claims: {claims}
- Do not claim: {avoid}
- Brand colors: {colors}

IMAGE-GENERATION PROMPT
Create a polished ecommerce creative draft for “{name}”. Treat any supplied product photo as the source of truth: preserve its shape, label, colors, proportions, and readable brand name. Audience: {audience}. Visual direction: {purpose} Brand palette: {colors}. Use only these factual benefits: {facts}. Do not invent ingredients, certifications, clinical results, awards, review quotes, or before-and-after outcomes. Avoid these claims: {avoid}. Create one clean concept for review, not a claim of Amazon compliance. Keep copy minimal and spelling accurate.

REVIEW BEFORE USE
- Does the product match the supplied reference?
- Are all claims supported by the product facts?
- Is text readable on a phone?
- Does this concept test one main change?
- Does it meet the selected placement's current rules?
"""


def calculate_metrics(results):
    df = results.copy()
    for col in ["impressions", "clicks", "orders"]:
        if col in df:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    df["CTR"] = df.apply(lambda r: r.get("clicks", 0) / r.get("impressions", 0) if r.get("impressions", 0) else 0, axis=1)
    df["Order rate"] = df.apply(lambda r: r.get("orders", 0) / r.get("clicks", 0) if r.get("clicks", 0) else 0, axis=1)
    return df


CHART_COLORS = ["#8E73DF", "#6FA8E4", "#F07C72", "#58AD7B", "#E5A956", "#87909B", "#C889B6"]


def show_minimal_line_chart(wide, value_title, height=230, zero=True):
    chart_data = wide.copy().apply(pd.to_numeric, errors="coerce").fillna(0)
    if chart_data.empty or len(chart_data.index) == 0:
        st.caption("No weekly values to chart yet.")
        return
    series_names = [str(name) for name in chart_data.columns]
    raw_values = chart_data.to_numpy().flatten().tolist()
    low, high = min(raw_values), max(raw_values)
    if zero:
        low = 0
        high = high * 1.08 if high else 1
    else:
        pad = (high - low) * .12 if high != low else max(abs(high) * .06, 1)
        low, high = low - pad, high + pad
    if high <= low:
        high = low + 1

    width = 900
    left, right, top, bottom = 62, 22, 18, 56
    plot_width, plot_height = width - left - right, height - top - bottom
    x_count = len(chart_data.index)
    x_positions = [left + (plot_width * i / max(x_count - 1, 1)) for i in range(x_count)]
    integer_ticks = "mention" in value_title.lower()
    label_fmt = (lambda value: f"{value:,.0f}") if integer_ticks else (lambda value: f"{value:.1f}")

    parts = [f"<div class='chart-shell'><svg viewBox='0 0 {width} {height}' role='img' aria-label='{escape(value_title)} trend' preserveAspectRatio='xMidYMid meet'>"]
    for tick_index in range(5):
        tick = low + (high - low) * tick_index / 4
        y = top + plot_height * (1 - tick_index / 4)
        parts.append(f"<line x1='{left}' y1='{y:.1f}' x2='{width-right}' y2='{y:.1f}' stroke='#eceaf0' stroke-width='1'/>")
        parts.append(f"<text x='{left-10}' y='{y+4:.1f}' text-anchor='end' fill='#85858d' font-size='11' font-family='Arial,sans-serif'>{label_fmt(tick)}</text>")

    for series_index, series in enumerate(chart_data.columns):
        color = CHART_COLORS[series_index % len(CHART_COLORS)]
        points = []
        for x, value in zip(x_positions, chart_data[series].tolist()):
            y = top + plot_height * (1 - (float(value) - low) / (high - low))
            points.append((x, y, float(value)))
        point_string = " ".join(f"{x:.1f},{y:.1f}" for x, y, _ in points)
        parts.append(f"<polyline points='{point_string}' fill='none' stroke='{color}' stroke-width='3' stroke-linecap='round' stroke-linejoin='round'/>")
        for x, y, value in points:
            parts.append(f"<circle cx='{x:.1f}' cy='{y:.1f}' r='4.2' fill='{color}' stroke='#fff' stroke-width='2'><title>{escape(str(series))}: {value:.2f}</title></circle>")

    for index, x in enumerate(x_positions):
        try:
            date_label = pd.Timestamp(chart_data.index[index]).strftime("%d %b")
        except Exception:
            date_label = str(chart_data.index[index])
        parts.append(f"<text x='{x:.1f}' y='{height-34}' text-anchor='middle' fill='#777982' font-size='11' font-family='Arial,sans-serif'>{escape(date_label)}</text>")

    if len(series_names) > 1:
        legend_y, legend_x = height - 11, left
        for index, series in enumerate(series_names):
            color = CHART_COLORS[index % len(CHART_COLORS)]
            parts.append(f"<circle cx='{legend_x}' cy='{legend_y-4}' r='4' fill='{color}'/>")
            parts.append(f"<text x='{legend_x+10}' y='{legend_y}' fill='#686a70' font-size='11' font-family='Arial,sans-serif'>{escape(series[:30])}</text>")
            legend_x += min(185, 22 + len(series) * 7)
    parts.append("</svg></div>")
    st.markdown("".join(parts), unsafe_allow_html=True)


def show_minimal_bar_chart(frame, category, value, value_title, height=235, color_field=None):
    plot = frame.copy()
    if plot.empty:
        st.caption("No values to compare yet.")
        return
    plot[value] = pd.to_numeric(plot[value], errors="coerce").fillna(0)
    if color_field and color_field in plot.columns:
        series_names = list(dict.fromkeys(plot[color_field].astype(str).tolist()))
    else:
        series_names = [""]
    group_values = plot.groupby(category, sort=False)[value].sum().sort_values(ascending=False)
    categories = group_values.index.tolist()
    width = 900
    left, right, top, bottom = 300, 38, 24, 35
    row_height = 27 if len(series_names) > 1 else 25
    dynamic_height = max(height, top + bottom + len(categories) * row_height + (22 if len(series_names) > 1 else 0))
    plot_width = width - left - right
    maximum = float(plot[value].max()) or 1
    maximum *= 1.08
    parts = [f"<div class='chart-shell'><svg viewBox='0 0 {width} {dynamic_height}' role='img' aria-label='{escape(value_title)} comparison' preserveAspectRatio='xMidYMid meet'>"]
    for tick_index in range(5):
        tick = maximum * tick_index / 4
        x = left + plot_width * tick_index / 4
        parts.append(f"<line x1='{x:.1f}' y1='{top}' x2='{x:.1f}' y2='{dynamic_height-bottom}' stroke='#efedf2' stroke-width='1'/>")
        parts.append(f"<text x='{x:.1f}' y='{dynamic_height-10}' text-anchor='middle' fill='#85858d' font-size='10' font-family='Arial,sans-serif'>{tick:,.0f}</text>")
    for category_index, category_value in enumerate(categories):
        row_top = top + category_index * row_height
        label = str(category_value)
        short_label = label if len(label) <= 42 else label[:39] + "…"
        parts.append(f"<text x='{left-12}' y='{row_top+15}' text-anchor='end' fill='#55575f' font-size='11' font-family='Arial,sans-serif'><title>{escape(label)}</title>{escape(short_label)}</text>")
        subset = plot[plot[category].eq(category_value)]
        for series_index, series in enumerate(series_names):
            if series:
                values = subset[subset[color_field].astype(str).eq(series)][value]
            else:
                values = subset[value]
            amount = float(values.sum()) if not values.empty else 0
            bar_height = 8 if len(series_names) > 1 else 14
            offset = series_index * (bar_height + 3)
            y = row_top + (1 if len(series_names) > 1 else 2) + offset
            bar_width = max(0, plot_width * amount / maximum)
            color = CHART_COLORS[series_index % len(CHART_COLORS)] if series else "#9A83E8"
            series_label = f" · {escape(series)}" if series else ""
            parts.append(f"<rect x='{left}' y='{y}' width='{bar_width:.1f}' height='{bar_height}' rx='4' fill='{color}'><title>{escape(label)}{series_label}: {amount:,.0f}</title></rect>")
    if len(series_names) > 1:
        legend_y, legend_x = dynamic_height - 10, left
        for index, series in enumerate(series_names):
            color = CHART_COLORS[index % len(CHART_COLORS)]
            parts.append(f"<circle cx='{legend_x}' cy='{legend_y-4}' r='4' fill='{color}'/>")
            parts.append(f"<text x='{legend_x+10}' y='{legend_y}' fill='#686a70' font-size='10' font-family='Arial,sans-serif'>{escape(series[:26])}</text>")
            legend_x += min(180, 22 + len(series) * 6)
    parts.append("</svg></div>")
    st.markdown("".join(parts), unsafe_allow_html=True)


def show_clean_table(frame):
    html_table = frame.to_html(index=False, escape=True, classes="clean-table")
    st.markdown(f"<div class='clean-table-wrap'>{html_table}</div>", unsafe_allow_html=True)


st.markdown("""
<style>
:root {
  --ink: #191a1d;
  --muted: #686a70;
  --line: #e7e6e5;
  --coral: #ff796e;
  --lavender: #cbbcff;
  --blue: #b9d8f4;
  --green: #b9e8c7;
  --peach: #ffd0ad;
  --paper: #f3f3f1;
  --surface: #ffffff;
}
.stApp {
  color: var(--ink);
  background: radial-gradient(ellipse at 88% 0%, rgba(203,188,255,.22), transparent 32%),
              radial-gradient(ellipse at 8% 20%, rgba(185,216,244,.18), transparent 32%),
              var(--paper);
}
.block-container { max-width: 1280px; padding-top: 1.2rem; padding-bottom: 3rem; }
.stApp, .stApp p, .stApp label, .stApp li, .stApp [data-testid="stMarkdownContainer"] { color: var(--ink); }
.stApp [data-testid="stCaptionContainer"], .stApp small { color: var(--muted) !important; }
[data-testid="stSidebar"] { background: #f7f7f5; border-right: 1px solid var(--line); }
[data-testid="stSidebar"] * { color: var(--ink); }
[data-testid="stSidebar"] [data-testid="stCaptionContainer"] { color: var(--muted) !important; }
.brand-row { display:flex; align-items:center; gap:.65rem; margin: .3rem 0 1.1rem; color:var(--ink); font-weight:750; letter-spacing:-.02em; }
.brand-mark { width:32px; height:32px; border-radius:11px; display:inline-flex; align-items:center; justify-content:center; background:linear-gradient(145deg,#bca8ff,#8aa8f8); color:white; font-size:18px; box-shadow:0 5px 16px rgba(120,104,205,.18); }
.eyebrow { color:#76767c; text-transform:uppercase; letter-spacing:.13em; font-size:.72rem; font-weight:750; }
.hero-title { font-size:clamp(2.2rem,4.8vw,4.1rem); font-weight:760; letter-spacing:-.055em; line-height:1.02; color:var(--ink); margin:.35rem 0 .85rem; max-width:850px; }
.hero-copy { font-size:1.08rem; line-height:1.62; color:#62646b; max-width:780px; }
.chart-shell { width:100%; overflow:hidden; padding:12px 14px 8px; border:1px solid #ebe9ef; border-radius:16px; background:#fff; box-shadow:0 6px 20px rgba(30,31,38,.035); }
.chart-shell svg { display:block; width:100%; height:auto; }
.demo-banner { display:inline-flex; gap:.5rem; align-items:center; margin-top:1.1rem; padding:.65rem 1rem; border:1px solid #eadcb7; border-radius:999px; background:#fff4d8; color:#493b1f; font-size:.82rem; font-weight:700; }
.flow-wrap { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:12px; margin:1.35rem 0 1.8rem; }
.flow-card { min-height:150px; padding:19px 20px 17px; border:1px solid rgba(25,26,29,.06); border-radius:21px; background:#fff; box-shadow:0 8px 26px rgba(30,31,38,.045); position:relative; }
.flow-card:nth-child(1) { background:linear-gradient(145deg,#fff 40%,#eee8ff); }
.flow-card:nth-child(2) { background:linear-gradient(145deg,#fff 40%,#e4f0fc); }
.flow-card:nth-child(3) { background:linear-gradient(145deg,#fff 40%,#e7f6eb); }
.flow-card:nth-child(4) { background:linear-gradient(145deg,#fff 40%,#fff0e3); }
.flow-card:not(:last-child):after { content:'→'; position:absolute; right:-15px; top:25px; z-index:2; width:28px; height:28px; border-radius:50%; display:flex; align-items:center; justify-content:center; background:#191a1d; color:#fff; font-weight:700; }
.flow-index { color:#777980; font-size:.72rem; font-weight:750; letter-spacing:.09em; text-transform:uppercase; }
.flow-title { color:#191a1d; font-size:1.05rem; font-weight:750; margin:.55rem 0 .35rem; }
.flow-copy { color:#686a70; font-size:.9rem; line-height:1.5; }
.panel { padding:21px 23px; border:1px solid rgba(25,26,29,.07); border-radius:22px; background:#fff; box-shadow:0 8px 26px rgba(30,31,38,.045); }
.muted { color:var(--muted); }
.metric-grid { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:14px; margin:1.5rem 0 1.1rem; }
.metric-card { min-height:140px; padding:19px 21px; border-radius:22px; border:1px solid rgba(25,26,29,.05); box-shadow:0 8px 26px rgba(30,31,38,.05); }
.metric-card:nth-child(1) { background:linear-gradient(145deg,#bda9ff,#d9ceff); }
.metric-card:nth-child(2) { background:linear-gradient(145deg,#b4d9f7,#d9ecfb); }
.metric-card:nth-child(3) { background:linear-gradient(145deg,#aee3bc,#d2f0d9); }
.metric-card:nth-child(4) { background:linear-gradient(145deg,#ffc69c,#ffe2c7); }
.metric-label { color:rgba(25,26,29,.68); font-size:.75rem; font-weight:700; letter-spacing:.07em; text-transform:uppercase; }
.metric-value { color:#151619; font-size:1.75rem; font-weight:750; letter-spacing:-.04em; margin-top:.65rem; line-height:1.1; }
.metric-note { color:rgba(25,26,29,.68); font-size:.8rem; margin-top:.45rem; }
.notice-card { padding:15px 18px; border:1px solid #ded6f4; border-radius:16px; background:linear-gradient(135deg,#f3efff,#edf6fc); color:#26232f; line-height:1.55; }
.variant-card { min-height:125px; padding:17px 18px; border:1px solid rgba(25,26,29,.06); border-radius:20px; background:#fff; box-shadow:0 7px 22px rgba(30,31,38,.04); }
.variant-name { color:#6b6d74; font-size:.72rem; font-weight:700; letter-spacing:.07em; text-transform:uppercase; }
.variant-rate { color:#191a1d; font-size:1.65rem; font-weight:750; margin-top:.55rem; }
.variant-note { color:#686a70; font-size:.82rem; margin-top:.35rem; }
.clean-table-wrap { max-height:420px; overflow:auto; border:1px solid #e7e5eb; border-radius:16px; background:#fff; box-shadow:0 6px 20px rgba(30,31,38,.035); }
table.clean-table { width:100%; border-collapse:separate; border-spacing:0; background:#fff !important; color:#25262a !important; font-size:.86rem; }
table.clean-table th { position:sticky; top:0; z-index:1; background:#f0ecfb !important; color:#2a2732 !important; font-weight:700; text-align:left; padding:11px 13px; border-bottom:1px solid #e5e0f0; }
table.clean-table td { background:#fff !important; color:#303137 !important; padding:10px 13px; border-bottom:1px solid #f0eff2; vertical-align:top; }
table.clean-table tr:last-child td { border-bottom:0; }
table.clean-table tbody tr:hover td { background:#f8f6ff !important; }
.stApp h1,.stApp h2,.stApp h3,.stApp h4 { color:var(--ink); }
.stApp [data-testid="stRadio"] label, .stApp [data-testid="stCheckbox"] label { color:var(--ink) !important; }
.stApp [data-testid="stExpander"] { background:#fff !important; border:1px solid var(--line); border-radius:16px; overflow:hidden; }
.stApp [data-testid="stExpander"] details,
.stApp [data-testid="stExpander"] summary,
.stApp [data-testid="stExpander"] [data-testid="stExpanderDetails"],
.stApp [data-testid="stExpander"] [data-testid="stExpanderToggleIcon"] { background:#fff !important; color:#191a1d !important; }
.stApp [data-testid="stExpander"] * { color:#191a1d !important; }
.stApp [data-testid="stExpander"] summary:hover { background:#f2efff !important; }
.stApp [data-testid="stAlert"] { background:#f2efff !important; border:1px solid #ded6f4 !important; border-radius:16px !important; color:#24252a !important; }
.stApp [data-testid="stAlert"] * { color:#24252a !important; }
.stButton button, .stDownloadButton button,
[data-testid="stFileUploader"] button { background:linear-gradient(135deg,#d9ceff,#c8e1f6) !important; color:#22212a !important; border:1px solid #cfc5ee !important; border-radius:14px !important; font-weight:700; box-shadow:0 5px 14px rgba(95,82,150,.10); }
.stButton button *, .stDownloadButton button *, [data-testid="stFileUploader"] button * { color:#22212a !important; fill:#22212a !important; }
.stButton button:hover, .stDownloadButton button:hover,
[data-testid="stFileUploader"] button:hover { background:linear-gradient(135deg,#cbbcff,#b5d7f3) !important; color:#191a1d !important; border-color:#b4a4e8 !important; }
.stButton [data-testid="stBaseButton-primary"] { background:linear-gradient(135deg,#ffc6a2,#ffb0ae) !important; color:#29201e !important; border-color:#f4b39e !important; }
.stTextInput input,.stTextArea textarea,.stSelectbox [role="combobox"],.stMultiSelect [role="combobox"],.stFileUploader section { background:#fff !important; color:var(--ink) !important; border-color:#d8d8d6 !important; border-radius:12px; }
.stDateInput input, [data-testid="stDateInput"] input, [data-baseweb="input"] input { background:#fff !important; color:#191a1d !important; border-color:#d8d8d6 !important; }
[data-testid="stDateInput"] [data-baseweb="input"] > div { background:#fff !important; border-color:#d8d8d6 !important; }
[data-baseweb="select"] > div, [data-baseweb="select"] input { background:#fff !important; color:#191a1d !important; }
[data-baseweb="popover"], [data-baseweb="menu"], [role="listbox"], ul[role="listbox"] { background:#fff !important; color:#191a1d !important; border:1px solid #e4e0ed !important; border-radius:12px !important; box-shadow:0 14px 34px rgba(33,28,48,.14) !important; }
[role="option"], li[role="option"] { background:#fff !important; color:#191a1d !important; }
[role="option"]:hover, li[role="option"]:hover, [role="option"][aria-selected="true"] { background:#eee9ff !important; color:#191a1d !important; }
[data-testid="stFileUploader"] section { border:1px dashed #b9abd9 !important; border-radius:16px !important; }
[data-testid="stFileUploader"] small, [data-testid="stFileUploader"] [data-testid="stFileUploaderDropzoneInstructions"] { color:#686a70 !important; }
.stApp [data-testid="stSlider"] [role="slider"] { background:#9a83e8 !important; border-color:#fff !important; }
.stApp [data-testid="stMetricDelta"] { background:#e4f3e9; color:#34734c !important; border-radius:999px; padding:.2rem .55rem; }
.stApp [data-testid="stMetricDelta"] * { color:#34734c !important; }
.stApp [data-testid="stDataFrame"], .stApp [data-testid="stTable"] { background:#fff; }
.stApp [data-testid="stAlert"] { color:#24252a; }
.stApp [data-testid="stMetricLabel"], .stApp [data-testid="stMetricValue"] { color:#191a1d !important; }
div[role="radiogroup"] label { color:var(--ink); }
@media (max-width: 900px) { .metric-grid { grid-template-columns:repeat(2,minmax(0,1fr)); } }
@media (max-width: 850px) { .flow-wrap { grid-template-columns:repeat(2,minmax(0,1fr)); } .flow-card:nth-child(2):after { display:none; } }
@media (max-width: 560px) { .flow-wrap,.metric-grid { grid-template-columns:1fr; } .flow-card:after { display:none!important; } }
</style>
""", unsafe_allow_html=True)


with st.sidebar:
    st.markdown('<div class="brand-row"><span class="brand-mark">✳</span> Listing Creative Lab</div>', unsafe_allow_html=True)
    st.caption("WORKSPACE")
    page = st.radio("Navigate", WORKSPACE_PAGES, label_visibility="collapsed")
    st.divider()
    st.caption("WORKBOOK")
    sample_path = Path(__file__).parent / "demo_inputs.xlsx"
    if sample_path.exists():
        st.download_button("Download example workbook", sample_path.read_bytes(), "demo_inputs.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    uploaded = st.file_uploader("Upload your workbook", type=["xlsx"], help="Keep the five expected tab names and column headings.")
    st.caption("The workbook updates the evidence and recommendations. Image creation happens in ChatGPT; upload the finished options in Image studio.")

try:
    if uploaded:
        data = read_workbook(uploaded)
    elif sample_path.exists():
        data = read_workbook(sample_path)
    else:
        st.info("Open the menu and upload your workbook to get started.")
        st.stop()
except Exception as exc:
    st.error(f"I couldn't read that workbook: {exc}")
    st.stop()

archive_path = Path(__file__).parent / "weekly_archive.xlsx"
if "archive_data" not in st.session_state:
    try:
        archive_source = archive_path if archive_path.exists() else sample_path
        st.session_state.archive_data = add_week_dates(read_workbook(archive_source), pd.Timestamp.today())
    except Exception:
        st.session_state.archive_data = add_week_dates(data, pd.Timestamp.today())

product = get_product(data)
hypotheses = make_hypotheses(data)
product_name = clean(product.get("product_name"), "Your product")


def section_heading(kicker, title, description):
    st.markdown(f'<div class="eyebrow">{kicker}</div><h1 style="letter-spacing:-.04em;color:#191a1d;margin:.25rem 0 .4rem">{title}</h1><div class="muted">{description}</div>', unsafe_allow_html=True)


if page == "Home":
    st.markdown('<div class="eyebrow">AMAZON LISTING CREATIVE WORKSPACE</div>', unsafe_allow_html=True)
    st.markdown('<div class="hero-title">From shopper signals<br>to stronger listing creative.</div>', unsafe_allow_html=True)
    st.markdown('<div class="hero-copy">Turn weekly marketplace signals into focused creative briefs, review image directions, and track what changes over time.</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="demo-banner">{DEMO_WARNING}</div>', unsafe_allow_html=True)

    archive_search = st.session_state.archive_data["Search Data"].copy()
    if not archive_search.empty and {"week_start", "impressions", "clicks"}.issubset(archive_search.columns):
        archive_search["week_start"] = pd.to_datetime(archive_search["week_start"], errors="coerce")
        archive_search["impressions"] = pd.to_numeric(archive_search["impressions"], errors="coerce").fillna(0)
        archive_search["clicks"] = pd.to_numeric(archive_search["clicks"], errors="coerce").fillna(0)
        latest_week = archive_search["week_start"].max()
        latest_rows = archive_search[archive_search["week_start"].eq(latest_week)]
        latest_ctr = latest_rows["clicks"].sum() / latest_rows["impressions"].sum() if latest_rows["impressions"].sum() else 0
        week_count = archive_search["week_start"].nunique()
    else:
        latest_ctr, week_count = 0, 0
    search_term_count = data["Search Data"]["keyword"].nunique() if "keyword" in data["Search Data"] else len(data["Search Data"])
    question_topics = data["Customer Questions"]["topic"].nunique() if "topic" in data["Customer Questions"] else len(data["Customer Questions"])
    shortlist_count = len(make_hypotheses(data, limit=5))
    st.markdown(f"""
    <div class="metric-grid">
      <div class="metric-card"><div class="metric-label">Latest weekly search CTR</div><div class="metric-value">{latest_ctr:.2%}</div><div class="metric-note">From the archived search totals</div></div>
      <div class="metric-card"><div class="metric-label">Shopper signal coverage</div><div class="metric-value">{search_term_count} + {question_topics}</div><div class="metric-note">Unique search terms and question topics in this upload</div></div>
      <div class="metric-card"><div class="metric-label">Prioritized ideas</div><div class="metric-value">{shortlist_count}</div><div class="metric-note">Shortlist ready for a creative brief</div></div>
      <div class="metric-card"><div class="metric-label">Weeks tracked</div><div class="metric-value">{week_count}</div><div class="metric-note">Choose a week comparison in Weekly archive</div></div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("### The workflow")
    st.markdown("""
    <div class="flow-wrap">
      <div class="flow-card"><div class="flow-index">01 · INPUT</div><div class="flow-title">Shopper signals</div><div class="flow-copy">Search phrases, customer questions, product facts, and competitor observations.</div></div>
      <div class="flow-card"><div class="flow-index">02 · ANALYZE</div><div class="flow-title">Creative hypothesis</div><div class="flow-copy">Spot a possible gap and explain which evidence points to it.</div></div>
      <div class="flow-card"><div class="flow-index">03 · CREATE</div><div class="flow-title">Brief and image prompt</div><div class="flow-copy">Turn the idea into instructions for a main image, gallery image, or A+ module.</div></div>
      <div class="flow-card"><div class="flow-index">04 · COMPARE</div><div class="flow-title">Review and learn</div><div class="flow-copy">Compare drafts, then review test metrics or prepare a real experiment.</div></div>
    </div>
    """, unsafe_allow_html=True)

    left, right = st.columns([1.25, .75], gap="large")
    with left:
        st.markdown("<div class='panel'>", unsafe_allow_html=True)
        st.markdown("#### What happens in this prototype")
        st.markdown("The workbook is read using preset rules. The app builds a short evidence-based hypothesis list, turns a selected idea into a brief, and tracks dated weekly signals and results in the archive.")
        st.markdown("</div>", unsafe_allow_html=True)
    with right:
        st.markdown("<div class='panel'>", unsafe_allow_html=True)
        st.markdown("#### Current product")
        st.markdown(f"**{product_name}**")
        st.caption(f"{clean(product.get('product_type'), 'Product type not set')} · {clean(product.get('target_customer'), 'Audience not set')}")
        search_term_count = data["Search Data"]["keyword"].nunique() if "keyword" in data["Search Data"] else len(data["Search Data"])
        topic_count = data["Customer Questions"]["topic"].nunique() if "topic" in data["Customer Questions"] else len(data["Customer Questions"])
        st.markdown(f"**{search_term_count}** unique search terms&nbsp;&nbsp; · &nbsp;&nbsp;**{topic_count}** question topics", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
    st.caption("Open the menu in the upper-left corner to explore signals, build briefs, review images, or track weekly changes.")

elif page == "Shopper signals":
    section_heading("01 · INPUT", "Shopper signals", "Review the workbook evidence that feeds the creative recommendations.")
    st.write(f"Product: **{product_name}**")
    for sheet in EXPECTED[1:]:
        with st.expander(sheet, expanded=(sheet == "Search Data")):
            frame = data[sheet].copy()
            if "week_start" in frame.columns:
                frame["week_start"] = pd.to_datetime(frame["week_start"], errors="coerce")
                latest = frame["week_start"].max()
                if pd.notna(latest):
                    frame = frame[frame["week_start"].eq(latest)]
                    st.caption(f"Showing rows for the week of {latest:%d %b %Y}.")
                    frame["week_start"] = frame["week_start"].dt.strftime("%d %b %Y")
            if sheet == "Search Data" and "impressions" in frame.columns:
                frame = frame.sort_values("impressions", ascending=False)
            if sheet == "Customer Questions" and "frequency" in frame.columns:
                frame = frame.sort_values("frequency", ascending=False)
            show_clean_table(frame.head(12))

elif page == "Brief builder":
    section_heading("02–03 · ANALYZE & CREATE", "Build a creative brief", "Choose a hypothesis and turn it into a direction for image generation.")
    idea_limit = st.slider("How many prioritized ideas to review?", min_value=3, max_value=7, value=5)
    hypotheses = make_hypotheses(data, limit=idea_limit)
    st.caption(f"Showing {len(hypotheses)} prioritized ideas. The app combines repeated topics, ranks signals by search exposure and shopper/competitor recurrence, and keeps the shortlist to {idea_limit} or fewer.")
    labels = [f"{i+1}. {idea['title']}" for i, idea in enumerate(hypotheses)]
    selected = st.selectbox("Choose an idea", range(len(hypotheses)), format_func=lambda i: labels[i])
    idea = hypotheses[selected]
    st.markdown(f"**{idea['area']}**")
    st.markdown(f"<div class='notice-card'><b>Why this idea?</b><br>{escape(idea['why'])}</div>", unsafe_allow_html=True)
    col_a, col_b = st.columns(2)
    col_a.markdown(f"**Test idea**  \n{idea['test']}")
    col_b.markdown(f"**Measure**  \n{idea['metric']}")
    kind = st.radio("Creative format", ["Main image", "Secondary image", "A+ content"], horizontal=True)
    brief = make_brief(idea, product, kind)
    st.text_area("Copy this brief and prompt into ChatGPT", brief, height=420)
    st.download_button("Download brief", brief, "creative_brief.txt", "text/plain")

elif page == "Image studio":
    section_heading("03 · CREATE", "Image studio", "Generate the creative in ChatGPT, then bring the options here for a side-by-side review.")
    files = st.file_uploader("Upload image options", type=["png", "jpg", "jpeg", "webp"], accept_multiple_files=True)
    if files:
        cols = st.columns(min(3, len(files)))
        for i, file in enumerate(files):
            with cols[i % len(cols)]:
                st.image(file, caption=f"Option {chr(65+i)} · {file.name}", use_container_width=True)
                st.checkbox(f"Prefer option {chr(65+i)}", key=f"preferred_{file.name}")
        chosen = [chr(65+i) for i, file in enumerate(files) if st.session_state.get(f"preferred_{file.name}")]
        if chosen:
            st.success("Your preferred concept: " + ", ".join(chosen))
    else:
        st.markdown("<div class='panel'>Upload two or more options to compare them. Your selection records a creative preference; it does not measure shopper response.</div>", unsafe_allow_html=True)

elif page == "Weekly archive":
    section_heading("05 · TRACK OVER TIME", "Weekly archive", "Save weekly workbooks, compare shopper signals, and follow click-through rate across weeks.")
    st.markdown("<div class='notice-card'>The app keeps an archive file while local storage is available. Download a backup after updates; Streamlit Community Cloud may clear generated files after a restart.</div>", unsafe_allow_html=True)

    restore_file = st.file_uploader("Restore a previous archive workbook", type=["xlsx"], key="archive_restore")
    if restore_file and st.button("Restore archive", key="restore_archive_button"):
        try:
            restored = add_week_dates(read_workbook(restore_file), pd.Timestamp.today())
            st.session_state.archive_data = restored
            try:
                save_archive_file(restored, archive_path)
            except Exception:
                pass
            st.success("Archive restored. Your weekly charts are updated.")
        except Exception as exc:
            st.error(f"I couldn't restore that archive: {exc}")

    st.markdown("#### Add this week's workbook")
    week_start = st.date_input("Week starts on (used only when the workbook has no week_start dates)", value=pd.Timestamp.today().normalize().date())
    if uploaded:
        if st.button("Add or replace these weeks in the archive", type="primary", key="archive_week_button"):
            merged = merge_weekly_data(st.session_state.archive_data, data, week_start)
            st.session_state.archive_data = merged
            try:
                save_archive_file(merged, archive_path)
                st.success("Weekly data added. Rows with matching week dates were replaced to prevent duplicates.")
            except Exception:
                st.warning("The archive is updated for this session. Download the archive below to keep it, then restore that file next time.")
    else:
        st.caption("Upload this week's workbook from the hamburger menu, then add it here. Re-uploading a week replaces that week's rows instead of duplicating them.")

    archive = st.session_state.archive_data
    backup = archive_bytes(archive)
    st.download_button("Download weekly archive backup", backup, "weekly_archive.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_archive")
    st.caption("This prototype's archive is shared by people using the same app. Use demo data unless the app is private and access is controlled.")

    searches = archive["Search Data"].copy()
    questions = archive["Customer Questions"].copy()
    competitors = archive["Competitors"].copy()
    results = archive["Test Results"].copy()
    week_sets = []
    for frame in (searches, questions, competitors, results):
        if "week_start" in frame.columns:
            frame["week_start"] = pd.to_datetime(frame["week_start"], errors="coerce").dt.normalize()
            week_sets.extend(frame["week_start"].dropna().unique().tolist())
    week_options = sorted(pd.Timestamp(value).normalize() for value in set(week_sets))
    if not week_options:
        st.warning("No dated weekly rows are available yet. Add a dated workbook to begin tracking history.")
    else:
        first_col, last_col = st.columns(2)
        with first_col:
            start_week = st.selectbox("Compare from week", week_options, index=0, format_func=lambda d: d.strftime("Week of %d %b %Y"), key="archive_start_week")
        with last_col:
            end_week = st.selectbox("Compare with week", week_options, index=len(week_options)-1, format_func=lambda d: d.strftime("Week of %d %b %Y"), key="archive_end_week")

        if not searches.empty and {"impressions", "clicks", "orders", "week_start"}.issubset(searches.columns):
            for col in ["impressions", "clicks", "orders"]:
                searches[col] = pd.to_numeric(searches[col], errors="coerce").fillna(0)
            search_weekly = searches.groupby("week_start").agg(
                impressions=("impressions", "sum"), clicks=("clicks", "sum"), orders=("orders", "sum")
            ).sort_index()
            search_weekly["CTR"] = search_weekly.apply(lambda r: r["clicks"] / r["impressions"] if r["impressions"] else 0, axis=1)
            search_weekly["CTR (%)"] = search_weekly["CTR"] * 100
            search_weekly["Order rate"] = search_weekly.apply(lambda r: r["orders"] / r["clicks"] if r["clicks"] else 0, axis=1)
            search_weekly["Order rate (%)"] = search_weekly["Order rate"] * 100
            start_data = search_weekly.loc[start_week] if start_week in search_weekly.index else None
            end_data = search_weekly.loc[end_week] if end_week in search_weekly.index else None
            if start_data is not None and end_data is not None:
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Search CTR", f"{end_data['CTR']:.2%}", f"{(end_data['CTR'] - start_data['CTR'])*100:+.2f} percentage points")
                m2.metric("Search order rate", f"{end_data['Order rate']:.2%}", f"{(end_data['Order rate'] - start_data['Order rate'])*100:+.2f} percentage points")
                m3.metric("Impressions", f"{end_data['impressions']:,.0f}", f"{end_data['impressions'] - start_data['impressions']:+,.0f} vs start week")
                m4.metric("Orders", f"{end_data['orders']:,.0f}", f"{end_data['orders'] - start_data['orders']:+,.0f} vs start week")
            left, right = st.columns(2)
            with left:
                st.markdown("#### Search CTR progression")
                show_minimal_line_chart(search_weekly[["CTR (%)"]], "CTR (%)", height=205, zero=False)
            with right:
                st.markdown("#### Search order-rate progression")
                show_minimal_line_chart(search_weekly[["Order rate (%)"]], "Order rate (%)", height=205, zero=False)

        if not questions.empty and "week_start" in questions.columns:
            if "frequency" not in questions.columns:
                questions["frequency"] = 1
            questions["frequency"] = pd.to_numeric(questions["frequency"], errors="coerce").fillna(0)
            if "sentiment" not in questions.columns:
                questions["sentiment"] = "Question"
            sentiment_weekly = questions.groupby(["week_start", "sentiment"])["frequency"].sum().unstack(fill_value=0).sort_index()
            st.markdown("#### Shopper questions and concerns over time")
            show_minimal_line_chart(sentiment_weekly, "Weekly mentions", height=205)
            if "topic" in questions.columns:
                q_compare = questions[questions["week_start"].isin([start_week, end_week])].groupby(["topic", "week_start"])["frequency"].sum().unstack(fill_value=0)
                for date in [start_week, end_week]:
                    if date not in q_compare.columns:
                        q_compare[date] = 0
                q_compare = q_compare[[start_week, end_week]]
                q_compare.columns = [start_week.strftime("Week of %d %b"), end_week.strftime("Week of %d %b")]
                q_compare["total"] = q_compare.sum(axis=1)
                q_compare = q_compare.sort_values("total", ascending=False).head(8).drop(columns="total").reset_index()
                st.markdown("#### Most common topics: selected weeks")
                q_long = q_compare.melt(id_vars="topic", var_name="week", value_name="mentions")
                show_minimal_bar_chart(q_long, "topic", "mentions", "Mentions", height=260, color_field="week")

        if not competitors.empty and "week_start" in competitors.columns and "image_gap" in competitors.columns:
            gap_compare = competitors[competitors["week_start"].isin([start_week, end_week])].groupby(["image_gap", "week_start"]).size().unstack(fill_value=0)
            for date in [start_week, end_week]:
                if date not in gap_compare.columns:
                    gap_compare[date] = 0
            gap_compare = gap_compare[[start_week, end_week]]
            gap_compare.columns = [start_week.strftime("Week of %d %b"), end_week.strftime("Week of %d %b")]
            gap_compare["total"] = gap_compare.sum(axis=1)
            gap_compare = gap_compare.sort_values("total", ascending=False).head(8).drop(columns="total").reset_index()
            st.markdown("#### Competitor image gaps: selected weeks")
            gap_long = gap_compare.melt(id_vars="image_gap", var_name="week", value_name="observations")
            show_minimal_bar_chart(gap_long, "image_gap", "observations", "Competitor observations", height=260, color_field="week")

        if not results.empty and {"week_start", "variant_name", "impressions", "clicks", "orders"}.issubset(results.columns):
            for col in ["impressions", "clicks", "orders"]:
                results[col] = pd.to_numeric(results[col], errors="coerce").fillna(0)
            test_weekly = results.groupby(["week_start", "variant_name"]).agg(
                impressions=("impressions", "sum"), clicks=("clicks", "sum"), orders=("orders", "sum")
            )
            test_weekly["CTR"] = test_weekly.apply(lambda r: r["clicks"] / r["impressions"] if r["impressions"] else 0, axis=1)
            test_weekly["Order rate"] = test_weekly.apply(lambda r: r["orders"] / r["clicks"] if r["clicks"] else 0, axis=1)
            test_weekly["CTR (%)"] = test_weekly["CTR"] * 100
            test_weekly["Order rate (%)"] = test_weekly["Order rate"] * 100
            st.markdown("#### Test variant CTR progression")
            show_minimal_line_chart(test_weekly["CTR (%)"].unstack("variant_name").sort_index(), "CTR (%)", height=215, zero=False)
            st.markdown("#### Test variant order-rate progression")
            show_minimal_line_chart(test_weekly["Order rate (%)"].unstack("variant_name").sort_index(), "Order rate (%)", height=215, zero=False)
            st.caption("Rates are calculated from the archived weekly totals. A real A/B test should also account for test duration, traffic allocation, and statistical confidence.")

elif page == "Experiment review":
    section_heading("04 · COMPARE", "Experiment review", "Review the example metrics and see what a real shopper test would measure.")
    st.caption("CTR = clicks ÷ impressions. Order rate = orders ÷ clicks.")
    source = data["Test Results"].copy()
    if {"variant_name", "impressions", "clicks", "orders"}.issubset(source.columns):
        for col in ["impressions", "clicks", "orders"]:
            source[col] = pd.to_numeric(source[col], errors="coerce").fillna(0)
        metrics = source.groupby("variant_name", as_index=False).agg(
            impressions=("impressions", "sum"), clicks=("clicks", "sum"), orders=("orders", "sum")
        )
        metrics = calculate_metrics(metrics)
        metrics["CTR (%)"] = metrics["CTR"] * 100
        metrics["Order rate (%)"] = metrics["Order rate"] * 100
        if not metrics.empty:
            variant_columns = st.columns(min(4, len(metrics)))
            for index, (_, row) in enumerate(metrics.iterrows()):
                with variant_columns[index % len(variant_columns)]:
                    st.markdown(
                        f"<div class='variant-card'><div class='variant-name'>{escape(clean(row['variant_name']))}</div>"
                        f"<div class='variant-rate'>{row['CTR']:.2%} CTR</div>"
                        f"<div class='variant-note'>{row['orders']:,.0f} orders · {row['Order rate']:.2%} order rate</div></div>",
                        unsafe_allow_html=True,
                    )
            st.markdown("#### CTR by creative option")
            show_minimal_bar_chart(metrics[["variant_name", "CTR (%)"]], "variant_name", "CTR (%)", "CTR (%)", height=230)
            winner = metrics.sort_values("CTR", ascending=False).iloc[0]
            st.markdown(
                f"<div class='notice-card'><b>Highest CTR in this workbook:</b> {escape(clean(winner['variant_name']))} · {winner['CTR']:.2%}</div>",
                unsafe_allow_html=True,
            )
    st.markdown("<div class='panel'><b>For a real test</b><br>Amazon would show different content versions to separate groups of shoppers and compare actual outcomes. This page displays the workbook values only.</div>", unsafe_allow_html=True)

st.markdown("<div style='margin-top:2rem;color:#778198;font-size:.82rem'>Listing Creative Lab · concept-to-test workflow</div>", unsafe_allow_html=True)
