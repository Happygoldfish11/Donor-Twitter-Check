"""
Twitter/X Profile Finder
A tool for matching people from a spreadsheet to their Twitter/X profiles.
"""

import streamlit as st
import time
import pandas as pd
import sys, os
sys.path.insert(0, os.path.dirname(__file__))

from src.matcher import Person, find_twitter
from src.spreadsheet import load_persons, write_results

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Twitter Profile Finder",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Styles ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500&family=IBM+Plex+Sans:wght@300;400;500;600&display=swap');

html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }

.stApp { background: #0a0f1e; color: #e2e8f0; }

.hero-wrap {
    border: 1px solid #1e293b;
    border-radius: 12px;
    padding: 2.5rem 2rem 2rem;
    background: linear-gradient(135deg, #0f172a 0%, #0a0f1e 100%);
    margin-bottom: 1.5rem;
    position: relative;
    overflow: hidden;
}
.hero-wrap::before {
    content: '';
    position: absolute;
    top: -40px; right: -40px;
    width: 200px; height: 200px;
    background: radial-gradient(circle, rgba(56,189,248,0.08) 0%, transparent 70%);
    pointer-events: none;
}
.hero-title {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.7rem;
    font-weight: 500;
    color: #f1f5f9;
    letter-spacing: -0.5px;
    margin: 0 0 0.4rem;
}
.hero-sub {
    color: #64748b;
    font-size: 0.9rem;
    margin: 0;
    font-weight: 300;
}
.tag {
    display: inline-block;
    background: #0ea5e920;
    color: #38bdf8;
    border: 1px solid #0ea5e940;
    border-radius: 4px;
    padding: 1px 8px;
    font-size: 0.72rem;
    font-family: 'IBM Plex Mono', monospace;
    margin-bottom: 1rem;
}
.card {
    background: #0f172a;
    border: 1px solid #1e293b;
    border-radius: 10px;
    padding: 1.25rem 1.5rem;
    margin: 0.75rem 0;
}
.metric-row {
    display: flex;
    gap: 0.75rem;
    flex-wrap: wrap;
    margin: 1rem 0;
}
.metric {
    flex: 1;
    min-width: 110px;
    background: #0f172a;
    border: 1px solid #1e293b;
    border-radius: 8px;
    padding: 1rem 0.75rem;
    text-align: center;
}
.metric .num {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.8rem;
    font-weight: 500;
    line-height: 1;
    margin-bottom: 4px;
}
.metric .lbl {
    font-size: 0.7rem;
    text-transform: uppercase;
    letter-spacing: 1px;
    color: #475569;
}
.result-row {
    display: flex;
    align-items: center;
    gap: 0.75rem;
    padding: 0.6rem 1rem;
    border-radius: 6px;
    border: 1px solid #1e293b;
    margin: 0.3rem 0;
    background: #0f172a;
    transition: border-color 0.15s;
}
.result-row:hover { border-color: #334155; }
.badge {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.65rem;
    padding: 2px 7px;
    border-radius: 3px;
    font-weight: 500;
    white-space: nowrap;
    min-width: 72px;
    text-align: center;
}
.badge-HIGH     { background: #14532d40; color: #4ade80; border: 1px solid #16a34a50; }
.badge-MEDIUM   { background: #71350040; color: #fbbf24; border: 1px solid #d9770050; }
.badge-LOW      { background: #7c2d1240; color: #fb923c; border: 1px solid #ea580c50; }
.badge-NOMATCH  { background: #7f1d1d40; color: #f87171; border: 1px solid #dc262650; }
.score-bar-wrap {
    flex: 1;
    background: #1e293b;
    border-radius: 3px;
    height: 4px;
    min-width: 60px;
}
.score-bar { height: 4px; border-radius: 3px; transition: width 0.4s; }
.person-name { font-weight: 500; font-size: 0.9rem; min-width: 160px; color: #e2e8f0; }
.handle-link { font-family: 'IBM Plex Mono', monospace; font-size: 0.8rem; color: #38bdf8; text-decoration: none; }
.score-num { font-family: 'IBM Plex Mono', monospace; font-size: 0.75rem; color: #475569; min-width: 28px; text-align: right; }
.info-box {
    background: #0c1a2e;
    border-left: 3px solid #0ea5e9;
    border-radius: 0 6px 6px 0;
    padding: 0.6rem 0.9rem;
    font-size: 0.82rem;
    color: #64748b;
    margin: 0.5rem 0;
}
.warn-box {
    background: #1a1200;
    border-left: 3px solid #d97706;
    border-radius: 0 6px 6px 0;
    padding: 0.6rem 0.9rem;
    font-size: 0.82rem;
    color: #92400e;
    margin: 0.4rem 0;
}
.breakdown-pill {
    display: inline-block;
    background: #1e293b;
    color: #94a3b8;
    border-radius: 4px;
    padding: 1px 6px;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.68rem;
    margin: 1px;
}
a { color: #38bdf8; }
</style>
""", unsafe_allow_html=True)


# ── Hero header ───────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero-wrap">
    <div class="tag">v1.0 · powered by Claude web search</div>
    <div class="hero-title">Twitter / X Profile Finder</div>
    <p class="hero-sub">Upload a spreadsheet of names, addresses, and employers —<br>
    get back matched Twitter profiles with confidence scores.</p>
</div>
""", unsafe_allow_html=True)


# ── Sidebar — API keys ────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Configuration")
    st.markdown("---")

    anthropic_key = st.text_input(
        "Anthropic API Key",
        type="password",
        placeholder="sk-ant-...",
        help="Used to power Claude's web search. Get one at console.anthropic.com"
    )

    st.markdown("---")
    st.markdown("**Optional: additional search backends**")

    serpapi_key = st.text_input(
        "SerpAPI Key (optional)",
        type="password",
        placeholder="For Google search fallback",
    )

    st.markdown("---")
    st.markdown("#### How confidence works")
    st.markdown("""
<div style='font-size:0.8rem; color:#475569; line-height:1.8'>
<span style='color:#4ade80'>✓ HIGH ≥ 75</span> — Strong match<br>
<span style='color:#fbbf24'>~ MED 45–74</span> — Likely match<br>
<span style='color:#fb923c'>! LOW 20–44</span> — Possible match<br>
<span style='color:#f87171'>✗ NONE &lt;20</span> — No match found
</div>
""", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("""
<div style='font-size:0.75rem; color:#334155'>
<b>Scoring signals:</b><br>
+40 Exact name match<br>
+25 Employer in bio<br>
+20 City in profile<br>
+10 State in profile<br>
+15 Bio keyword overlap<br>
+5 Verified account<br>
+5 Account age > 2yr
</div>
""", unsafe_allow_html=True)


# ── Upload section ────────────────────────────────────────────────────────────
st.markdown("### 1. Upload your spreadsheet")
st.markdown("""
<div class="info-box">
Required columns: <b>First Name</b>, <b>Last Name</b> &nbsp;·&nbsp;
Recommended: <b>Address</b>, <b>Employer</b> (improves confidence scores significantly)
</div>
""", unsafe_allow_html=True)

uploaded = st.file_uploader("", type=["xlsx", "csv"], label_visibility="collapsed")

if uploaded:
    try:
        df_orig, persons, warnings = load_persons(uploaded)
    except ValueError as e:
        st.error(f"❌ {e}")
        st.stop()

    for w in warnings:
        st.markdown(f'<div class="warn-box">⚠️ {w}</div>', unsafe_allow_html=True)

    st.markdown(f"**{len(persons)} people loaded.** Preview:")
    st.dataframe(df_orig.head(5), use_container_width=True, hide_index=True)

    # ── Run button ────────────────────────────────────────────────────────────
    st.markdown("### 2. Run the search")

    if not anthropic_key and not serpapi_key:
        st.markdown('<div class="warn-box">⚠️ Add your Anthropic API key in the sidebar to enable search.</div>', unsafe_allow_html=True)

    col1, col2 = st.columns([3, 1])
    with col1:
        run = st.button("🔍 Find Twitter Profiles", type="primary", use_container_width=True,
                        disabled=(not anthropic_key and not serpapi_key))
    with col2:
        dry_run = st.button("🧪 Test (first 3 only)", use_container_width=True,
                            disabled=(not anthropic_key and not serpapi_key))

    if run or dry_run:
        sample = persons[:3] if dry_run else persons
        results = []

        st.markdown("### 3. Results")
        progress = st.progress(0, text="Starting search...")
        status = st.empty()
        live_results = st.empty()

        for i, person in enumerate(sample):
            status.markdown(
                f'<div class="info-box">Searching for <b>{person.full_name}</b> ({i+1} of {len(sample)})...</div>',
                unsafe_allow_html=True
            )

            result = find_twitter(
                person,
                anthropic_key=anthropic_key or "",
                serpapi_key=serpapi_key or "",
            )
            results.append(result)

            progress.progress((i + 1) / len(sample),
                              text=f"{i+1}/{len(sample)} complete")

            # Live results table
            rows_html = ""
            for r in results:
                p = r.profile
                conf_class = r.confidence.replace(" ", "")
                bar_color = {"HIGH": "#4ade80", "MEDIUM": "#fbbf24",
                             "LOW": "#fb923c", "NO MATCH": "#f87171"}.get(r.confidence, "#6b7280")
                handle_html = (
                    f'<a class="handle-link" href="{p.profile_url}" target="_blank">{p.handle}</a>'
                    if p else '<span style="color:#334155">—</span>'
                )
                breakdown_html = " ".join(
                    f'<span class="breakdown-pill">{k.replace("_"," ")} +{v}</span>'
                    for k, v in r.score_breakdown.items()
                ) if r.score_breakdown else ""

                rows_html += f"""
                <div class="result-row">
                    <span class="badge badge-{conf_class}">{r.emoji} {r.confidence}</span>
                    <span class="person-name">{r.person.full_name}</span>
                    <span>{handle_html}</span>
                    <div style="flex:1">
                        <div class="score-bar-wrap">
                            <div class="score-bar" style="width:{r.score}%;background:{bar_color}"></div>
                        </div>
                    </div>
                    <span class="score-num">{r.score}</span>
                </div>
                {f'<div style="padding:0 1rem 0.25rem;font-size:0.72rem;color:#475569">{breakdown_html}</div>' if breakdown_html else ''}
                """

            live_results.markdown(rows_html, unsafe_allow_html=True)
            time.sleep(0.1)

        status.empty()
        progress.empty()

        # ── Summary stats ─────────────────────────────────────────────────────
        counts = {"HIGH": 0, "MEDIUM": 0, "LOW": 0, "NO MATCH": 0}
        for r in results:
            counts[r.confidence] += 1

        st.markdown(f"""
        <div class="metric-row">
            <div class="metric">
                <div class="num" style="color:#e2e8f0">{len(results)}</div>
                <div class="lbl">Searched</div>
            </div>
            <div class="metric">
                <div class="num" style="color:#4ade80">{counts['HIGH']}</div>
                <div class="lbl">High confidence</div>
            </div>
            <div class="metric">
                <div class="num" style="color:#fbbf24">{counts['MEDIUM']}</div>
                <div class="lbl">Medium</div>
            </div>
            <div class="metric">
                <div class="num" style="color:#fb923c">{counts['LOW']}</div>
                <div class="lbl">Low</div>
            </div>
            <div class="metric">
                <div class="num" style="color:#f87171">{counts['NO MATCH']}</div>
                <div class="lbl">No match</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # ── Download ──────────────────────────────────────────────────────────
        st.markdown("### 4. Download results")
        uploaded.seek(0)
        out_buf = write_results(uploaded, results)
        out_name = uploaded.name.replace(".csv", "").replace(".xlsx", "") + "_twitter_results.xlsx"

        st.download_button(
            label="⬇️ Download Annotated Spreadsheet",
            data=out_buf,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )

        if dry_run and len(persons) > 3:
            st.info(f"Test run complete. Click 'Find Twitter Profiles' to process all {len(persons)} people.")

else:
    # ── Empty state ───────────────────────────────────────────────────────────
    st.markdown("""
    <div class="card" style="text-align:center;padding:3rem 2rem">
        <div style="font-size:2.5rem;margin-bottom:1rem">📋</div>
        <div style="font-size:1rem;color:#475569;margin-bottom:0.5rem">No file uploaded yet</div>
        <div style="font-size:0.82rem;color:#334155">
            Upload an .xlsx or .csv file with First Name, Last Name, Address, and Employer columns.<br>
            The more columns you provide, the higher your confidence scores will be.
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Sample template download ──────────────────────────────────────────────
    import io as _io
    sample_df = pd.DataFrame({
        "First Name": ["Jane", "Michael", "Sarah"],
        "Last Name":  ["Smith", "Chen", "Johnson"],
        "Address":    ["123 Main St, New York, NY 10001",
                       "456 Oak Ave, San Francisco, CA 94102",
                       "789 Pine Rd, Chicago, IL 60601"],
        "Employer":   ["Goldman Sachs", "Google", "United Airlines"],
    })
    buf = _io.BytesIO()
    sample_df.to_excel(buf, index=False)
    buf.seek(0)

    st.download_button(
        "📥 Download Sample Template",
        data=buf,
        file_name="twitter_finder_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
