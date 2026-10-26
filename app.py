import io
import re
from collections import Counter

import streamlit as st
import fitz
from docx import Document

st.set_page_config(
    page_title="NEXUS CV | Resume & ATS Checker",
    page_icon="✦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------- Futuristic UI ----------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
:root { color-scheme: dark; }
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.stApp {
  background:
    radial-gradient(ellipse at 80% 0%, rgba(38, 75, 180, .18), transparent 38%),
    radial-gradient(ellipse at 0% 80%, rgba(104, 43, 220, .12), transparent 32%),
    #050d1b;
  color: #eef4ff;
}
[data-testid="stHeader"] { background: rgba(5,13,27,.82); }
[data-testid="stSidebar"] {
  background: linear-gradient(180deg, #081327 0%, #060d1b 100%);
  border-right: 1px solid rgba(100,145,255,.20);
}
[data-testid="stSidebar"] * { color: #dce7ff; }
.block-container { padding-top: 1.35rem; max-width: 1500px; }
h1,h2,h3 { letter-spacing: -.035em; color: #f5f8ff !important; }
.brand {
  font-size: 1.55rem; font-weight: 800; letter-spacing: -.04em;
  background: linear-gradient(90deg,#a56bff,#39d9ff);
  -webkit-background-clip: text; -webkit-text-fill-color: transparent;
}
.brand-sub { color:#9fb3d9; font-size:.82rem; margin-top:-4px; }
.topbar {
  display:flex; justify-content:space-between; align-items:center; gap:1rem;
  border:1px solid rgba(81,133,255,.24); border-radius:18px;
  padding:12px 17px; background:rgba(10,23,44,.74); margin-bottom:1.4rem;
}
.privacy-pill {
  border:1px solid rgba(0,231,184,.48); color:#53f3ce; border-radius:999px;
  padding:7px 13px; font-size:.78rem; font-weight:700; background:rgba(0,231,184,.06);
}
.hero-kicker { color:#5edfff; font-size:.82rem; font-weight:700; text-transform:uppercase; letter-spacing:.13em; }
.hero-title { font-size:clamp(2rem,3.4vw,3.2rem); line-height:1.08; font-weight:800; margin:.4rem 0 .55rem; }
.gradient-text {
  background:linear-gradient(90deg,#42d8ff,#9b6bff); -webkit-background-clip:text;
  -webkit-text-fill-color:transparent;
}
.muted { color:#9db1d5; }
.panel {
  background:linear-gradient(145deg,rgba(12,28,52,.94),rgba(7,18,36,.96));
  border:1px solid rgba(88,133,225,.25); border-radius:18px; padding:1.15rem;
  box-shadow:0 12px 35px rgba(0,0,0,.12);
}
.panel-title { font-weight:700; font-size:1.02rem; margin-bottom:.85rem; color:#f0f5ff; }
.step {
  display:inline-flex; width:28px; height:28px; align-items:center; justify-content:center;
  border-radius:50%; margin-right:8px; font-weight:800;
  background:linear-gradient(135deg,#8d48ff,#1eb9ff); color:white;
}
.score-wrap {
  border-radius:16px; padding:16px; text-align:center;
  background:radial-gradient(circle at center,rgba(66,91,255,.16),rgba(8,20,39,.88) 68%);
  border:1px solid rgba(95,126,255,.24);
}
.score-number { font-size:3.1rem; font-weight:800; line-height:1.1; color:#f5f8ff; }
.metric-card {
  background:rgba(13,30,54,.9); border:1px solid rgba(91,132,213,.22);
  border-radius:14px; padding:14px; min-height:94px;
}
.metric-label { color:#9eb4da; font-size:.8rem; }
.metric-value { color:#f3f7ff; font-size:1.4rem; font-weight:800; margin-top:4px; }
.keyword {
  display:inline-block; margin:4px 5px 4px 0; padding:6px 10px; border-radius:999px;
  background:rgba(39,205,255,.09); border:1px solid rgba(39,205,255,.35);
  color:#8deaff; font-size:.8rem;
}
.keyword-missing {
  display:inline-block; margin:4px 5px 4px 0; padding:6px 10px; border-radius:999px;
  background:rgba(255,70,105,.08); border:1px solid rgba(255,90,115,.40);
  color:#ffb1c0; font-size:.8rem;
}
.tip {
  background:rgba(13,30,54,.75); border:1px solid rgba(87,127,206,.18);
  border-radius:12px; padding:12px 14px; margin:8px 0;
}
div[data-testid="stFileUploader"] section {
  background:rgba(9,23,44,.68); border:1px dashed #378eea; border-radius:14px;
}
.stTextArea textarea {
  background:#07162a !important; color:#e9f2ff !important;
  border:1px solid rgba(92,139,225,.35) !important; border-radius:13px !important;
}
.stButton > button, .stDownloadButton > button {
  border:0 !important; border-radius:12px !important; min-height:46px !important;
  color:white !important; font-weight:800 !important;
  background:linear-gradient(100deg,#9148ff 0%,#416cff 50%,#10cce9 100%) !important;
  box-shadow:0 7px 22px rgba(68,93,255,.24);
}
.stButton > button:hover, .stDownloadButton > button:hover { filter:brightness(1.12); transform:translateY(-1px); }
div[data-testid="stAlert"] { border-radius:12px; }
hr { border-color:rgba(98,135,205,.2); }
.small-note { font-size:.78rem; color:#91a7cc; }
</style>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown('<div class="brand">✦ NEXUS CV</div><div class="brand-sub">AI Resume & ATS Checker</div>', unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("### ◈ Workspace")
    st.markdown("🔎 **Resume Scanner**")
    st.markdown("📄 Resume Builder · *Coming soon*")
    st.markdown("◷ History · *Not stored*")
    st.markdown("⚙️ Settings")
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown("""
    <div class="panel">
      <div style="font-weight:700;color:#64f4d3">🔒 Private by design</div>
      <p class="small-note">This app does not intentionally save uploaded resumes to a database or disk. Processed content may exist temporarily in application memory.</p>
    </div>
    """, unsafe_allow_html=True)

st.markdown("""
<div class="topbar">
  <div style="font-weight:700;font-size:1rem">Resume Scanner <span class="muted">/ Workspace</span></div>
  <div class="privacy-pill">🔒 &nbsp; PRIVATE BY DESIGN · NO RESUME DATABASE</div>
</div>
<div class="hero-kicker">✦ Your next opportunity starts here</div>
<div class="hero-title">Match your <span class="gradient-text">resume</span> to the role.</div>
<div class="muted" style="font-size:1.05rem;margin-bottom:1.35rem">Instant feedback. Clear improvements. Your career, in focus.</div>
""", unsafe_allow_html=True)

STOPWORDS = {
    "a", "an", "and", "are", "as", "at", "be", "been", "by", "for", "from", "has",
    "have", "in", "is", "it", "of", "on", "or", "our", "that", "the", "their",
    "this", "to", "was", "were", "will", "with", "you", "your", "we", "they",
    "using", "use", "work", "working", "role", "job", "candidate", "ability",
    "skills", "skill", "experience", "required", "preferred", "including"
}

def extract_text(uploaded_file):
    data = uploaded_file.getvalue()
    name = uploaded_file.name.lower()
    if name.endswith(".pdf"):
        with fitz.open(stream=data, filetype="pdf") as pdf:
            return "\n".join(page.get_text() for page in pdf)
    if name.endswith(".docx"):
        doc = Document(io.BytesIO(data))
        parts = [p.text for p in doc.paragraphs]
        for table in doc.tables:
            for row in table.rows:
                parts.append(" | ".join(cell.text for cell in row.cells))
        return "\n".join(parts)
    raise ValueError("Please upload a PDF or DOCX file.")

def words(text):
    found = re.findall(r"[A-Za-z][A-Za-z0-9+#.-]{1,}", text.lower())
    return [w.strip(".-") for w in found if len(w.strip(".-")) > 1 and w not in STOPWORDS]

def keyword_analysis(resume_text, job_text):
    resume_words = set(words(resume_text))
    job_counts = Counter(words(job_text))
    keywords = [word for word, _ in job_counts.most_common(40)]
    matched = [word for word in keywords if word in resume_words]
    missing = [word for word in keywords if word not in resume_words]
    score = round(100 * len(matched) / max(len(keywords), 1))
    return score, matched, missing

def make_report(score, matched, missing, suggestions):
    return "\n".join([
        "NEXUS CV — RESUME & ATS CHECK REPORT",
        "=" * 38,
        f"Estimated keyword match: {score}/100",
        "",
        "Matched keywords:", ", ".join(matched) or "None detected",
        "",
        "Potentially missing keywords:", ", ".join(missing[:25]) or "None detected",
        "",
        "Improvement suggestions:", *[f"- {s}" for s in suggestions],
        "",
        "This is an estimated keyword comparison, not a score from an employer's ATS."
    ])

left, right = st.columns([0.93, 1.07], gap="large")

with left:
    st.markdown('<div class="panel-title"><span class="step">1</span> Upload your resume</div>', unsafe_allow_html=True)
    resume_file = st.file_uploader("PDF or DOCX · processed temporarily", type=["pdf", "docx"], label_visibility="collapsed")
    if resume_file:
        st.success(f"✓ {resume_file.name} is ready for analysis")
    st.markdown('<div style="height:12px"></div><div class="panel-title"><span class="step">2</span> Paste the job description</div>', unsafe_allow_html=True)
    job_text = st.text_area(
        "Job description",
        height=220,
        placeholder="Paste the job description here — responsibilities, skills, qualifications...",
        label_visibility="collapsed",
        max_chars=15000
    )
    st.markdown('<div class="small-note">Tip: Include the complete job description for a more useful comparison.</div>', unsafe_allow_html=True)
    analyze = st.button("✦  Analyze resume  →", type="primary", use_container_width=True)
    st.markdown("""
    <div class="small-note" style="margin-top:12px">🔐 Files are processed for this session. Avoid uploading highly sensitive information unless you are comfortable with the hosting provider's privacy terms.</div>
    """, unsafe_allow_html=True)

with right:
    st.markdown('<div class="panel-title">◉ ATS Match Insights</div>', unsafe_allow_html=True)
    if analyze:
        if not resume_file:
            st.error("Upload a PDF or DOCX resume first.")
        elif not job_text.strip():
            st.error("Paste the job description first.")
        else:
            try:
                resume_text = extract_text(resume_file)
                if not resume_text.strip():
                    st.error("No readable text found. Scanned PDFs may need OCR before upload.")
                else:
                    score, matched, missing = keyword_analysis(resume_text, job_text)
                    suggestions = []
                    if len(resume_text.split()) < 150:
                        suggestions.append("Your resume looks brief. Add relevant education, projects, skills, or achievements where accurate.")
                    if "@" not in resume_text:
                        suggestions.append("Check that your contact email is clearly visible.")
                    if not re.search(r"\b(education|degree|bachelor|university|college)\b", resume_text, re.I):
                        suggestions.append("Add a clearly labelled Education section if relevant.")
                    if not re.search(r"\b(project|projects)\b", resume_text, re.I):
                        suggestions.append("Include relevant projects and explain your contribution where applicable.")
                    if missing:
                        suggestions.append("Review missing keywords and include only those that truthfully reflect your skills.")
                    suggestions.append("Use clear headings, consistent dates, and simple formatting; avoid complex tables and text boxes.")
                    suggestions.append("Verify spelling, contact details, dates, and all claims before applying.")

                    label = "Strong keyword overlap" if score >= 70 else "Moderate keyword overlap" if score >= 40 else "Low keyword overlap"
                    st.markdown(f"""
                    <div class="score-wrap">
                      <div class="muted" style="font-size:.85rem">ESTIMATED KEYWORD MATCH</div>
                      <div class="score-number">{score}<span style="font-size:1.15rem;color:#9db1d5"> / 100</span></div>
                      <div style="color:#57eacb;font-weight:700;margin-top:5px">✦ {label}</div>
                    </div>
                    """, unsafe_allow_html=True)
                    st.progress(score / 100)
                    a, b, c = st.columns(3)
                    with a:
                        st.markdown(f'<div class="metric-card"><div class="metric-label">Keywords matched</div><div class="metric-value">{len(matched)}</div></div>', unsafe_allow_html=True)
                    with b:
                        st.markdown(f'<div class="metric-card"><div class="metric-label">Potential gaps</div><div class="metric-value">{len(missing)}</div></div>', unsafe_allow_html=True)
                    with c:
                        st.markdown(f'<div class="metric-card"><div class="metric-label">Resume length</div><div class="metric-value">{len(resume_text.split())} <span style="font-size:.8rem">words</span></div></div>', unsafe_allow_html=True)

                    st.markdown("#### ✓ Matched keywords")
                    if matched:
                        st.markdown("".join(f'<span class="keyword">{w}</span>' for w in matched[:20]), unsafe_allow_html=True)
                    else:
                        st.write("No matching keywords detected.")

                    st.markdown("#### ⚠ Potentially missing keywords")
                    if missing:
                        st.markdown("".join(f'<span class="keyword-missing">{w}</span>' for w in missing[:15]), unsafe_allow_html=True)
                    else:
                        st.write("No major keyword gaps detected.")

                    st.markdown("#### ✨ Recommended improvements")
                    for item in suggestions:
                        st.markdown(f'<div class="tip">✧ &nbsp; {item}</div>', unsafe_allow_html=True)

                    st.download_button(
                        "⬇  Download analysis report (.txt)",
                        data=make_report(score, matched, missing, suggestions).encode("utf-8"),
                        file_name="nexus_cv_analysis_report.txt",
                        mime="text/plain",
                        use_container_width=True
                    )
                    st.caption("Keyword overlap is an estimate, not a guaranteed ATS pass score. This version does not automatically rewrite your resume.")
                    del resume_text
            except Exception:
                st.error("We couldn't read this file. Try a valid text-based PDF or DOCX. Document contents are not shown in error messages.")
    else:
        st.markdown("""
        <div class="score-wrap" style="padding:30px 15px">
          <div style="font-size:2.4rem;margin-bottom:8px">◉</div>
          <div style="font-size:1.2rem;font-weight:700">Your insights will appear here</div>
          <div class="muted" style="margin-top:8px">Upload a resume and add a job description, then select Analyze resume.</div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("#### ✨ What you'll get")
        for item in [
            "Estimated keyword overlap with the job description",
            "Matched and potentially missing keywords",
            "Practical suggestions to improve clarity",
            "A downloadable analysis report",
        ]:
            st.markdown(f'<div class="tip">✓ &nbsp; {item}</div>', unsafe_allow_html=True)

st.markdown("---")
st.markdown('<div class="small-note">✦ NEXUS CV · Privacy-conscious resume analysis · No login or resume database in this prototype. Hosting providers may retain technical logs; review their policies before production use.</div>', unsafe_allow_html=True)
