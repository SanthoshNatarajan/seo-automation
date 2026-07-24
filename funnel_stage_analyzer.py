"""
Funnel Stage Analyzer  |  SEO Automation Series - Episode 3
-----------------------------------------------------------
Ask it for a website. It fetches the whole site (blogs + pages) from the
sitemap, classifies every URL as TOFU / MOFU / BOFU by search intent, and
opens an HTML presentation:

    Slide 1  - graph view of the TOFU / MOFU / BOFU mix
    Slide 2  - number of blogs and pages, broken down by intent
    Slide 3  - every URL, classified and filterable

100% FREE. No API key, no subscription. Runs on keyword intent rules.

Setup (one time):
    pip install requests beautifulsoup4

Run:
    python funnel_stage_analyzer.py
    -> it will ask you to type a website, e.g. systechgroup.in
"""

import csv
import re
import sys
import time
import webbrowser
from collections import Counter
from datetime import date

import requests
from bs4 import BeautifulSoup

HEADERS = {"User-Agent": "Mozilla/5.0 (FunnelStageAnalyzer/1.0)"}
FETCH_PAGE_CONTENT = True   # read each page's title+description for better accuracy
REQUEST_DELAY = 0.3         # polite pause between page fetches (seconds)

# Target mix for a lead-generating site. Tune to your own strategy.
TARGET_MIX = {"TOFU": 55, "MOFU": 30, "BOFU": 15}

# ------------------- INTENT RULES (this is the strategy) --------------------
# Checked against URL + title + meta description. Most specific stage wins:
# BOFU (ready to buy) -> MOFU (comparing) -> TOFU (just learning).

BOFU_SIGNALS = [
    "best institute", "best training", "which is the best", "top institute",
    "cost of", "price", "pricing", "fees", "with placement", "placement",
    "enroll", "enrol", "admission", "near me", "demo class", "free trial",
    "book a", "contact us", "alternatives to", "vs ", " vs", "-vs-", "compare",
]
MOFU_SIGNALS = [
    "course after", "courses after", "interview question", "salary", "jobs",
    "job opportunities", "career", "how to become", "certification", "roadmap",
    "difference between", "scope", "benefits of", "why choose", "is it worth",
    "which course", "top courses", "trending courses", "professional courses",
    "for beginners", "eligibility", "syllabus", "duration",
]
# City names only signal buying intent when paired with a course/institute word,
# so "digital marketing course in chennai" is BOFU but "salary in chennai" is not.
CITY_WORDS = ["chennai", "coimbatore", "trichy", "tiruchirappalli", "madurai",
              "bangalore", "in tamil"]
COMMERCIAL_WORDS = ["course", "training", "institute", "class", "coaching",
                    "certification", "program", "bootcamp"]

def classify(text):
    t = text.lower()
    city_buy = (any(c in t for c in CITY_WORDS)
                and any(w in t for w in COMMERCIAL_WORDS)
                and not any(x in t for x in ["salary", "interview", "jobs", "job "]))
    if city_buy or any(s in t for s in BOFU_SIGNALS):
        return "BOFU", "buying signal: cost / placement / best-in-city / comparison"
    if any(s in t for s in MOFU_SIGNALS):
        return "MOFU", "consideration signal: career / salary / course choice"
    return "TOFU", "no buying signal: educational / awareness topic"

# ----------------------------- SITEMAP --------------------------------------

def normalize(site):
    site = site.strip().rstrip("/")
    if not site.startswith("http"):
        site = "https://" + site
    return site

def find_sitemaps(site):
    """Return list of child sitemap URLs, following any sitemap index."""
    candidates = [f"{site}/sitemap_index.xml", f"{site}/sitemap.xml"]
    for url in candidates:
        try:
            xml = requests.get(url, headers=HEADERS, timeout=20).text
            locs = re.findall(r"<loc>(.*?)</loc>", xml)
            if not locs:
                continue
            # If entries are themselves sitemaps, it's an index -> return them.
            child = [l for l in locs if l.endswith(".xml")]
            return child if child else [url]
        except Exception:
            continue
    return []

def collect_urls(site):
    """Return list of dicts: {url, kind} where kind is 'blog' or 'page'."""
    sitemaps = find_sitemaps(site)
    if not sitemaps:
        return []
    items = {}
    for sm in sitemaps:
        try:
            xml = requests.get(sm, headers=HEADERS, timeout=20).text
        except Exception:
            continue
        locs = [l for l in re.findall(r"<loc>(.*?)</loc>", xml) if not l.endswith(".xml")]
        # Decide blog vs page by which sitemap it came from, else by URL.
        name = sm.lower()
        if "post" in name or "blog" in name:
            kind = "blog"
        elif "page" in name or "product" in name:
            kind = "page"
        else:
            kind = None
        for u in locs:
            k = kind or ("blog" if re.search(r"/blog[-/]|/post[-/]", u) else "page")
            items[u] = k
    return [{"url": u, "kind": k} for u, k in sorted(items.items())]

def get_signals(url):
    """Fetch title + meta description for sharper classification."""
    try:
        soup = BeautifulSoup(requests.get(url, headers=HEADERS, timeout=15).text,
                             "html.parser")
        title = soup.title.get_text(strip=True) if soup.title else ""
        meta = soup.find("meta", attrs={"name": "description"})
        desc = meta["content"].strip() if meta and meta.get("content") else ""
        return f"{title} {desc}"
    except Exception:
        return ""

# ----------------------------- HTML PRESENTATION ----------------------------

def build_html(site, rows, out_path):
    total = len(rows)
    overall = Counter(r["stage"] for r in rows)
    pct = {s: round(100 * overall.get(s, 0) / total, 1) if total else 0
           for s in ("TOFU", "MOFU", "BOFU")}

    def split(kind):
        c = Counter(r["stage"] for r in rows if r["kind"] == kind)
        return {"TOFU": c.get("TOFU", 0), "MOFU": c.get("MOFU", 0),
                "BOFU": c.get("BOFU", 0), "total": sum(c.values())}
    blogs, pages = split("blog"), split("page")

    import json
    data = {
        "site": site, "date": date.today().isoformat(), "total": total,
        "pct": pct, "counts": dict(overall), "target": TARGET_MIX,
        "blogs": blogs, "pages": pages,
        "rows": [{"url": r["url"], "kind": r["kind"], "stage": r["stage"],
                  "reason": r["reason"]} for r in rows],
    }
    html = TEMPLATE.replace("__DATA__", json.dumps(data))
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)

TEMPLATE = r"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Funnel Stage Analysis</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
:root{--tofu:#3b82f6;--mofu:#f59e0b;--bofu:#ef4444;--ink:#0f172a;--mut:#64748b;--bg:#eef2f7;--card:#fff;--line:#e2e8f0}
*{box-sizing:border-box;margin:0}
body{font-family:'Segoe UI',system-ui,sans-serif;background:var(--bg);color:var(--ink);padding:28px 16px}
.wrap{max-width:1040px;margin:0 auto}
.slide{background:var(--card);border:1px solid var(--line);border-top:5px solid #1e40af;border-radius:14px;padding:26px 30px;margin-bottom:22px;box-shadow:0 2px 10px rgba(15,23,42,.05)}
.eyebrow{font-size:12px;letter-spacing:.12em;text-transform:uppercase;color:#1e40af;font-weight:700}
h1{font-size:24px;margin:2px 0 2px}h2{font-size:19px;margin:0 0 14px}
.sub{color:var(--mut);font-size:13px;margin-bottom:16px}
.grid3{display:grid;grid-template-columns:repeat(3,1fr);gap:14px}
.kpi{border:1px solid var(--line);border-radius:12px;padding:16px}
.kpi .s{font-size:12px;text-transform:uppercase;letter-spacing:.06em;font-weight:700}
.kpi .n{font-size:34px;font-weight:800;line-height:1.1}
.kpi .m{font-size:12px;color:var(--mut)}
.gp{color:#16a34a;font-size:12px}.gn{color:#dc2626;font-size:12px}
.charts{display:grid;grid-template-columns:1fr 1fr;gap:20px;align-items:center}
.legend{font-size:12px;color:var(--mut);margin-top:10px;line-height:1.7}
.dot{display:inline-block;width:10px;height:10px;border-radius:2px;margin-right:5px;vertical-align:middle}
table{width:100%;border-collapse:collapse;font-size:13px;margin-top:6px}
th,td{text-align:left;padding:9px 10px;border-bottom:1px solid var(--line)}
th{color:var(--mut);font-size:11px;text-transform:uppercase;letter-spacing:.05em}
td a{color:#2563eb;text-decoration:none;word-break:break-all}
.pill{padding:2px 9px;border-radius:999px;font-size:11px;font-weight:800;color:#fff}
.p-TOFU{background:var(--tofu)}.p-MOFU{background:var(--mofu)}.p-BOFU{background:var(--bofu)}
.tag{font-size:11px;color:var(--mut);text-transform:uppercase}
.filters button{border:1px solid #cbd5e1;background:#fff;border-radius:999px;padding:5px 14px;margin:0 6px 6px 0;cursor:pointer;font-size:13px}
.filters button.on{background:var(--ink);color:#fff;border-color:var(--ink)}
.split{display:grid;grid-template-columns:1fr 1fr;gap:18px}
.mini{border:1px solid var(--line);border-radius:12px;padding:16px}
.mini h3{font-size:15px;margin:0 0 4px}.mini .big{font-size:30px;font-weight:800}
.bar{height:10px;border-radius:6px;background:var(--line);overflow:hidden;display:flex;margin-top:10px}
.bar i{display:block;height:100%}
@media(max-width:820px){.charts,.split,.grid3{grid-template-columns:1fr}}
</style></head><body><div class="wrap">

<div class="slide">
  <div class="eyebrow">SEO Automation Series - Episode 3</div>
  <h1>Funnel Stage Analysis</h1>
  <div class="sub"><b id="site"></b> - <span id="total"></span> URLs analysed - <span id="date"></span></div>
  <div class="grid3" id="kpis"></div>
</div>

<div class="slide">
  <div class="eyebrow">Slide 1</div><h2>The content mix - TOFU / MOFU / BOFU</h2>
  <div class="charts">
    <div><canvas id="donut" height="240"></canvas></div>
    <div><canvas id="bars" height="240"></canvas>
      <div class="legend">
        <span><i class="dot" style="background:var(--tofu)"></i><b>TOFU</b> - awareness (informational). Builds traffic.</span><br>
        <span><i class="dot" style="background:var(--mofu)"></i><b>MOFU</b> - consideration (commercial). Educates & nurtures.</span><br>
        <span><i class="dot" style="background:var(--bofu)"></i><b>BOFU</b> - decision (commercial + transactional). Drives leads.</span>
      </div>
    </div>
  </div>
</div>

<div class="slide">
  <div class="eyebrow">Slide 2</div><h2>Blogs vs Pages - by intent</h2>
  <div class="split">
    <div class="mini"><h3>Blogs</h3><div class="big" id="blogTotal"></div>
      <div class="tag">posts</div><div class="bar" id="blogBar"></div>
      <table id="blogTbl"></table></div>
    <div class="mini"><h3>Pages</h3><div class="big" id="pageTotal"></div>
      <div class="tag">pages</div><div class="bar" id="pageBar"></div>
      <table id="pageTbl"></table></div>
  </div>
</div>

<div class="slide">
  <div class="eyebrow">Slide 3</div><h2>Every URL, classified</h2>
  <div class="filters" id="filters">
    <button class="on" data-f="ALL">All</button>
    <button data-f="TOFU">TOFU</button><button data-f="MOFU">MOFU</button>
    <button data-f="BOFU">BOFU</button>
    <button data-f="blog">Blogs</button><button data-f="page">Pages</button></div>
  <table><thead><tr><th style="width:44px">#</th><th>URL</th><th>Type</th><th>Stage</th><th>Why</th></tr></thead>
  <tbody id="tbody"></tbody></table>
</div>
</div>
<script>
const D=__DATA__,C={TOFU:'#3b82f6',MOFU:'#f59e0b',BOFU:'#ef4444'};
site.textContent=D.site;total.textContent=D.total;date.textContent=D.date;
kpis.innerHTML=['TOFU','MOFU','BOFU'].map(s=>{
 const g=+(D.pct[s]-D.target[s]).toFixed(1),cls=g>=0?'gp':'gn',sg=g>=0?'+':'';
 return `<div class="kpi"><div class="s" style="color:${C[s]}">${s}</div>
 <div class="n">${D.pct[s]}%</div><div class="m">${D.counts[s]||0} URLs</div>
 <div class="${cls}">${sg}${g} pts vs ${D.target[s]}% target</div></div>`}).join('');
new Chart(donut,{type:'doughnut',data:{labels:['TOFU','MOFU','BOFU'],
 datasets:[{data:['TOFU','MOFU','BOFU'].map(s=>D.pct[s]),backgroundColor:Object.values(C),borderWidth:2}]},
 options:{plugins:{legend:{position:'bottom'},title:{display:true,text:'Share of content (%)'}},cutout:'60%'}});
new Chart(bars,{type:'bar',data:{labels:['TOFU','MOFU','BOFU'],
 datasets:[{label:'Your site %',data:['TOFU','MOFU','BOFU'].map(s=>D.pct[s]),backgroundColor:Object.values(C)},
 {label:'Target %',data:['TOFU','MOFU','BOFU'].map(s=>D.target[s]),backgroundColor:'#cbd5e1'}]},
 options:{plugins:{legend:{position:'bottom'}},scales:{y:{beginAtZero:true,max:100}}}});
function fill(o,totalEl,barEl,tblEl){
 totalEl.textContent=o.total;
 barEl.innerHTML=['TOFU','MOFU','BOFU'].map(s=>{
  const w=o.total?100*o[s]/o.total:0;return `<i style="width:${w}%;background:${C[s]}"></i>`}).join('');
 tblEl.innerHTML=`<tr><th>Stage</th><th>Count</th><th>Share</th></tr>`+
  ['TOFU','MOFU','BOFU'].map(s=>{const p=o.total?(100*o[s]/o.total).toFixed(1):0;
  return `<tr><td><span class="pill p-${s}">${s}</span></td><td>${o[s]}</td><td>${p}%</td></tr>`}).join('');}
fill(D.blogs,blogTotal,blogBar,blogTbl);
fill(D.pages,pageTotal,pageBar,pageTbl);
function render(f){tbody.innerHTML=D.rows.filter(r=>f==='ALL'||r.stage===f||r.kind===f)
 .map((r,i)=>`<tr><td style="color:var(--mut);font-weight:700">${i+1}</td>
 <td><a href="${r.url}" target="_blank">${r.url.replace(D.site,'')}</a></td>
 <td class="tag">${r.kind}</td><td><span class="pill p-${r.stage}">${r.stage}</span></td>
 <td>${r.reason}</td></tr>`).join('');}
document.querySelectorAll('.filters button').forEach(b=>b.onclick=()=>{
 document.querySelectorAll('.filters button').forEach(x=>x.classList.remove('on'));
 b.classList.add('on');render(b.dataset.f);});
render('ALL');
</script></body></html>"""

# ----------------------------- MAIN -----------------------------------------

def main():
    print("=" * 56)
    print("  Funnel Stage Analyzer  -  SEO Automation Series Ep 3")
    print("=" * 56)
    site = input("\nEnter website (e.g. systechgroup.in): ").strip()
    if not site:
        sys.exit("No website entered.")
    site = normalize(site)
    print(f"\nAnalysing {site} ...")

    items = collect_urls(site)
    if not items:
        sys.exit("Could not find a sitemap. Try the full URL, or check "
                 "yoursite.com/sitemap.xml exists.")
    print(f"Found {len(items)} URLs "
          f"({sum(i['kind']=='blog' for i in items)} blogs, "
          f"{sum(i['kind']=='page' for i in items)} pages)")

    rows = []
    for n, it in enumerate(items, 1):
        text = it["url"].replace("-", " ").replace("/", " ")
        if FETCH_PAGE_CONTENT:
            text += " " + get_signals(it["url"])
            time.sleep(REQUEST_DELAY)
        stage, reason = classify(text)
        rows.append({"url": it["url"], "kind": it["kind"],
                     "stage": stage, "reason": reason})
        if n % 20 == 0:
            print(f"  classified {n}/{len(items)}")

    domain = re.sub(r"https?://", "", site).replace("/", "_")
    html_path = f"funnel_report_{domain}.html"
    csv_path = f"funnel_report_{domain}.csv"

    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["url", "kind", "stage", "reason"])
        w.writeheader(); w.writerows(rows)

    build_html(site, rows, html_path)

    c = Counter(r["stage"] for r in rows)
    print("\n" + "-" * 40)
    print(f"  TOFU {c['TOFU']}   MOFU {c['MOFU']}   BOFU {c['BOFU']}")
    print("-" * 40)
    print(f"\nSaved: {html_path}")
    print(f"Saved: {csv_path}")
    try:
        webbrowser.open("file://" + __import__("os").path.abspath(html_path))
        print("Opening the presentation in your browser...")
    except Exception:
        print("Open the HTML file above in any browser.")

if __name__ == "__main__":
    main()
