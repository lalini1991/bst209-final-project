#!/usr/bin/env python3
"""Generate funding_dashboard.html — a read-only artifact view of Funding_Dashboard.xlsx.

The xlsx remains the single source of truth; this page is a presentation layer.
Run:  python3 scripts/generate_dashboard.py
"""
import openpyxl, html
from collections import Counter

XLSX = "/home/user/bst209-final-project/Funding_Dashboard.xlsx"
OUT  = "/home/user/bst209-final-project/funding_dashboard.html"
DATE = "2026-09-22"

wb = openpyxl.load_workbook(XLSX)
ws = wb["Funding Dashboard"]

def esc(s): return html.escape(str(s if s is not None else ""))

def cur_sym(c):
    return {"USD":"$","GBP":"£","CHF":"CHF ","CAD":"C$","EUR":"€"}.get(c, (c+" ") if c else "")

def fmt_amt(v, c):
    if v is None or v == "": return ""
    try:
        n = float(v); s = cur_sym(c)
        if n >= 1_000_000 and n % 1_000_000 == 0: return f"{s}{int(n/1_000_000)}M"
        return f"{s}{n:,.0f}"
    except Exception:
        return esc(v)

rows = []
for r in range(2, ws.max_row + 1):
    name = ws.cell(r, 2).value
    if not name: continue
    rows.append(dict(
        wsn=ws.cell(r,1).value, name=name, funder=ws.cell(r,3).value or "",
        typ=ws.cell(r,4).value or "", geo=ws.cell(r,5).value or "",
        amt=fmt_amt(ws.cell(r,6).value, ws.cell(r,7).value),
        elig=ws.cell(r,8).value or "", fit=ws.cell(r,9).value,
        deadline=ws.cell(r,10).value or "", hours=ws.cell(r,11).value or "",
        status=(ws.cell(r,12).value or "").strip(), owner=ws.cell(r,13).value or "",
        action=ws.cell(r,14).value or "", link=ws.cell(r,15).value or "",
    ))

stat = Counter(x["status"] for x in rows)
fits = [x["fit"] for x in rows if isinstance(x["fit"], (int, float))]
avg_fit = round(sum(fits)/len(fits), 2) if fits else 0
total = len(rows)

STATUS_CLASS = {"Eligible":"ok","Won":"won","Researching":"research","Drafting":"draft",
    "Submitted":"submitted","Not eligible":"noteligible","Dead":"dead"}
WS_TITLE = {1:"US ecosystem grants &amp; accelerators", 2:"Angel investors &amp; pre-seed funds",
    3:"Philanthropic funders · Penda / WHO angle", "Triage":"Unidentified links — to resolve"}
WS_SUB = {1:"Equity-free where possible · Harvard, Boston, MA state, MIT, SBDC, NIH",
    2:"Dilutive by nature · pre-revenue / pre-seed check writers with health or AI theses",
    3:"Post-deployment monitoring of clinical AI in LMIC settings · safety &amp; equity framing",
    "Triage":"share.google links that could not be opened this session · paste the destination URL or grant name to verify"}

def fit_dots(f):
    if not isinstance(f,(int,float)):
        return '<span class="fit na" title="No fit score yet">—</span>'
    f = int(f)
    dots = "".join(f'<i class="{"on" if i<f else "off"}"></i>' for i in range(5))
    return f'<span class="fit" title="Fit {f} of 5" aria-label="Fit {f} of 5">{dots}</span>'

def deadline_flag(d):
    d = str(d)
    if "URGENT" in d or "21 Aug 2026" in d: return "urgent"
    if not d.strip(): return "cold"
    if "No open call" in d or "Unknown" in d or "invitation" in d.lower(): return "cold"
    if "Rolling" in d: return "rolling"
    return "dated"

# Application Calendar (hand-curated tab)
cal = wb["Application Calendar"]; cal_rows = []; r = 5
while cal.cell(r,3).value:
    cal_rows.append(dict(pri=cal.cell(r,1).value, deadline=cal.cell(r,2).value, opp=cal.cell(r,3).value,
        funder=cal.cell(r,4).value, ws=cal.cell(r,5).value, fit=cal.cell(r,6).value,
        hours=cal.cell(r,7).value, action=cal.cell(r,8).value)); r += 1

tiles = "".join(f'''
      <button type="button" class="tile filterable" data-filter="{esc(k)}" aria-pressed="false">
        <span class="tval">{stat.get(k,0)}</span>
        <span class="tlab"><i class="dot {STATUS_CLASS.get(k,'research')}"></i>{esc(k)}</span>
      </button>''' for k in ["Eligible","Researching","Not eligible","Dead"])

cal_cards = ""
for c in cal_rows:
    urgent = "urgent" if str(c["pri"]).startswith("1") else ""
    cal_cards += f'''
      <li class="calcard {urgent}">
        <div class="calhead"><span class="calpri">{esc(c["pri"])}</span><span class="caldl">{esc(c["deadline"])}</span></div>
        <h3>{esc(c["opp"])}</h3>
        <p class="calfunder">{esc(c["funder"])} · {fit_dots(c["fit"])} · <span class="calhrs">~{esc(c["hours"])}h</span></p>
        <p class="calaction">{esc(c["action"])}</p>
      </li>'''

def render_group(key):
    items = [x for x in rows if x["wsn"] == key]
    if not items: return ""
    order = {"Eligible":0,"Won":0,"Researching":1,"Drafting":1,"Submitted":1,"Not eligible":2,"Dead":3}
    items.sort(key=lambda x:(order.get(x["status"],1), -(x["fit"] if isinstance(x["fit"],(int,float)) else 0)))
    lis = ""
    for x in items:
        sc = STATUS_CLASS.get(x["status"],"research")
        link = f'<a class="offlink" href="{esc(x["link"])}" target="_blank" rel="noopener">Open link ↗</a>' if x["link"] else ""
        hrs = f'<span class="metaitem"><b>Est. effort</b> {esc(x["hours"])} h</span>' if x["hours"] else ""
        amt = f'<span class="amt">{esc(x["amt"])}</span>' if x["amt"] else '<span class="amt none">—</span>'
        lis += f'''
        <details class="row s-{sc}" data-status="{esc(x["status"])}">
          <summary>
            <span class="rfit">{fit_dots(x["fit"])}</span>
            <span class="rmain"><span class="rname">{esc(x["name"])}</span>
              <span class="rmeta">{esc(x["funder"])}{(" · " + esc(x["geo"])) if x["geo"] else ""}</span></span>
            {amt}
            <span class="rdl {deadline_flag(x["deadline"])}">{esc(x["deadline"]) or "—"}</span>
            <span class="pill {sc}">{esc(x["status"])}</span>
            <span class="chev" aria-hidden="true">›</span>
          </summary>
          <div class="rbody">
            <p class="raction"><b>Next action</b> {esc(x["action"])}</p>
            <p class="relig">{esc(x["elig"])}</p>
            <div class="rmetarow">{hrs}<span class="metaitem"><b>Owner</b> {esc(x["owner"])}</span>{link}</div>
          </div>
        </details>'''
    num = f"Workstream {key}" if key in (1,2,3) else "Needs triage"
    return f'''
    <section class="wsgroup">
      <div class="wshead"><span class="wsnum">{num}</span>
        <h2>{WS_TITLE[key]}</h2>
        <p class="wssub">{WS_SUB[key]} · <b>{len(items)}</b> {"items" if key=="Triage" else "opportunities"}</p></div>
      <div class="rows">{lis}</div>
    </section>'''

groups = "".join(render_group(k) for k in (1,2,3,"Triage"))

HTML = f'''<title>SafeAI Health — Non-Dilutive Funding Dashboard</title>
<style>
  :root {{
    --bg:#f5f7f6; --surface:#ffffff; --surface-2:#eef2f0; --raise:#f9fbfa;
    --ink:#15201e; --ink-soft:#41504c; --muted:#6c7a76; --faint:#93a09b;
    --border:#dbe3df; --border-2:#c8d2cd;
    --accent:#0f766e; --accent-ink:#0b5c55; --accent-tint:#e2f0ed;
    --ok:#15803d; --ok-tint:#dcf3e3; --research:#a15c07; --research-tint:#f7ecd6;
    --noteligible:#5b6b74; --noteligible-tint:#e7ecee; --dead:#b42318; --dead-tint:#f7e0dd;
    --submitted:#1d4ed8; --submitted-tint:#dfe6fb; --won:#0f766e; --draft:#1d4ed8;
    --urgent:#c2410c; --urgent-tint:#fbe4d5;
    --shadow:0 1px 2px rgba(20,40,35,.04),0 4px 16px rgba(20,40,35,.05);
    --serif:ui-serif,Georgia,"Times New Roman",serif;
    --sans:system-ui,-apple-system,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;
  }}
  @media (prefers-color-scheme:dark) {{ :root:not([data-theme="light"]) {{
      --bg:#0d1413; --surface:#131d1b; --surface-2:#1a2523; --raise:#172220;
      --ink:#e9efec; --ink-soft:#b3c0bb; --muted:#8b9994; --faint:#6d7c77;
      --border:#25322f; --border-2:#31413d;
      --accent:#2dd4bf; --accent-ink:#5ee0d0; --accent-tint:#123330;
      --ok:#4ade80; --ok-tint:#123626; --research:#e0a44a; --research-tint:#3a2c12;
      --noteligible:#9fb0b0; --noteligible-tint:#222e2e; --dead:#f4776a; --dead-tint:#3a1d1a;
      --submitted:#8aa6ff; --submitted-tint:#1a2340; --won:#2dd4bf; --draft:#8aa6ff;
      --urgent:#fb923c; --urgent-tint:#3a2413;
      --shadow:0 1px 2px rgba(0,0,0,.3),0 6px 20px rgba(0,0,0,.35);
  }} }}
  :root[data-theme="dark"] {{
      --bg:#0d1413; --surface:#131d1b; --surface-2:#1a2523; --raise:#172220;
      --ink:#e9efec; --ink-soft:#b3c0bb; --muted:#8b9994; --faint:#6d7c77;
      --border:#25322f; --border-2:#31413d;
      --accent:#2dd4bf; --accent-ink:#5ee0d0; --accent-tint:#123330;
      --ok:#4ade80; --ok-tint:#123626; --research:#e0a44a; --research-tint:#3a2c12;
      --noteligible:#9fb0b0; --noteligible-tint:#222e2e; --dead:#f4776a; --dead-tint:#3a1d1a;
      --submitted:#8aa6ff; --submitted-tint:#1a2340; --won:#2dd4bf; --draft:#8aa6ff;
      --urgent:#fb923c; --urgent-tint:#3a2413;
      --shadow:0 1px 2px rgba(0,0,0,.3),0 6px 20px rgba(0,0,0,.35);
  }}
  * {{ box-sizing:border-box; }}
  body {{ margin:0; background:var(--bg); color:var(--ink); font-family:var(--sans); line-height:1.5;
    -webkit-font-smoothing:antialiased; font-variant-numeric:tabular-nums; }}
  .wrap {{ max-width:1080px; margin:0 auto; padding:clamp(20px,4vw,44px) clamp(16px,4vw,32px) 72px; }}
  a {{ color:var(--accent-ink); }}
  h1,h2,h3 {{ font-family:var(--serif); text-wrap:balance; letter-spacing:-.01em; }}
  header.top {{ border-bottom:1px solid var(--border); padding-bottom:22px; margin-bottom:26px; }}
  .eyebrow {{ font-family:var(--sans); text-transform:uppercase; letter-spacing:.14em; font-size:.7rem; font-weight:700; color:var(--accent-ink); margin:0 0 8px; }}
  header.top h1 {{ margin:0; font-size:clamp(1.7rem,3.6vw,2.5rem); font-weight:600; }}
  header.top .lede {{ color:var(--ink-soft); margin:.5rem 0 0; max-width:64ch; }}
  .tiles {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(130px,1fr)); gap:12px; margin:22px 0 0; }}
  .tile {{ background:var(--surface); border:1px solid var(--border); border-radius:12px; padding:14px 16px; display:flex; flex-direction:column; gap:6px; box-shadow:var(--shadow); }}
  .tile.lead {{ background:var(--accent-tint); border-color:transparent; }}
  button.tile {{ font:inherit; text-align:left; color:inherit; }}
  button.tile.filterable {{ cursor:pointer; transition:transform .12s ease, border-color .12s ease, box-shadow .12s ease; }}
  button.tile.filterable:hover {{ border-color:var(--accent); transform:translateY(-1px); }}
  button.tile.filterable:focus-visible {{ outline:2px solid var(--accent); outline-offset:2px; }}
  button.tile.filterable[aria-pressed="true"] {{ border-color:var(--accent); box-shadow:inset 0 0 0 1px var(--accent); background:var(--accent-tint); }}
  .tval {{ font-family:var(--serif); font-size:1.9rem; font-weight:600; line-height:1; }}
  .tlab {{ font-size:.78rem; color:var(--muted); display:flex; align-items:center; gap:6px; }}
  .dot {{ width:9px; height:9px; border-radius:50%; display:inline-block; flex:none; }}
  .dot.ok{{background:var(--ok);}} .dot.research{{background:var(--research);}} .dot.noteligible{{background:var(--noteligible);}} .dot.dead{{background:var(--dead);}}
  .filterhint {{ margin:12px 2px 0; font-size:.78rem; color:var(--faint); }}
  .filterbar {{ display:flex; align-items:center; gap:14px; margin:14px 0 0; background:var(--accent-tint); border-radius:10px; padding:9px 14px; font-size:.85rem; color:var(--accent-ink); font-weight:600; }}
  .clearbtn {{ font:inherit; font-weight:700; cursor:pointer; margin-left:auto; color:var(--accent-ink); background:var(--surface); border:1px solid var(--border-2); border-radius:20px; padding:5px 13px; }}
  .clearbtn:hover {{ border-color:var(--accent); }}
  .clearbtn:focus-visible {{ outline:2px solid var(--accent); outline-offset:2px; }}
  .wsgroup[hidden], details.row[hidden] {{ display:none; }}
  .banner {{ display:flex; gap:12px; align-items:flex-start; margin:24px 0 6px; background:var(--surface-2); border:1px solid var(--border); border-left:3px solid var(--urgent); border-radius:10px; padding:13px 16px; font-size:.86rem; color:var(--ink-soft); }}
  .banner b {{ color:var(--ink); }}
  .banner .bi {{ color:var(--urgent); font-weight:800; }}
  .section-label {{ font-family:var(--sans); text-transform:uppercase; letter-spacing:.13em; font-size:.72rem; font-weight:700; color:var(--muted); margin:40px 0 14px; display:flex; align-items:center; gap:10px; }}
  .section-label::after {{ content:""; flex:1; height:1px; background:var(--border); }}
  ul.calendar {{ list-style:none; padding:0; margin:0; display:grid; grid-template-columns:repeat(auto-fill,minmax(250px,1fr)); gap:14px; }}
  .calcard {{ background:var(--surface); border:1px solid var(--border); border-radius:13px; padding:15px 16px 16px; box-shadow:var(--shadow); display:flex; flex-direction:column; gap:7px; }}
  .calcard.urgent {{ border-color:var(--urgent); background:linear-gradient(180deg,var(--urgent-tint),var(--surface) 60%); }}
  .calhead {{ display:flex; justify-content:space-between; align-items:center; gap:8px; }}
  .calpri {{ font-size:.66rem; font-weight:800; text-transform:uppercase; letter-spacing:.08em; color:var(--accent-ink); background:var(--accent-tint); padding:3px 8px; border-radius:20px; }}
  .calcard.urgent .calpri {{ color:#fff; background:var(--urgent); }}
  .caldl {{ font-size:.8rem; font-weight:700; color:var(--ink); }}
  .calcard.urgent .caldl {{ color:var(--urgent); }}
  .calcard h3 {{ margin:2px 0 0; font-size:1.02rem; font-weight:600; line-height:1.25; }}
  .calfunder {{ margin:0; font-size:.8rem; color:var(--muted); display:flex; flex-wrap:wrap; align-items:center; gap:6px; }}
  .calaction {{ margin:2px 0 0; font-size:.82rem; color:var(--ink-soft); }}
  .calhrs {{ color:var(--faint); }}
  .fit {{ display:inline-flex; gap:2.5px; align-items:center; }}
  .fit i {{ width:7px; height:7px; border-radius:2px; background:var(--border-2); display:block; }}
  .fit i.on {{ background:var(--accent); }}
  .fit.na {{ color:var(--faint); }}
  .wshead {{ margin:34px 0 12px; }}
  .wsnum {{ font-size:.7rem; text-transform:uppercase; letter-spacing:.13em; font-weight:700; color:var(--accent-ink); }}
  .wshead h2 {{ margin:3px 0 2px; font-size:1.4rem; font-weight:600; }}
  .wssub {{ margin:0; color:var(--muted); font-size:.85rem; }}
  .rows {{ display:flex; flex-direction:column; gap:7px; }}
  details.row {{ background:var(--surface); border:1px solid var(--border); border-radius:11px; box-shadow:var(--shadow); overflow:hidden; }}
  details.row[open] {{ border-color:var(--border-2); }}
  .row summary {{ list-style:none; cursor:pointer; display:grid; grid-template-columns:44px minmax(0,1fr) auto auto auto 18px; align-items:center; gap:14px; padding:12px 15px; }}
  .row summary::-webkit-details-marker {{ display:none; }}
  .row summary:hover {{ background:var(--raise); }}
  .row:focus-within {{ outline:2px solid var(--accent); outline-offset:1px; }}
  .rmain {{ min-width:0; display:flex; flex-direction:column; gap:2px; }}
  .rname {{ font-weight:600; font-size:.95rem; color:var(--ink); }}
  .rmeta {{ font-size:.78rem; color:var(--muted); white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
  .amt {{ font-weight:600; font-size:.9rem; color:var(--ink); white-space:nowrap; }}
  .amt.none {{ color:var(--faint); font-weight:400; }}
  .rdl {{ font-size:.78rem; color:var(--ink-soft); white-space:nowrap; max-width:190px; overflow:hidden; text-overflow:ellipsis; text-align:right; }}
  .rdl.urgent {{ color:var(--urgent); font-weight:700; }}
  .rdl.rolling {{ color:var(--accent-ink); }}
  .rdl.cold {{ color:var(--faint); }}
  .pill {{ font-size:.72rem; font-weight:700; padding:4px 10px; border-radius:20px; white-space:nowrap; text-align:center; }}
  .pill.ok{{color:var(--ok);background:var(--ok-tint);}} .pill.research{{color:var(--research);background:var(--research-tint);}}
  .pill.noteligible{{color:var(--noteligible);background:var(--noteligible-tint);}} .pill.dead{{color:var(--dead);background:var(--dead-tint);}}
  .pill.submitted{{color:var(--submitted);background:var(--submitted-tint);}} .pill.draft{{color:var(--draft);background:var(--submitted-tint);}} .pill.won{{color:#fff;background:var(--accent);}}
  .chev {{ color:var(--faint); transition:transform .18s ease; font-size:1.1rem; text-align:center; }}
  details.row[open] .chev {{ transform:rotate(90deg); }}
  .rbody {{ padding:2px 15px 15px 73px; border-top:1px solid var(--border); margin-top:2px; display:flex; flex-direction:column; gap:8px; }}
  .raction {{ margin:10px 0 0; font-size:.88rem; color:var(--ink); }}
  .raction b, .metaitem b {{ font-size:.7rem; text-transform:uppercase; letter-spacing:.06em; color:var(--muted); margin-right:6px; font-weight:700; }}
  .relig {{ margin:0; font-size:.83rem; color:var(--ink-soft); }}
  .rmetarow {{ display:flex; flex-wrap:wrap; gap:8px 20px; align-items:center; padding-top:2px; }}
  .metaitem {{ font-size:.82rem; color:var(--ink-soft); }}
  .offlink {{ font-size:.82rem; font-weight:600; }}
  .s-dead summary .rname, .s-noteligible summary .rname {{ color:var(--ink-soft); }}
  .s-dead {{ opacity:.82; }}
  footer {{ margin-top:44px; padding-top:18px; border-top:1px solid var(--border); color:var(--muted); font-size:.8rem; }}
  footer b {{ color:var(--ink-soft); }}
  @media (max-width:680px) {{
    .row summary {{ grid-template-columns:34px minmax(0,1fr) auto 16px; row-gap:8px; }}
    .rdl {{ grid-column:2/4; text-align:left; max-width:none; }}
    .pill {{ grid-column:1/3; justify-self:start; }}
    .amt {{ grid-column:3/4; }}
    .rbody {{ padding-left:15px; }}
  }}
  @media (prefers-reduced-motion:reduce) {{ * {{ transition:none!important; }} }}
</style>

<div class="wrap">
  <header class="top">
    <p class="eyebrow">SafeAI Health · Operating dashboard</p>
    <h1>Non-Dilutive Funding Pipeline</h1>
    <p class="lede">Grants, pre-seed investors, and philanthropic funders for a pre-incorporation,
      pre-revenue clinical-AI governance &amp; post-deployment monitoring venture. A read-only view of
      <b>Funding_Dashboard.xlsx</b> (the single source of truth), generated {DATE}.</p>
    <div class="tiles">
      <button type="button" class="tile lead filterable" data-filter="all" aria-pressed="false">
        <span class="tval">{total}</span><span class="tlab">All opportunities</span>
      </button>
      {tiles}
      <div class="tile static"><span class="tval">{avg_fit}</span><span class="tlab">Average fit (of 5)</span></div>
    </div>
    <p class="filterhint">Tap a status to filter the pipeline below.</p>
    <div id="filterbar" class="filterbar" hidden>
      <span id="filterlabel"></span>
      <button type="button" id="clearfilter" class="clearbtn">Show all ×</button>
    </div>
  </header>

  <div class="banner">
    <span class="bi">!</span>
    <span><b>Verification status.</b> This session's network policy blocked direct reads of official
      pages, so amounts, deadlines, and rules are <b>search-verified against official-domain content</b> —
      not live-page confirmations. The 8 rows under “Unidentified links” are share.google shorteners that
      could not be opened at all; paste each destination URL or grant name to verify. Confirm the
      <b>21 Aug 2026</b> Coefficient Giving deadline on its official page first.</span>
  </div>

  <p class="section-label">Application calendar — top 5 by deadline &amp; fit</p>
  <ul class="calendar">{cal_cards}
  </ul>

  <p class="section-label">Full pipeline</p>
  {groups}

  <footer>
    <p>Status legend: <b>Eligible</b> = eligible &amp; actionable now · <b>Researching</b> = eligible but
    watching for an open call or confirming fit · <b>Not eligible</b> = structural mismatch at current
    stage · <b>Dead</b> = disqualified. Owner: Lara. Fit score 1–5. Investors (Workstream 2) are dilutive
    and tagged Type = Investor so they never mix into grant reporting.</p>
    <p>Generated from Funding_Dashboard.xlsx · {DATE}. Update the workbook, not this page.</p>
  </footer>
</div>

<script>
(function () {{
  var tiles  = Array.prototype.slice.call(document.querySelectorAll('.tile.filterable'));
  var rows   = Array.prototype.slice.call(document.querySelectorAll('details.row'));
  var groups = Array.prototype.slice.call(document.querySelectorAll('.wsgroup'));
  var bar    = document.getElementById('filterbar');
  var label  = document.getElementById('filterlabel');
  var clear  = document.getElementById('clearfilter');
  var active = null;
  function apply() {{
    var shown = 0;
    rows.forEach(function (r) {{
      var match = (!active || active === 'all' || r.getAttribute('data-status') === active);
      r.hidden = !match; if (match) shown++; else r.open = false;
    }});
    groups.forEach(function (g) {{
      var any = Array.prototype.slice.call(g.querySelectorAll('details.row')).some(function (r) {{ return !r.hidden; }});
      g.hidden = !any;
    }});
    tiles.forEach(function (t) {{ t.setAttribute('aria-pressed', (active === t.getAttribute('data-filter')) ? 'true' : 'false'); }});
    if (active && active !== 'all') {{ bar.hidden = false; label.textContent = 'Showing ' + shown + ' \\u00b7 ' + active; }}
    else {{ bar.hidden = true; }}
  }}
  tiles.forEach(function (t) {{ t.addEventListener('click', function () {{
    var f = t.getAttribute('data-filter'); active = (active === f || f === 'all') ? null : f; apply();
  }}); }});
  clear.addEventListener('click', function () {{ active = null; apply(); }});
}})();
</script>'''

with open(OUT, "w") as f:
    f.write(HTML)
print("Wrote", OUT, "|", total, "rows | statuses:", dict(stat))
