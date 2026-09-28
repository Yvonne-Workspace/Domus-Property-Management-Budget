"""Full-screen owner story. Same feel as the Thornhill meeting deck, filled from the live budget."""

from __future__ import annotations

import base64
import html
from pathlib import Path


def _e(s) -> str:
    return html.escape(str(s or ""), quote=True)


def _r(n: float) -> str:
    return f"R{abs(float(n or 0)):,.0f}"


def _short(s: str, n: int = 22) -> str:
    s = str(s or "")
    return s if len(s) <= n else s[: n - 1] + "…"


def render_meeting_html(pack: dict, logo_path: Path | None = None) -> bytes:
    logo = ""
    if logo_path and Path(logo_path).exists():
        b64 = base64.b64encode(Path(logo_path).read_bytes()).decode()
        logo = f'<img class="logo" alt="Domus" src="data:image/jpeg;base64,{b64}">'
    slides = "\n".join(_slides(pack, logo))
    name = _e(pack.get("name") or "Budget")
    year = _e(pack.get("year") or "")
    page = _PAGE.replace("<!--TITLE-->", f"{name} — Budget")
    page = page.replace("<!--SLIDES-->", slides)
    page = page.replace("<!--FOOT-->", _e(f"{pack.get('name') or 'Budget'} · {pack.get('year') or ''}"))
    return page.encode("utf-8")


def _slides(p: dict, logo: str) -> list[str]:
    out = [_cover(p, logo), _plain(p), _headlines(p), _split(p), _bill(p)]
    why = _why(p)
    if why:
        out.append(why)
    run = _running(p)
    if run:
        out.append(run)
    if p.get("muni"):
        out.append(_muni(p))
    if p.get("repairs"):
        out.append(_repairs(p))
    out.append(_reserve(p))
    if p.get("jobs"):
        out.append(_jobs(p))
    if p.get("years"):
        out.append(_ten(p))
    out.append(_ask(p))
    out.append(_end(p, logo))
    return out


def _cover(p, logo):
    who = _e(p["who"])
    return f"""
<section class="slide dark">
  <div class="blob" style="width:420px;height:420px;background:var(--green);top:-140px;right:-90px;"></div>
  <div class="blob" style="width:200px;height:200px;background:#1a4736;bottom:-70px;right:120px;"></div>
  <div style="position:relative;padding-top:36px;">
    <div class="r" style="--d:40">{logo}</div>
    <div class="kicker r" style="--d:120;margin-top:18px;">{_e(p['name'])}</div>
    <h1 class="r" style="--d:200;font-size:58px;">Your budget, in plain words</h1>
    <div class="r" style="--d:280;font-family:var(--head);font-size:24px;color:var(--moss);margin-top:10px;">{_e(p['year'])}</div>
    <p class="r" style="--d:360;color:#bfd8c6;margin-top:8px;font-size:16px;">So you can see what your money pays for. Nothing is hidden.</p>
    <div style="display:flex;gap:16px;margin-top:36px;">
      <div class="card rp" style="--d:480;width:230px;"><div class="num" style="font-size:34px;" data-count="{int(p['units'])}">0</div><div class="mut" style="margin-top:4px;">owners sharing the cost</div></div>
      <div class="card rp" style="--d:580;width:250px;"><div class="num" style="font-size:34px;" data-count="{int(round(p['owner_month']))}" data-prefix="R">R0</div><div class="mut" style="margin-top:4px;">a month for {who}</div></div>
      <div class="card rp" style="--d:680;width:250px;"><div class="num" style="font-size:34px;" data-count="{int(round(p['collect_year']))}" data-prefix="R">R0</div><div class="mut" style="margin-top:4px;">collected for the whole year</div></div>
    </div>
  </div>
</section>"""


def _plain(p):
    return f"""
<section class="slide">
  <div class="kicker r">Start here</div>
  <h1 class="r" style="--d:60">Three simple questions</h1>
  <p class="sub r" style="--d:120">Every number in this budget answers one of these. If you remember only this slide, that is enough.</p>
  <div class="grid" style="grid-template-columns:repeat(3,1fr);margin-top:22px;">
    <div class="card r" style="--d:220;height:210px;">
      <h3>Money in</h3>
      <div style="font-size:11.5px;font-weight:700;color:var(--amber);letter-spacing:.6px;margin:8px 0;">WHAT OWNERS PAY</div>
      <p>A levy is the money owners pay each month so the complex can be looked after.</p>
    </div>
    <div class="card r" style="--d:320;height:210px;">
      <h3>Money out</h3>
      <div style="font-size:11.5px;font-weight:700;color:var(--amber);letter-spacing:.6px;margin:8px 0;">WHAT IT IS SPENT ON</div>
      <p>Security, gardens, insurance, repairs, and the people who pay the bills. Real costs. Not extras.</p>
    </div>
    <div class="card r" style="--d:420;height:210px;">
      <h3>Money saved</h3>
      <div style="font-size:11.5px;font-weight:700;color:var(--amber);letter-spacing:.6px;margin:8px 0;">THE RESERVE</div>
      <p>A little is saved every month. Then a big job does not become a sudden bill.</p>
    </div>
  </div>
  <div class="card pale row r" style="--d:540;margin-top:18px;">
    <p><b style="color:var(--green)">Please remember:</b> {_e(p['calm'])}</p>
  </div>
</section>"""


def _headlines(p):
    change = p.get("pct")
    if change is None:
        ch = '<div class="num" style="font-size:28px;">—</div><div style="font-weight:600;font-size:13px;margin:8px 0 6px;">compared with last year</div><div class="mut">We will show this once last year’s levy is filled in</div>'
    else:
        sign = "+" if change >= 0 else "−"
        word = "higher" if change >= 0.5 else ("lower" if change <= -0.5 else "about the same")
        ch = f'<div class="num" style="font-size:33px;color:var(--amber);" data-count="{abs(change):.1f}" data-prefix="{sign}" data-suffix="%" data-dec="1">{sign}0%</div><div style="font-weight:600;font-size:13px;margin:8px 0 6px;">{word} than last year</div><div class="mut">The next slides show why, in plain words</div>'
    return f"""
<section class="slide">
  <div class="kicker r">At a glance</div>
  <h1 class="r" style="--d:60">The only numbers that matter</h1>
  <div class="grid" style="grid-template-columns:repeat(4,1fr);margin-top:22px;">
    <div class="card rp" style="--d:180;min-height:168px;"><div class="num" style="font-size:30px;" data-count="{int(round(p['ordinary']))}" data-prefix="R">R0</div><div style="font-weight:600;font-size:13px;margin:8px 0 6px;">to run the complex</div><div class="mut">The ordinary levy for the year</div></div>
    <div class="card rp" style="--d:260;min-height:168px;"><div class="num" style="font-size:30px;" data-count="{int(round(p['ordinary']/12))}" data-prefix="R">R0</div><div style="font-weight:600;font-size:13px;margin:8px 0 6px;">each month, all owners</div><div class="mut">What must reach the bank</div></div>
    <div class="card rp" style="--d:340;min-height:168px;"><div class="num" style="font-size:30px;" data-count="{int(round(p['owner_month']))}" data-prefix="R">R0</div><div style="font-weight:600;font-size:13px;margin:8px 0 6px;">for {_e(p['who'])}</div><div class="mut">Levy, reserve and CSOS together</div></div>
    <div class="card rp" style="--d:420;min-height:168px;">{ch}</div>
  </div>
  <div class="card r" style="--d:540;margin-top:18px;">
    <h3>How we worked out the levy</h3>
    <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:18px;margin-top:16px;">
      <div class="row rx" style="--d:640"><div class="badge" style="width:34px;height:34px;background:var(--green);color:#fff;font-family:var(--head);font-weight:700;">1</div><p>We listed what the complex must pay this year.</p></div>
      <div class="row rx" style="--d:740"><div class="badge" style="width:34px;height:34px;background:var(--green);color:#fff;font-family:var(--head);font-weight:700;">2</div><p>We took off money that owners pay back, such as their own electricity.</p></div>
      <div class="row rx" style="--d:840"><div class="badge" style="width:34px;height:34px;background:var(--green);color:#fff;font-family:var(--head);font-weight:700;">3</div><p>What is left is the levy. Each owner pays their share.</p></div>
    </div>
  </div>
</section>"""


def _split(p):
    costs = p.get("costs") or []
    if not costs:
        return """<section class="slide"><h1>Where the levy goes</h1><p class="sub">Fill in the costs first. This slide will then show where the money goes.</p></section>"""
    colors = ["#1b5e43", "#2e7d5b", "#4e9a72", "#7fb685", "#c9751b", "#a7c4a9"]
    circ = 439.82
    rings = []
    off = 0.0
    legend = []
    total = sum(c["amount"] for c in costs) or 1
    for i, c in enumerate(costs):
        length = circ * (c["amount"] / total)
        col = colors[i % len(colors)]
        rings.append(
            f'<circle class="dseg" cx="100" cy="100" r="70" fill="none" stroke="{col}" stroke-width="34" data-len="{length:.2f}" data-off="{-off:.2f}"></circle>'
        )
        off += length
        legend.append(
            f'<div class="card row rx" style="--d:{280+i*70};padding:10px 14px;">'
            f'<span style="width:11px;height:11px;border-radius:3px;background:{col};flex:0 0 auto;"></span>'
            f'<div style="flex:1"><b style="color:var(--green);font-size:14px;">{_e(c["label"])}</b>'
            f'<div class="mut">{_e(c["blurb"])}</div></div>'
            f'<span class="money" style="font-size:16px;">{_r(c["amount"])}</span></div>'
        )
    return f"""
<section class="slide">
  <div class="kicker r">The split</div>
  <h1 class="r" style="--d:60">Where every rand of the levy goes</h1>
  <p class="sub r" style="--d:100">This is the ordinary levy only. The reserve is extra, and it is saved, not spent on these bills.</p>
  <div style="display:flex;gap:26px;margin-top:8px;align-items:center;">
    <div class="rp" style="--d:180;flex:0 0 auto;">
      <svg width="300" height="300" viewBox="0 0 200 200">
        <g transform="rotate(-90 100 100)">{''.join(rings)}</g>
        <text x="100" y="96" text-anchor="middle" style="font-family:var(--head);font-size:18px;font-weight:800;fill:#1b5e43;">{_r(p['ordinary'])}</text>
        <text x="100" y="114" text-anchor="middle" style="font-family:var(--body);font-size:8px;fill:#5f6f66;">LEVY FOR THE YEAR</text>
      </svg>
    </div>
    <div style="flex:1;display:flex;flex-direction:column;gap:8px;">{''.join(legend)}</div>
  </div>
</section>"""


def _bill(p):
    rows = []
    for i, line in enumerate(p.get("invoice") or []):
        rows.append(
            f'<tr style="--d:{220+i*70}"><td><b>{_e(line["name"])}</b></td>'
            f'<td class="money">{_r(line["monthly"])}</td><td>{_e(line["plain"])}</td></tr>'
        )
    rows.append(
        f'<tr class="tot" style="--d:700"><td>TOTAL</td><td class="money" style="font-size:17px;">{_r(p["owner_month"])}</td>'
        f'<td style="font-style:italic;color:#c7dbcd;">For {_e(p["who"])}, each month</td></tr>'
    )
    return f"""
<section class="slide">
  <div class="kicker r">What you actually pay</div>
  <h1 class="r" style="--d:60">Your monthly bill, line by line</h1>
  <p class="sub r" style="--d:100">Each line has a job. You can see them separately, so nothing is buried in one number.</p>
  <table style="margin-top:16px;">
    <thead><tr><th style="width:240px;">On your statement</th><th style="width:160px;">A month</th><th>What it pays for</th></tr></thead>
    <tbody>{''.join(rows)}</tbody>
  </table>
  <div class="card pale row r" style="--d:780;margin-top:14px;"><p>{_e(p['bill_note'])}</p></div>
</section>"""


def _why(p):
    risers = p.get("risers") or []
    if not risers:
        return ""
    top = max(r["more"] for r in risers) or 1
    bars = []
    cards = []
    colors = ["#c9751b", "#1b5e43", "#2e7d5b", "#4e9a72", "#7fb685"]
    for i, r in enumerate(risers[:5]):
        w = max(6, r["more"] / top * 100)
        bars.append(
            f'<div class="track r" style="--d:{220+i*80}"><span class="tlabel">{_e(_short(r["label"]))}</span>'
            f'<div class="bar" style="--w:{w:.1f}%;background:{colors[i%5]};"></div>'
            f'<span class="tval">+{_r(r["more"])}</span></div>'
        )
        if i < 3:
            cards.append(
                f'<div class="card r" style="--d:{300+i*90}"><h3>{_e(r["label"])} &nbsp;+{_r(r["more"])}</h3>'
                f'<p style="margin-top:8px;">{_e(r["plain"])}</p></div>'
            )
    title = "Why the levy is higher" if (p.get("pct") or 0) >= 0.5 else "What changed from last year"
    return f"""
<section class="slide">
  <div class="kicker r">The honest answer</div>
  <h1 class="r" style="--d:60">{title}</h1>
  <p class="sub r" style="--d:100">{_e(p.get('why_lead') or 'These are the lines that cost more than last year. The rest barely moved.')}</p>
  <div style="display:flex;gap:22px;margin-top:16px;">
    <div style="flex:1;display:flex;flex-direction:column;gap:12px;">{''.join(bars)}
      <div class="card pale r" style="--d:700;padding:12px 16px;"><p style="font-style:italic;">This is not money for luxuries. It is the real cost of looking after the place you live.</p></div>
    </div>
    <div style="flex:0 0 420px;display:flex;flex-direction:column;gap:10px;">{''.join(cards)}</div>
  </div>
</section>"""


def _running(p):
    lines = p.get("big") or []
    if not lines:
        return ""
    top = max(x["yearly"] for x in lines) or 1
    colors = ["#1b5e43", "#2e7d5b", "#4e9a72", "#7fb685", "#a7c4a9"]
    bars = []
    for i, r in enumerate(lines[:6]):
        w = max(4, r["yearly"] / top * 100)
        bars.append(
            f'<div class="track r" style="--d:{200+i*70}"><span class="tlabel">{_e(_short(r["label"]))}</span>'
            f'<div class="bar" style="--w:{w:.1f}%;background:{colors[i%5]};"></div>'
            f'<span class="tval">{_r(r["yearly"])}</span></div>'
        )
    return f"""
<section class="slide">
  <div class="kicker r">Money out</div>
  <h1 class="r" style="--d:60">The cost of keeping it running</h1>
  <p class="sub r" style="--d:100">These are the bigger bills. Each one has a job. None of them is a hidden extra.</p>
  <div style="display:flex;flex-direction:column;gap:14px;margin-top:26px;max-width:980px;">{''.join(bars)}</div>
  <p class="mut r" style="--d:680;margin-top:16px;font-style:italic;">Last year’s real cost is where we started. We then allowed a fair increase, or we typed the quote we already have.</p>
</section>"""


def _muni(p):
    m = p["muni"]
    gaps = m.get("gaps") or []
    cards = []
    for i, g in enumerate(gaps[:4]):
        gap = g["gap"]
        if gap > 50:
            word, col = f"Shared part {_r(gap)}", "var(--amber)"
            note = "Owners pay back less than the bill. The rest stays in the levy."
        elif gap < -50:
            word, col = f"Over {_r(gap)}", "var(--green)"
            note = "Owners pay back more than this bill. That helps the levy."
        else:
            word, col = "It balances", "var(--green)"
            note = "What we pay and what owners pay back are almost the same."
        cards.append(
            f'<div class="card rx" style="--d:{360+i*80}"><div class="row" style="justify-content:space-between;">'
            f'<b style="color:var(--green);">{_e(g["name"])}</b><span class="money" style="color:{col};">{word}</span></div>'
            f'<div class="mut" style="margin-top:4px;">Bill {_r(g["gross"])} · paid back {_r(g["rec"])}. {note}</div></div>'
        )
    net = m["net"]
    if net > 50:
        lead = "Some of the city bill is for the common property. That shared part is inside the levy. You are not charged for it twice."
    else:
        lead = "Owners pay back what they use. That money is taken off the levy. You are not charged for it twice."
    return f"""
<section class="slide">
  <div class="kicker r">The city account</div>
  <h1 class="r" style="--d:60">You do not pay the city bill twice</h1>
  <p class="sub r" style="--d:100">{lead}</p>
  <div style="display:flex;gap:18px;margin-top:18px;">
    <div class="card rp" style="--d:200;flex:1;"><div class="mut">We pay the city</div><div class="num" style="font-size:32px;margin-top:6px;" data-count="{int(round(m['gross']))}" data-prefix="R">R0</div></div>
    <div class="card rp" style="--d:280;flex:1;"><div class="mut">Owners pay back</div><div class="num" style="font-size:32px;margin-top:6px;" data-count="{int(round(m['rec']))}" data-prefix="R">R0</div></div>
    <div class="card amber rp" style="--d:360;flex:1;"><div class="mut">Left in the levy</div><div class="num" style="font-size:32px;margin-top:6px;" data-count="{int(round(max(net,0)))}" data-prefix="R">R0</div></div>
  </div>
  <div style="display:flex;flex-direction:column;gap:8px;margin-top:14px;">{''.join(cards)}</div>
</section>"""


def _repairs(p):
    cards = []
    for i, r in enumerate(p["repairs"][:4]):
        cards.append(
            f'<div class="card rp" style="--d:{200+i*80};min-height:180px;">'
            f'<div class="num" style="font-size:28px;" data-count="{int(round(r["yearly"]))}" data-prefix="R">R0</div>'
            f'<div style="font-weight:600;font-size:14px;margin:8px 0;">{_e(r["desc"])}</div>'
            f'<div class="mut">{_e(r["plain"])}</div></div>'
        )
    return f"""
<section class="slide">
  <div class="kicker r">When something breaks</div>
  <h1 class="r" style="--d:60">Small repairs, already allowed for</h1>
  <p class="sub r" style="--d:100">This money is for ordinary breakages. Big planned jobs come from the reserve, not from a surprise collection.</p>
  <div class="grid" style="grid-template-columns:repeat({min(4,len(cards))},1fr);margin-top:20px;">{''.join(cards)}</div>
</section>"""


def _reserve(p):
    return f"""
<section class="slide dark">
  <div class="kicker r">Saving instead of surprising</div>
  <h1 class="r" style="--d:60">The reserve is a savings jar</h1>
  <p class="sub r" style="--d:100">It is not spent on the monthly bills. Those are in the levy. This jar is for the big jobs.</p>
  <div class="row" style="margin-top:28px;gap:12px;align-items:stretch;">
    <div class="card rp" style="--d:220;flex:1;"><div class="mut">Already in the bank</div><div class="num" style="font-size:30px;margin-top:8px;" data-count="{int(round(p['opening']))}" data-prefix="R">R0</div><div class="mut" style="margin-top:6px;">Interest already earned is inside this amount.</div></div>
    <div class="r" style="--d:300;font-family:var(--head);font-size:28px;color:var(--moss);">+</div>
    <div class="card rp" style="--d:340;flex:1;"><div class="mut">Put in this year</div><div class="num" style="font-size:30px;margin-top:8px;" data-count="{int(round(p['reserve']))}" data-prefix="R">R0</div><div class="mut" style="margin-top:6px;">From the owners, a little each month.</div></div>
    <div class="r" style="--d:400;font-family:var(--head);font-size:28px;color:var(--moss);">−</div>
    <div class="card rp" style="--d:440;flex:1;"><div class="mut">Used on this year’s jobs</div><div class="num" style="font-size:30px;margin-top:8px;" data-count="{int(round(p['projects']))}" data-prefix="R">R0</div><div class="mut" style="margin-top:6px;">Only the work planned for this year.</div></div>
  </div>
  <div class="card r" style="--d:560;margin-top:18px;background:var(--green);"><div style="font-family:var(--head);font-size:26px;font-weight:800;">Should still be in the jar: {_r(p['projected'])}</div></div>
</section>"""


def _jobs(p):
    cards = []
    for i, j in enumerate(p["jobs"][:4]):
        cards.append(
            f'<div class="card rp" style="--d:{200+i*80};min-height:150px;">'
            f'<div class="num" style="font-size:26px;" data-count="{int(round(j["yearly"]))}" data-prefix="R">R0</div>'
            f'<div style="font-weight:600;margin-top:8px;">{_e(j["desc"])}</div></div>'
        )
    return f"""
<section class="slide">
  <div class="kicker r">This year’s big jobs</div>
  <h1 class="r" style="--d:60">Paid from the savings, not from a new levy</h1>
  <div class="grid" style="grid-template-columns:repeat({min(4,len(cards))},1fr);margin-top:22px;">{''.join(cards)}</div>
  <div class="card solid row r" style="--d:620;margin-top:18px;justify-content:space-between;">
    <div style="font-family:var(--head);font-size:22px;font-weight:800;">Together: {_r(p['projects'])}</div>
    <div style="font-size:14px;">Already taken off the reserve. Not added on top of the levy.</div>
  </div>
</section>"""


def _ten(p):
    years = p["years"]
    mx = max((y["amount"] for y in years), default=1) or 1
    cols = []
    labels = []
    for y in years:
        h = 4 if y["amount"] < 1 else max(8, y["amount"] / mx * 210)
        hot = y["amount"] >= mx * 0.98 and y["amount"] > 1
        col = "#c9751b" if hot else "#1b5e43"
        cols.append(
            f'<div style="flex:1;text-align:center;"><div class="tval" style="font-size:10px;margin-bottom:4px;">{_r(y["amount"]) if y["amount"]>1 else "R0"}</div>'
            f'<div class="col" style="--h:{h:.0f}px;background:{col};"></div></div>'
        )
        labels.append(f'<div style="flex:1;text-align:center;" class="mut">{_e(y["label"])}</div>')
    return f"""
<section class="slide">
  <div class="kicker r">Looking ahead</div>
  <h1 class="r" style="--d:60">The next ten years</h1>
  <p class="sub r" style="--d:80">This is the plan, so a big year is not a surprise. The tall bar is the busy year.</p>
  <div style="margin-top:18px;">
    <div style="display:flex;align-items:end;gap:8px;height:280px;border-bottom:2px solid #dce5dd;">{''.join(cols)}</div>
    <div style="display:flex;gap:8px;padding-top:8px;">{''.join(labels)}</div>
  </div>
  <p class="mut r" style="--d:500;margin-top:12px;font-style:italic;">These jobs are paid from the reserve, not from the ordinary levy.</p>
</section>"""


def _ask(p):
    items = [
        ("1", "Approve the budget", f"Ordinary levies of {_r(p['ordinary'])} for the year."),
        ("2", f"Approve what {_e(p['who'])} pays", f"{_r(p['owner_month'])} a month, levy plus reserve plus CSOS."),
        ("3", "Approve the reserve", f"{_r(p['reserve'])} put into the savings jar this year."),
    ]
    if p.get("projects", 0) > 1:
        items.append(("4", "Note the big jobs", f"{_r(p['projects'])} of planned work, paid from the reserve."))
    blocks = []
    for i, (n, title, body) in enumerate(items):
        blocks.append(
            f'<div class="card row rx" style="--d:{180+i*90}"><div class="badge" style="background:var(--moss);color:var(--forest);font-family:var(--head);font-weight:800;">{n}</div>'
            f'<div><h3 style="font-size:18px;">{title}</h3><div class="mut" style="margin-top:3px;">{body}</div></div></div>'
        )
    return f"""
<section class="slide dark">
  <div class="kicker r">Please approve</div>
  <h1 class="r" style="--d:60">What we are asking of you</h1>
  <div style="display:flex;flex-direction:column;gap:10px;margin-top:20px;">{''.join(blocks)}</div>
  <p class="mut r" style="--d:640;margin-top:14px;font-style:italic;">We prepared this from last year’s real costs, so you can say yes with your eyes open.</p>
</section>"""


def _end(p, logo):
    return f"""
<section class="slide dark">
  <div class="blob" style="width:430px;height:430px;background:var(--green);top:-130px;right:-100px;"></div>
  <div style="position:relative;padding-top:150px;">
    <div class="r" style="--d:80">{logo}</div>
    <h1 class="r" style="--d:180;font-size:64px;margin-top:18px;">Any questions?</h1>
    <p class="r" style="--d:320;font-size:18px;color:#c7dbcd;margin-top:16px;max-width:720px;">We are happy to walk through any line. The full budget is available to every owner.</p>
    <p class="r" style="--d:420;margin-top:18px;color:var(--moss);font-size:14px;">Domus · prepared with care for the owners of {_e(p['name'])}</p>
  </div>
</section>"""


_PAGE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title><!--TITLE--></title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@700;800&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
:root{
  --forest:#12342a; --forest2:#194536; --green:#1b5e43; --green2:#2e7d5b;
  --moss:#7fb685; --moss2:#a7c4a9; --pale:#e4efe6; --light:#f4f7f3;
  --amber:#c9751b; --amber2:#e8a24a; --ink:#16281f; --muted:#5f6f66;
  --head:'Playfair Display', Georgia, serif;
  --body:'Inter', Calibri, 'Segoe UI', Arial, sans-serif;
}
*{box-sizing:border-box; margin:0; padding:0;}
html,body{height:100%;}
body{background:#0c231c; font-family:var(--body); color:var(--ink); overflow:hidden; display:flex; align-items:center; justify-content:center;}
#viewport{position:fixed; inset:0; display:flex; align-items:center; justify-content:center;}
#stage{width:1280px; height:720px; position:relative; transform-origin:center center; border-radius:14px; overflow:hidden; box-shadow:0 30px 90px rgba(0,0,0,.55);}
.slide{position:absolute; inset:0; padding:48px 52px 58px; opacity:0; visibility:hidden; transform:translateX(42px) scale(.985); transition:opacity .5s ease, transform .55s cubic-bezier(.22,.9,.3,1), visibility .55s; background:var(--light);}
.slide.dark{background:var(--forest);}
.slide.active{opacity:1; visibility:visible; transform:none;}
.slide.prev{transform:translateX(-42px) scale(.985);}
.r{opacity:0; transform:translateY(16px); transition:opacity .55s ease, transform .6s cubic-bezier(.22,.9,.3,1);}
.slide.active .r{opacity:1; transform:none; transition-delay:calc(var(--d,0) * 1ms);}
.rx{opacity:0; transform:translateX(-18px);}
.slide.active .rx{opacity:1; transform:none;}
.rp{opacity:0; transform:scale(.9);}
.slide.active .rp{opacity:1; transform:none;}
.kicker{font-size:12px; font-weight:700; letter-spacing:2.6px; text-transform:uppercase; color:var(--amber); margin-bottom:8px;}
.dark .kicker{color:var(--moss);}
h1{font-family:var(--head); font-size:40px; font-weight:800; color:var(--green); line-height:1.08;}
.dark h1{color:#fff;}
.sub{font-size:15.5px; color:var(--muted); margin-top:8px; max-width:1080px; line-height:1.45;}
.dark .sub{color:#c7dbcd;}
h3{font-family:var(--head); font-size:18px; color:var(--green); font-weight:700;}
.dark h3{color:#fff;}
p{font-size:14px; line-height:1.5; color:var(--ink);}
.dark p{color:#d6e6da;}
.mut{color:var(--muted); font-size:12.5px; line-height:1.45;}
.dark .mut{color:#aec9b6;}
.num{font-family:var(--head); font-weight:800; color:var(--green); letter-spacing:-.6px;}
.dark .num{color:#fff;}
.card{background:#fff; border-radius:11px; padding:16px 18px; box-shadow:0 3px 14px rgba(27,94,67,.11);}
.dark .card{background:var(--forest2); box-shadow:none;}
.card.pale{background:var(--pale); box-shadow:none;}
.card.amber{background:var(--amber);} .card.amber *{color:#fff !important;} .card.amber .mut{color:#fdf0e0 !important;}
.card.solid{background:var(--green);} .card.solid *{color:#fff !important;}
.badge{width:36px; height:36px; border-radius:50%; background:var(--pale); display:flex; align-items:center; justify-content:center; flex:0 0 auto;}
.grid{display:grid; gap:14px;}
.row{display:flex; align-items:center; gap:12px;}
.bar{height:24px; border-radius:5px; width:0; transition:width 1.1s cubic-bezier(.22,.9,.3,1);}
.slide.active .bar{width:var(--w);}
.col{width:100%; border-radius:5px 5px 0 0; height:0; align-self:end; transition:height 1.1s cubic-bezier(.22,.9,.3,1);}
.slide.active .col{height:var(--h);}
.track{display:flex; align-items:center; gap:9px;}
.tlabel{width:150px; text-align:right; font-size:12.5px; color:var(--muted); flex:0 0 auto;}
.tval{font-size:12px; font-weight:600; color:var(--ink); white-space:nowrap; width:92px;}
.dseg{transition:stroke-dasharray 1.3s cubic-bezier(.22,.9,.3,1); stroke-dasharray:0 440;}
table{width:100%; border-collapse:collapse; font-size:13.5px;}
thead th{background:var(--green); color:#fff; text-align:left; padding:8px 12px; font-weight:600;}
tbody td{padding:8px 12px; border-bottom:1px solid #dce5dd;}
tbody tr:nth-child(odd){background:#ebf1ea;} tbody tr:nth-child(even){background:#fff;}
tbody tr.tot td{background:var(--forest); color:#fff; font-weight:600;}
tbody tr{opacity:0; transform:translateX(-14px); transition:opacity .45s ease, transform .5s ease;}
.slide.active tbody tr{opacity:1; transform:none; transition-delay:calc(var(--d) * 1ms);}
.money{font-family:var(--head); font-weight:700;}
.blob{position:absolute; border-radius:50%; pointer-events:none;}
.logo{height:42px; background:#fff; border-radius:8px; padding:4px 8px;}
@keyframes drift{0%,100%{transform:translate(0,0);} 50%{transform:translate(-16px,18px);}}
.slide.active .blob{animation:drift 14s ease-in-out infinite;}
#bar{position:absolute; top:0; left:0; height:4px; background:var(--amber); width:0; transition:width .45s ease; z-index:6;}
#foot{position:absolute; bottom:0; left:0; right:0; height:36px; display:flex; align-items:center; justify-content:space-between; padding:0 52px; font-size:10.5px; letter-spacing:1.4px; text-transform:uppercase; color:var(--muted); z-index:5;}
.dark #foot{color:var(--moss);}
#dots{display:flex; gap:5px;}
.dot{width:7px; height:7px; border-radius:50%; background:#c3d3c6; cursor:pointer;}
.dot.on{background:var(--amber); width:18px; border-radius:4px;}
.dark .dot{background:#2f5c49;}
#ctl{position:fixed; bottom:16px; right:18px; display:flex; gap:8px; z-index:20;}
#ctl button{background:rgba(255,255,255,.1); color:#dbeade; border:1px solid rgba(255,255,255,.18); border-radius:8px; padding:7px 13px; font-family:var(--body); font-size:12px; cursor:pointer;}
#hint{position:fixed; bottom:20px; left:18px; color:#6d8c7c; font-size:11.5px; z-index:20;}
</style>
</head>
<body>
<div id="viewport"><div id="stage">
  <div id="bar"></div>
<!--SLIDES-->
  <div id="foot"><span><!--FOOT--></span><div id="dots"></div><span id="fnum">01</span></div>
</div></div>
<div id="hint">Arrow keys or space · F for full screen · A to play through</div>
<div id="ctl"><button id="prev">Prev</button><button id="auto">Play</button><button id="next">Next</button><button id="fs">Full screen</button></div>
<script>
(function(){
  var slides = Array.prototype.slice.call(document.querySelectorAll('.slide'));
  var i = 0, timer = null;
  var bar = document.getElementById('bar'), dots = document.getElementById('dots'), fnum = document.getElementById('fnum'), foot = document.getElementById('foot');
  slides.forEach(function(s, n){ var d = document.createElement('div'); d.className = 'dot'; d.addEventListener('click', function(){ go(n); }); dots.appendChild(d); });
  var dotEls = Array.prototype.slice.call(dots.children);
  function fit(){ var k = Math.min((window.innerWidth - 36) / 1280, (window.innerHeight - 36) / 720); document.getElementById('stage').style.transform = 'scale(' + k + ')'; }
  window.addEventListener('resize', fit); fit();
  function countUp(el){
    var target = parseFloat(el.dataset.count), dec = parseInt(el.dataset.dec || '0', 10), pre = el.dataset.prefix || '', suf = el.dataset.suffix || '', t0 = null;
    function step(t){ if (!t0) t0 = t; var p = Math.min((t - t0) / 1100, 1); var v = target * (1 - Math.pow(1 - p, 3));
      el.textContent = pre + v.toLocaleString('en-US', {minimumFractionDigits: dec, maximumFractionDigits: dec}) + suf;
      if (p < 1) requestAnimationFrame(step); }
    requestAnimationFrame(step);
  }
  function animate(s){
    s.querySelectorAll('[data-count]').forEach(function(el, n){ setTimeout(function(){ countUp(el); }, 280 + n * 80); });
    s.querySelectorAll('.dseg').forEach(function(c, n){
      c.style.strokeDasharray = '0 440'; c.style.strokeDashoffset = c.dataset.off;
      setTimeout(function(){ c.style.strokeDasharray = c.dataset.len + ' ' + (439.8 - parseFloat(c.dataset.len)); }, 250 + n * 120);
    });
  }
  function go(n){
    if (n < 0 || n >= slides.length) return;
    slides[i].classList.remove('active'); slides[i].classList.add('prev');
    var old = i; i = n;
    setTimeout(function(){ slides[old].classList.remove('prev'); }, 550);
    slides[i].classList.add('active'); slides[i].classList.remove('prev');
    animate(slides[i]);
    bar.style.width = ((i + 1) / slides.length * 100) + '%';
    fnum.textContent = ('0' + (i + 1)).slice(-2) + ' / ' + slides.length;
    dotEls.forEach(function(d, k){ d.classList.toggle('on', k === i); });
    foot.classList.toggle('dark', slides[i].classList.contains('dark'));
    document.body.style.background = slides[i].classList.contains('dark') ? '#0c231c' : '#0e2a21';
  }
  document.addEventListener('keydown', function(e){
    if (e.key === 'ArrowRight' || e.key === ' ' || e.key === 'PageDown'){ e.preventDefault(); go(i + 1); }
    if (e.key === 'ArrowLeft' || e.key === 'PageUp'){ e.preventDefault(); go(i - 1); }
    if (e.key === 'f' || e.key === 'F'){ if (!document.fullscreenElement) document.documentElement.requestFullscreen(); else document.exitFullscreen(); }
    if (e.key === 'a' || e.key === 'A') toggleAuto();
  });
  document.getElementById('next').onclick = function(){ go(i + 1); };
  document.getElementById('prev').onclick = function(){ go(i - 1); };
  document.getElementById('fs').onclick = function(){ if (!document.fullscreenElement) document.documentElement.requestFullscreen(); else document.exitFullscreen(); };
  var autoBtn = document.getElementById('auto');
  function toggleAuto(){
    if (timer){ clearInterval(timer); timer = null; autoBtn.textContent = 'Play'; }
    else { autoBtn.textContent = 'Pause'; timer = setInterval(function(){ if (i >= slides.length - 1){ clearInterval(timer); timer = null; autoBtn.textContent = 'Play'; return; } go(i + 1); }, 9000); }
  }
  autoBtn.onclick = toggleAuto;
  slides[0].classList.add('active'); animate(slides[0]);
  bar.style.width = (1 / slides.length * 100) + '%';
  dotEls[0].classList.add('on'); foot.classList.add('dark');
  fnum.textContent = '01 / ' + slides.length;
})();
</script>
</body>
</html>
"""
