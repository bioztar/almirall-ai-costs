"""Inline SVG charts for the dashboard (no JS, prints cleanly)."""
import html


def bars(rows, fmt, color="#0b5cad"):
    """rows: [(label, value)] -> horizontal bar chart."""
    if not rows:
        return "<p class=muted>No data.</p>"
    top = max(v for _, v in rows) or 1
    h, out = 26, []
    for i, (label, v) in enumerate(rows):
        w = 520 * v / top
        out.append(f'<text x="0" y="{i*h+17}" class="lbl">{html.escape(str(label))}</text>'
                   f'<rect x="150" y="{i*h+5}" width="{w:.1f}" height="16" rx="3" fill="{color}"/>'
                   f'<text x="{156+w:.1f}" y="{i*h+17}" class="val">{fmt(v)}</text>')
    return f'<svg viewBox="0 0 800 {len(rows)*h+6}" role="img">{"".join(out)}</svg>'


def columns(rows, color="#0b5cad"):
    """rows: [(label, value)] in time order -> vertical column chart, labels on every 7th column."""
    if not rows:
        return "<p class=muted>No data.</p>"
    top = max(v for _, v in rows) or 1
    n, H = len(rows), 160
    w = 780 / n
    out = []
    for i, (label, v) in enumerate(rows):
        h = H * v / top
        out.append(f'<rect x="{10+i*w:.1f}" y="{10+H-h:.1f}" width="{w*0.75:.1f}" height="{h:.1f}" fill="{color}">'
                   f'<title>{html.escape(str(label))}: {v:,.0f}</title></rect>')
        if i % 7 == 0 or i == n - 1:
            out.append(f'<text x="{10+i*w:.1f}" y="{H+28}" class="val">{html.escape(str(label))}</text>')
    out.append(f'<text x="10" y="8" class="val">max {top:,.0f}</text>')
    return f'<svg viewBox="0 0 800 {H+34}" role="img">{"".join(out)}</svg>'
