"""Placeholder-filled narrative text for a zone report, from ZoneStats."""

# Cat/domain names are Title-cased for display; acronyms preserved.
_ACRONYMS = {"NCI", "RAU", "GHQ", "BEES"}

# Connector phrase per domain rank (clamped to the last for deeper ranks).
_PHRASES = ["heavily concentrated in", "primarily in", "led by", "distributed across"]


def _disp(name):
    return name if name.upper() in _ACRONYMS else name.title()


def format_date_ordinal(d):
    suffix = "th" if 11 <= d.day % 100 <= 13 else {1: "st", 2: "nd", 3: "rd"}.get(d.day % 10, "th")
    return f"{d.day}{suffix} {d.strftime('%B %Y')}"


def build_title(zone_name, date_str):
    return f"{zone_name} Risk Analysis (Till {date_str})"


def build_narrative(stats, date_str):
    lines = []
    total = stats.total
    lines.append(f"A total of {total} risks identified across business units")
    if total == 0:
        return lines

    lines.append("Risk concentration is uneven, with strong clustering in a few key areas")
    top2 = stats.grand_by_cat and sorted(stats.grand_by_cat, key=lambda kv: -kv[1])[:2]
    if len(top2) == 2:
        (c1, n1), (c2, n2) = top2
        pct = round(100 * (n1 + n2) / total)
        lines.append(f"{_disp(c1)} ({n1} risks) and {_disp(c2)} ({n2} risks) "
                     f"together account for {pct}% of total exposure")

    lines.append("")
    lines.append("Key Insights by Risk Domain")
    for rank, (dom, dcount) in enumerate(stats.by_domain):
        superl = " – highest" if rank == 0 else ""
        lines.append(f"{_disp(dom)} ({stats.domain_pct[dom]}% of total risk{superl})")
        phrase = _PHRASES[min(rank, len(_PHRASES) - 1)]
        lines.append(f"{dcount} risks, {phrase}:")
        for cat, n in stats.top_cats[dom]:
            lines.append(f"    {_disp(cat)} ({n})")
        lines.append("")
    return lines
