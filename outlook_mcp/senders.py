"""Sender-level aggregation: turn per-message rows into per-sender rows.

Unsubscribing, muting and sweeping are all *per sender*, but every Graph
listing is per message. Grouping here (rather than in the model's head) is
what lets a 300-message folder be presented as ~50 decidable lines.
"""

# Trailing domain labels that carry no brand meaning. Not a real public-suffix
# list — enough to strip the tail off the domains bulk senders actually use.
_TLD_LABELS = frozenset(
    """com co uk net org io us de fr eu me info biz ie ca au nz nl es it se no dk
    fi pl ch at be in jp cn kr br mx za app shop store online site tech cloud dev
    ai xyz live one pro""".split()
)

# Labels that identify a mail stream or ESP rather than a brand.
_GENERIC_LABELS = frozenset(
    """email emails mail mails mailer mailing mailings e em m t news newsletter
    newsletters info send sends sending sendgrid mkt marketing campaign campaigns
    notification notifications notify noreply no-reply reply replies click clicks
    link links cust message messages messaging go cta ct hello contact smtp mta mg
    sg sparkpostmail exct exacttarget pages pg bounce bounces list lists updates
    update alerts alert digest engage engagement comms crm edm mandrillapp
    mailgun sendibm1 sendibm2 sendibt sendibt1 rsgsv mcsv tnsv cmail
    createsend""".split()
)

# Stripped off the surviving label to recover a glued-on brand
# ("emaildebenhams" -> "debenhams"). Deliberately tiny: a longer list starts
# eating real names, e.g. a bare "e" turns edfenergy into dfenergy and a "mail"
# suffix turns royalmail into royal.
_AFFIX_PREFIXES = ("emails", "email", "news")
_AFFIX_SUFFIXES = ("emails", "email")

_MIN_STEM = 4


def domain_of(address: str) -> str:
    """The full sending domain, subdomain included. Never collapsed — two
    subdomains of one brand can mean two completely different kinds of mail
    (billing vs marketing), so they stay distinct everywhere."""
    return (address or "").rpartition("@")[2].strip().lower()


def _strip_affix(label: str) -> str:
    for p in _AFFIX_PREFIXES:
        if label.startswith(p) and len(label) - len(p) >= _MIN_STEM:
            return label[len(p):]
    for s in _AFFIX_SUFFIXES:
        if label.endswith(s) and len(label) - len(s) >= _MIN_STEM:
            return label[: -len(s)]
    return label


def brand_for(address: str) -> str:
    """Heuristic human label for a sending address — "Debenhams" out of
    `news@emaildebenhams.com`. A display aid only: two addresses can produce
    the same brand, so it must never be used as a grouping or matching key.
    """
    domain = domain_of(address)
    if not domain:
        return (address or "").strip() or "unknown"
    labels = [l for l in domain.split(".") if l]
    while len(labels) > 1 and labels[-1] in _TLD_LABELS:
        labels.pop()
    candidates = [l for l in labels if l not in _GENERIC_LABELS] or labels
    stem = _strip_affix(max(candidates, key=len))
    return stem.replace("-", " ").strip().title() or domain


def group_senders(rows: list[dict], sender_key: str = "from") -> list[dict]:
    """Aggregate message rows into one row per *exact* sending address.

    Each row carries count, read/unread split and read_pct — measured
    engagement, which is the only honest basis for deciding what a mailbox
    owner actually wants. `brand_shared` marks addresses whose brand label
    collides with another address in the same result, so a
    billing-vs-marketing subdomain pair can never be silently merged.
    """
    by_sender: dict[str, dict] = {}
    for r in rows:
        addr = (r.get(sender_key) or "").strip().lower()
        if not addr:
            continue
        g = by_sender.get(addr)
        if g is None:
            g = by_sender[addr] = {
                "sender": addr,
                "brand": brand_for(addr),
                "domain": domain_of(addr),
                "count": 0,
                "read": 0,
                "unread": 0,
                "sample_ref": r.get("ref"),
                "newest": None,
                "oldest": None,
            }
        g["count"] += 1
        g["unread" if r.get("unread") else "read"] += 1
        date = r.get("date")
        if date:
            if g["newest"] is None or date > g["newest"]:
                g["newest"] = date
            if g["oldest"] is None or date < g["oldest"]:
                g["oldest"] = date

    brand_domains: dict[str, set[str]] = {}
    for g in by_sender.values():
        brand_domains.setdefault(g["brand"], set()).add(g["domain"])

    out = []
    for g in by_sender.values():
        g["read_pct"] = round(100 * g["read"] / g["count"]) if g["count"] else 0
        if len(brand_domains[g["brand"]]) > 1:
            g["brand_shared"] = True
        out.append(g)
    out.sort(key=lambda g: (-g["count"], g["sender"]))
    return out
