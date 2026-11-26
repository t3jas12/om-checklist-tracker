import re

_FAULT_EXACT = frozenset({
    "non operational",
    "not operational",
    "not ok",
    "not clean",
    "non working",
    "not working",
    "faulty",
    "offline",
    "failed",
    "fail",
    "abnormal",
    "damaged",
    "dirty",
    "dirt found",
    "no",
})

_SPARE_EXACT = frozenset({
    "spare",
    "standby",
    "n/a",
    "na",
    "pending",
    "partially cleaned",
})

_OK_EXACT = frozenset({
    "operational",
    "ok",
    "working",
    "healthy",
    "clean",
    "no dirt",
    "done",
    "pass",
    "yes",
})


def normalize_option(value):
    """Normalize option strings (e.g. 'Non-Operational' -> 'non operational')."""
    return re.sub(r"[\s_-]+", " ", (value or "").strip().lower()).strip()


def classify_choice_kind(value):
    """Classify a choice label into 'fault', 'spare', 'ok', or 'other'."""
    norm = normalize_option(value)
    if not norm:
        return "other"
    if norm in _FAULT_EXACT or norm.startswith(("non ", "not ")):
        return "fault"
    if norm in _SPARE_EXACT:
        return "spare"
    if norm in _OK_EXACT:
        return "ok"
    return "other"


def short_item_label(item):
    """Return a concise item label (e.g. 'PV3 - Status Check' -> 'PV3')."""
    desc = (item.description or "").strip()
    if " - " in desc and desc.lower().endswith(" check"):
        return desc.split(" - ", 1)[0].strip()
    return desc or f"#{item.checkpoint_no}"


def supports_choice_summary(template):
    """Return True only for Inverter checklists (Operational / Non-Operational / Spare summaries)."""
    if not template:
        return False
    resolver = getattr(template, "resolved_category", None)
    return resolver() == "Inverters" if callable(resolver) else False


def summarize_submission_choices(submission, answers=None):
    """
    Compute dynamic per-option counts for Inverter checklists with dropdown
    'choices' items (Operational / Non-Operational / Spare).

    Returns None if the template is not an Inverter checklist or has no
    'choices' items; otherwise returns a dict:
      {
        "total_items": int,
        "option_counts": [{"label": str, "count": int, "kind": str, "items": list[str]}, ...],
        "non_operational_count": int,
        "operational_count": int,
        "spare_count": int,
        "non_operational_labels": list[str],
      }
    """
    if not supports_choice_summary(getattr(submission, "template", None)):
        return None
    if answers is None:
        answers = list(submission.answers.select_related("item").order_by("item__position", "item__checkpoint_no"))

    choice_answers = [a for a in answers if a.item.answer_type == "choices"]
    if not choice_answers:
        return None

    # Preserve option ordering defined on the template items, then append any unexpected submitted values
    ordered_labels = []
    seen_norm = {}
    for ans in choice_answers:
        for opt in ans.item.get_choices_list():
            norm = normalize_option(opt)
            if norm and norm not in seen_norm:
                seen_norm[norm] = opt
                ordered_labels.append(opt)
        if ans.value:
            norm = normalize_option(ans.value)
            if norm and norm not in seen_norm:
                seen_norm[norm] = ans.value.strip()
                ordered_labels.append(ans.value.strip())

    counts_by_norm = {norm: {"label": label, "count": 0, "kind": classify_choice_kind(label), "items": []} for norm, label in seen_norm.items()}

    for ans in choice_answers:
        norm = normalize_option(ans.value)
        if not norm:
            continue
        bucket = counts_by_norm.get(norm)
        if bucket is None:
            label = ans.value.strip()
            bucket = {"label": label, "count": 0, "kind": classify_choice_kind(label), "items": []}
            counts_by_norm[norm] = bucket
            ordered_labels.append(label)
        bucket["count"] += 1
        bucket["items"].append(short_item_label(ans.item))

    option_counts = [counts_by_norm[normalize_option(lbl)] for lbl in ordered_labels]
    non_op_count = sum(b["count"] for b in option_counts if b["kind"] == "fault")
    op_count = sum(b["count"] for b in option_counts if b["kind"] == "ok")
    spare_count = sum(b["count"] for b in option_counts if b["kind"] == "spare")
    non_op_labels = []
    for b in option_counts:
        if b["kind"] == "fault":
            non_op_labels.extend(b["items"])

    return {
        "total_items": len(choice_answers),
        "option_counts": option_counts,
        "non_operational_count": non_op_count,
        "operational_count": op_count,
        "spare_count": spare_count,
        "non_operational_labels": non_op_labels,
    }


def build_executive_summary(submissions, sites=None):
    """
    Build a per-site, day-by-day Executive Summary highlighting every
    non-operational / faulty asset across all submitted checklists.

    Returns a list of site summary dicts ordered by site name:
      [
        {
          "site": Site,
          "total_non_operational": int,
          "days_with_issues": int,
          "submitted_days": int,
          "daily_rows": [
            {
              "date": date,
              "submissions_count": int,
              "non_operational_count": int,
              "by_checklist": [
                {
                  "submission": Submission,
                  "template_name": str,
                  "count": int,
                  "assets": [{"label": str, "status": str, "section": str, "comment": str}, ...],
                },
              ],
            },
          ],
        },
      ]
    """
    site_map = {}
    ordered_sites = []
    if sites is not None:
        for s in sites:
            site_map[s.id] = s
            ordered_sites.append(s)

    subs_by_site = {}
    for sub in submissions:
        if sub.site_id not in site_map:
            site_map[sub.site_id] = sub.site
            ordered_sites.append(sub.site)
        subs_by_site.setdefault(sub.site_id, []).append(sub)

    result = []
    for site in ordered_sites:
        site_subs = subs_by_site.get(site.id, [])
        by_day = {}
        for sub in site_subs:
            by_day.setdefault(sub.checklist_date, []).append(sub)

        daily_rows = []
        total_non_op = 0
        days_with_issues = 0

        for day in sorted(by_day.keys(), reverse=True):
            day_subs = sorted(by_day[day], key=lambda s: s.template.name.lower())
            day_non_op = 0
            by_checklist = []

            for sub in day_subs:
                answers = sorted(sub.answers.all(), key=lambda a: (a.item.position, a.item.checkpoint_no))
                assets = []
                for ans in answers:
                    if classify_choice_kind(ans.value) == "fault":
                        assets.append({
                            "label": short_item_label(ans.item),
                            "status": ans.value.strip(),
                            "section": ans.item.section or "",
                            "comment": (ans.comment or "").strip(),
                        })
                if assets:
                    day_non_op += len(assets)
                    by_checklist.append({
                        "submission": sub,
                        "template_name": sub.template.name,
                        "count": len(assets),
                        "assets": assets,
                    })

            total_non_op += day_non_op
            if day_non_op > 0:
                days_with_issues += 1

            daily_rows.append({
                "date": day,
                "submissions_count": len(day_subs),
                "non_operational_count": day_non_op,
                "by_checklist": by_checklist,
            })

        result.append({
            "site": site,
            "total_non_operational": total_non_op,
            "days_with_issues": days_with_issues,
            "submitted_days": len(daily_rows),
            "daily_rows": daily_rows,
        })

    return result

