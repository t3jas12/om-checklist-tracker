import calendar
from datetime import date, datetime, timedelta

from django.contrib import messages
from django.contrib.auth import get_user_model, update_session_auth_hash
from django.contrib.auth.decorators import login_required, user_passes_test
from django.core.exceptions import PermissionDenied, ValidationError
from django.db import IntegrityError, models, transaction
from django.http import HttpResponseForbidden
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone

from .forms import (
    AdminCredentialChangeForm,
    ChecklistUploadForm,
    DeadlineForm,
    EquipmentForm,
    SiteChecklistAssignForm,
    SiteForm,
    UserEditForm,
    UserProvisionForm,
)
from .models import (
    CATEGORY_CHOICES,
    CATEGORY_ORDER,
    Answer,
    ChecklistItem,
    ChecklistTemplate,
    GlobalSettings,
    Site,
    SiteChecklist,
    SiteEquipment,
    Submission,
    UserSite,
    resolve_equipment_label,
)
from .services.exports import excel_response, pdf_response, range_excel_response, range_pdf_response
from .services.summary import build_executive_summary, summarize_submission_choices
from .services.workbook_import import import_uploaded_csv, import_uploaded_excel, import_uploaded_pdf

User = get_user_model()


def is_admin(user):
    return user.is_authenticated and user.is_staff


def assigned_site(user):
    try:
        return user.site_assignment.site
    except (AttributeError, User.site_assignment.RelatedObjectDoesNotExist):
        return None


def _draft_session_key(user, site, checklist_date, template):
    return f"checklist_draft_{user.pk}_{site.pk}_{checklist_date.isoformat()}_{template.pk}"


def _collect_checklist_values(request, items):
    errors = []
    values = {}
    for item in items:
        value = request.POST.get(f"item_{item.id}", "").strip()
        comment = request.POST.get(f"comment_{item.id}", "").strip()
        if item.required and not value:
            errors.append(item.checkpoint_no)
        if len(value) > 300:
            errors.append(item.checkpoint_no)
        if item.answer_type == "number" and value:
            try:
                float(value)
            except ValueError:
                errors.append(item.checkpoint_no)
        if item.answer_type == "boolean" and value and value not in {"YES", "NO"}:
            errors.append(item.checkpoint_no)
        if item.answer_type == "choices" and value:
            allowed = {c.strip() for c in item.choices.split(",") if c.strip()}
            if allowed and value not in allowed:
                errors.append(item.checkpoint_no)
        values[str(item.id)] = {"value": value, "comment": comment}
    remarks = request.POST.get("remarks", "").strip()
    return values, remarks, sorted(set(errors))


def _preview_rows(items, values):
    return [{"item": item, **values.get(str(item.id), {"value": "", "comment": ""})} for item in items]


def _group_checklist_items(checklist_status_list):
    """Group site checklist options into ordered category sections."""
    groups = {}
    for item in checklist_status_list:
        cat = item.get("category") or item["template"].resolved_category()
        groups.setdefault(cat, []).append(item)

    ordered = []
    for cat in sorted(groups.keys(), key=lambda c: (CATEGORY_ORDER.get(c, 99), c.lower())):
        items = groups[cat]
        submitted_count = sum(1 for it in items if it.get("submission") is not None)
        available_count = sum(1 for it in items if it.get("available"))
        non_op_count = sum(
            it["choice_summary"]["non_operational_count"]
            for it in items
            if it.get("choice_summary")
        )
        ordered.append({
            "category": cat,
            "items": items,
            "total": len(items),
            "submitted_count": submitted_count,
            "available_count": available_count,
            "all_done": bool(items) and submitted_count == len(items),
            "non_op_count": non_op_count,
        })
    return ordered


def _site_category_map(site=None):
    qs = SiteChecklist.objects.select_related("template")
    if site is not None:
        qs = qs.filter(site=site)
    return {(sc.site_id, sc.template_id): sc.resolved_category() for sc in qs}


def _filter_submissions_by_category(submissions, category, site_cat_map):
    if not category:
        return submissions
    return [
        sub for sub in submissions
        if site_cat_map.get((sub.site_id, sub.template_id), sub.template.resolved_category()) == category
    ]


def _parse_filter_dates(request, default_start, default_end):
    month_start = default_end.replace(day=1)
    month_end = default_end.replace(day=calendar.monthrange(default_end.year, default_end.month)[1])
    if request.GET.get("all_days") == "1":
        earliest = Submission.objects.order_by("checklist_date").values_list("checklist_date", flat=True).first()
        start_date = earliest or month_start
        end_date = default_end
    elif request.GET.get("month", "").strip():
        try:
            m_dt = datetime.strptime(request.GET.get("month", "").strip(), "%Y-%m").date()
            start_date = m_dt.replace(day=1)
            end_date = m_dt.replace(day=calendar.monthrange(m_dt.year, m_dt.month)[1])
        except ValueError:
            start_date, end_date = month_start, month_end
    else:
        start_str = request.GET.get("start_date", "").strip()
        end_str = request.GET.get("end_date", "").strip()
        try:
            start_date = datetime.strptime(start_str, "%Y-%m-%d").date() if start_str else default_start
        except ValueError:
            start_date = default_start
        try:
            end_date = datetime.strptime(end_str, "%Y-%m-%d").date() if end_str else default_end
        except ValueError:
            end_date = default_end
    if start_date > end_date:
        start_date, end_date = end_date, start_date
    return start_date, end_date


@login_required
@user_passes_test(is_admin)
def executive_summary(request):
    today = timezone.localdate()
    month_start = today.replace(day=1)
    all_sites = list(Site.objects.filter(active=True).order_by("name"))
    site_id = request.GET.get("site", "").strip()
    start_date, end_date = _parse_filter_dates(request, month_start, today)

    qs = (
        Submission.objects.filter(checklist_date__range=[start_date, end_date])
        .select_related("site", "template", "submitted_by")
        .prefetch_related("answers__item")
        .order_by("-checklist_date", "site__name", "-submitted_at")
    )
    if site_id:
        qs = qs.filter(site_id=site_id)

    submissions = list(qs)
    target_sites = [s for s in all_sites if str(s.pk) == site_id] if site_id else all_sites
    executive_summaries = build_executive_summary(submissions, sites=target_sites)
    total_non_op = sum(item["total_non_operational"] for item in executive_summaries)
    sites_with_issues = sum(1 for item in executive_summaries if item["total_non_operational"] > 0)

    return render(request, "checklist/executive_summary.html", {
        "today": today,
        "all_sites": all_sites,
        "selected_site": site_id,
        "start_date": start_date,
        "end_date": end_date,
        "current_month_str": today.strftime("%Y-%m"),
        "executive_summaries": executive_summaries,
        "total_non_op": total_non_op,
        "sites_with_issues": sites_with_issues,
    })


@login_required
def dashboard(request):
    today = timezone.localdate()
    if is_admin(request.user):
        active_tab = request.GET.get("tab", "status")
        all_sites = Site.objects.filter(active=True).prefetch_related("checklists__template", "equipment", "user_assignment__user").order_by("name")

        # Today's status tab data
        today_submissions = list(
            Submission.objects.filter(checklist_date=today)
            .select_related("site", "template", "submitted_by")
            .prefetch_related("answers__item")
        )
        site_submissions = {}
        for s in today_submissions:
            s.choice_summary = summarize_submission_choices(s, list(s.answers.all()))
            site_submissions.setdefault(s.site_id, []).append(s)

        site_cat_map = _site_category_map()
        rows = []
        for site in all_sites:
            assigned_checklists = list(site.checklists.filter(active=True).select_related("template").order_by("position"))
            subs = site_submissions.get(site.id, [])
            submitted_template_ids = {s.template_id for s in subs}
            total_req = len(assigned_checklists) if assigned_checklists else 1
            completed_count = len(subs)
            is_complete = completed_count >= total_req if assigned_checklists else (completed_count > 0)
            choice_subs = [s for s in subs if s.choice_summary]
            non_op_today = sum(s.choice_summary["non_operational_count"] for s in choice_subs)

            group_counts = {}
            for sc in assigned_checklists:
                cat = sc.resolved_category()
                grp = group_counts.setdefault(cat, {"category": cat, "total": 0, "submitted": 0})
                grp["total"] += 1
                if sc.template_id in submitted_template_ids:
                    grp["submitted"] += 1
            group_progress = [
                group_counts[c]
                for c in sorted(group_counts.keys(), key=lambda k: (CATEGORY_ORDER.get(k, 99), k.lower()))
            ]

            rows.append({
                "site": site,
                "submissions": subs,
                "first_submission": subs[0] if subs else None,
                "assigned_checklists": assigned_checklists,
                "group_progress": group_progress,
                "completed_count": completed_count,
                "total_required": total_req,
                "is_complete": is_complete,
                "choice_subs": choice_subs,
                "non_op_today": non_op_today,
            })

        # Inline review tab data
        review_site_id = request.GET.get("site", "")
        review_template_id = request.GET.get("template", "")
        review_category = request.GET.get("category", "").strip()
        month_start = today.replace(day=1)
        month_end = today.replace(day=calendar.monthrange(today.year, today.month)[1])
        start_date, end_date = _parse_filter_dates(request, today, today)

        review_qs = Submission.objects.filter(checklist_date__range=[start_date, end_date]).select_related("site", "template", "submitted_by").prefetch_related("answers__item").order_by("-checklist_date", "site__name", "-submitted_at")
        if review_site_id:
            review_qs = review_qs.filter(site_id=review_site_id)
        if review_template_id:
            review_qs = review_qs.filter(template_id=review_template_id)

        review_submissions = _filter_submissions_by_category(list(review_qs), review_category, site_cat_map)
        string_summary_subs = []
        total_non_operational = 0
        for sub in review_submissions:
            sub.category = site_cat_map.get((sub.site_id, sub.template_id), sub.template.resolved_category())
            sub.choice_summary = summarize_submission_choices(sub, list(sub.answers.all()))
            if sub.choice_summary:
                string_summary_subs.append(sub)
                total_non_operational += sub.choice_summary["non_operational_count"]

        all_templates = ChecklistTemplate.objects.filter(active=True).order_by("name", "-version")
        all_categories = [c[0] for c in CATEGORY_CHOICES]

        return render(request, "checklist/admin_dashboard.html", {
            "rows": rows,
            "today": today,
            "deadline": GlobalSettings.get_solo().deadline,
            "active_tab": active_tab,
            "all_sites": all_sites,
            "all_templates": all_templates,
            "all_categories": all_categories,
            "review_site_id": review_site_id,
            "review_template_id": review_template_id,
            "review_category": review_category,
            "start_date": start_date,
            "end_date": end_date,
            "month_start": month_start,
            "month_end": month_end,
            "current_month_str": today.strftime("%Y-%m"),
            "review_submissions": review_submissions,
            "string_summary_subs": string_summary_subs,
            "total_non_operational": total_non_operational,
        })

    # Operator dashboard
    site = assigned_site(request.user)
    if site is None:
        return render(request, "checklist/no_site.html")
    checklist_status_list = _checklist_options_for_site(site, today)
    for item in checklist_status_list:
        sub = item.get("submission")
        item["choice_summary"] = summarize_submission_choices(sub) if sub else None
    grouped_checklists = _group_checklist_items(checklist_status_list)
    eligible_templates = [item["template"] for item in checklist_status_list if item["available"]]
    next_template = eligible_templates[0] if eligible_templates else None
    all_done = bool(checklist_status_list) and all(item["submission"] is not None for item in checklist_status_list)
    has_any = any(item["submission"] is not None for item in checklist_status_list)
    first_submission = next((item["submission"] for item in checklist_status_list if item["submission"]), None)

    return render(request, "checklist/user_dashboard.html", {
        "site": site,
        "submission": first_submission,
        "checklist_status_list": checklist_status_list,
        "grouped_checklists": grouped_checklists,
        "next_template": next_template,
        "eligible_count": len(eligible_templates),
        "all_done": all_done,
        "has_any": has_any,
        "today": today,
        "deadline": GlobalSettings.get_solo().deadline,
        "open": GlobalSettings.is_open(),
    })


def _checklist_options_for_site(site, today):
    """Return each active checklist with today's submission, category, and sequence eligibility."""
    assigned = list(
        site.checklists.filter(active=True, template__active=True)
        .select_related("template")
        .order_by("position", "template__name", "template_id")
    )
    assignments = [(sc.template, sc.position, sc.resolved_category()) for sc in assigned]
    if not assignments:
        fallback = ChecklistTemplate.objects.filter(active=True).order_by("-version").first()
        if fallback:
            assignments = [(fallback, None, fallback.resolved_category())]

    submissions = {
        submission.template_id: submission
        for submission in Submission.objects.filter(
            site=site,
            checklist_date=today,
            template_id__in=[template.pk for template, _, _ in assignments],
        )
    }
    pending_unsequenced = {
        template.pk for template, position, _ in assignments
        if position in (0, None) and template.pk not in submissions
    }
    first_pending_sequenced = next((
        template.pk for template, position, _ in assignments
        if position and template.pk not in submissions
    ), None)
    eligible_ids = pending_unsequenced
    if first_pending_sequenced is not None:
        eligible_ids.add(first_pending_sequenced)

    return [
        {
            "template": template,
            "position": position,
            "category": category,
            "submission": submissions.get(template.pk),
            "available": template.pk in eligible_ids,
        }
        for template, position, category in assignments
    ]


def _get_next_template_for_site(site, today):
    """Return the first eligible checklist, retaining the legacy default route behavior."""
    return next((item["template"] for item in _checklist_options_for_site(site, today) if item["available"]), None)


def _all_templates_for_site(site):
    """Return ordered list of templates assigned to site, or global fallback."""
    return [item["template"] for item in _checklist_options_for_site(site, timezone.localdate())]


def _selected_eligible_template(request, site, today):
    """Resolve a requested checklist, or default only when exactly one is eligible."""
    requested_id = request.GET.get("template_id") or request.POST.get("template_id")
    eligible = [item["template"] for item in _checklist_options_for_site(site, today) if item["available"]]
    if requested_id:
        try:
            requested_id = int(requested_id)
        except (TypeError, ValueError):
            return None, "That checklist selection is invalid. Choose an available checklist from the dashboard."
        template = next((item for item in eligible if item.pk == requested_id), None)
        if template is None:
            return None, "That checklist is not currently available. Choose an available checklist from the dashboard."
        return template, None
    if len(eligible) == 1:
        return eligible[0], None
    if len(eligible) > 1:
        return None, "Choose which available checklist you want to fill from the dashboard."
    return None, None


@login_required
def checklist_form(request):
    if is_admin(request.user):
        raise PermissionDenied
    site = assigned_site(request.user)
    if site is None or not site.active:
        return render(request, "checklist/no_site.html")
    today = timezone.localdate()
    if not GlobalSettings.is_open():
        messages.error(request, "The submission cutoff has passed. Today's checklist is locked.")
        return redirect("dashboard")

    template, selection_error = _selected_eligible_template(request, site, today)
    if selection_error:
        messages.error(request, selection_error)
        return redirect("dashboard")
    if template is None:
        if not _checklist_options_for_site(site, today):
            messages.info(request, "No active checklists are configured for this site.")
        else:
            messages.info(request, "All checklists for today have been submitted.")
        return redirect("dashboard")

    checklist_options = _checklist_options_for_site(site, today)
    all_templates = [item["template"] for item in checklist_options]
    done_ids = {item["template"].pk for item in checklist_options if item["submission"]}
    checklist_progress = [{"template": t, "done": t.pk in done_ids} for t in all_templates]
    current_index = next((i for i, t in enumerate(all_templates) if t.pk == template.pk), 0)

    items = list(template.items.prefetch_related().all())
    draft_key = _draft_session_key(request.user, site, today, template)
    values = request.session.get(draft_key, {})
    if not isinstance(values, dict):
        values = {}
    valid_item_ids = {str(item.id) for item in items}
    if set(values) - valid_item_ids - {"remarks"}:
        values = {}
        request.session.pop(draft_key, None)
    remarks = values.get("remarks", "")
    display_values = {key: record.get("value", "") for key, record in values.items() if key != "remarks" and isinstance(record, dict)}
    display_comments = {key: record.get("comment", "") for key, record in values.items() if key != "remarks" and isinstance(record, dict)}

    ctx = {
        "site": site, "template": template, "template_id": template.pk, "items": items, "today": today,
        "deadline": GlobalSettings.get_solo().deadline,
        "equipment_label": resolve_equipment_label(site, template),
        "checklist_progress": checklist_progress, "current_index": current_index,
        "total_checklists": len(all_templates),
    }
    if request.method == "POST":
        if not GlobalSettings.is_open():
            messages.error(request, "The submission cutoff has passed. Your checklist was not submitted.")
            return redirect("dashboard")
        values, remarks, missing = _collect_checklist_values(request, items)
        if missing:
            messages.error(request, "Please complete all required fields and enter valid numbers. Checkpoints: " + ", ".join(map(str, missing)))
            values["remarks"] = remarks
            display_values = {str(key): record["value"] for key, record in values.items() if key != "remarks"}
            display_comments = {str(key): record["comment"] for key, record in values.items() if key != "remarks"}
            return render(request, "checklist/checklist_form.html", {**ctx, "values": display_values, "comments": display_comments, "remarks": remarks})
        values["remarks"] = remarks
        request.session[draft_key] = values
        request.session.modified = True
        return render(request, "checklist/checklist_preview.html", {**ctx, "rows": _preview_rows(items, values), "remarks": remarks})
    return render(request, "checklist/checklist_form.html", {**ctx, "values": display_values, "comments": display_comments, "remarks": remarks})


@login_required
def finalize_checklist(request):
    if is_admin(request.user):
        raise PermissionDenied
    site = assigned_site(request.user)
    today = timezone.localdate()
    if site is None or not site.active:
        return redirect("dashboard")
    if not GlobalSettings.is_open():
        messages.error(request, "The submission cutoff has passed. Today's checklist is locked.")
        return redirect("dashboard")

    template, selection_error = _selected_eligible_template(request, site, today)
    if selection_error:
        messages.error(request, selection_error)
        return redirect("dashboard")
    if template is None:
        messages.info(request, "All checklists for today have been submitted or no checklist was selected.")
        return redirect("dashboard")

    draft_key = _draft_session_key(request.user, site, today, template)
    values = request.session.get(draft_key)
    if not isinstance(values, dict):
        messages.error(request, "Your review draft has expired. Please enter the checklist again.")
        return redirect("checklist-form")
    remarks = values.get("remarks", "")
    items = list(template.items.all())
    errors = []
    for item in items:
        record = values.get(str(item.id), {})
        value = record.get("value", "")
        if item.required and not value:
            errors.append(item.checkpoint_no)
        if len(value) > 300:
            errors.append(item.checkpoint_no)
        if item.answer_type == "number" and value:
            try:
                float(value)
            except ValueError:
                errors.append(item.checkpoint_no)
        if item.answer_type == "boolean" and value and value not in {"YES", "NO"}:
            errors.append(item.checkpoint_no)
        if item.answer_type == "choices" and value:
            allowed = {c.strip() for c in item.choices.split(",") if c.strip()}
            if allowed and value not in allowed:
                errors.append(item.checkpoint_no)
    if errors:
        messages.error(request, "The saved review is no longer valid. Please review the checklist again.")
        return redirect("checklist-form")

    checklist_options = _checklist_options_for_site(site, today)
    all_templates = [item["template"] for item in checklist_options]
    done_ids = {item["template"].pk for item in checklist_options if item["submission"]}
    checklist_progress = [{"template": t, "done": t.pk in done_ids} for t in all_templates]
    current_index = next((i for i, t in enumerate(all_templates) if t.pk == template.pk), 0)
    ctx = {
        "site": site, "template": template, "template_id": template.pk, "items": items, "today": today,
        "remarks": remarks, "deadline": GlobalSettings.get_solo().deadline,
        "equipment_label": resolve_equipment_label(site, template),
        "rows": _preview_rows(items, values),
        "checklist_progress": checklist_progress, "current_index": current_index,
        "total_checklists": len(all_templates),
    }

    if request.method == "GET":
        return render(request, "checklist/checklist_lockin.html", ctx)
    if request.method != "POST":
        return HttpResponseForbidden("Final confirmation must be submitted.")
    if request.POST.get("confirm_lockin") != "yes":
        messages.error(request, "Confirm that the reviewed checklist is accurate before locking it in.")
        return redirect("finalize-checklist")
    if not GlobalSettings.is_open():
        messages.error(request, "The submission cutoff has passed. Your checklist was not submitted.")
        return redirect("dashboard")
    try:
        with transaction.atomic():
            locked_site = Site.objects.select_for_update().get(pk=site.pk, active=True)
            if not GlobalSettings.is_open():
                messages.error(request, "The submission cutoff has passed. Your checklist was not submitted.")
                return redirect("dashboard")
            still_eligible, _ = _selected_eligible_template(request, locked_site, today)
            if still_eligible is None or still_eligible.pk != template.pk:
                messages.error(request, "This checklist is no longer available in the current sequence. Choose another checklist from the dashboard.")
                return redirect("dashboard")
            submission = Submission.objects.create(
                site=locked_site, template=template, submitted_by=request.user, checklist_date=today,
                remarks=remarks, signer_name=(request.user.get_full_name() or request.user.username),
                equipment_snapshot=resolve_equipment_label(locked_site, template),
            )
            Answer.objects.bulk_create([
                Answer(submission=submission, item=item, value=values[str(item.id)]["value"], comment=values[str(item.id)].get("comment", ""))
                for item in items
            ])
    except (IntegrityError, ValidationError):
        messages.error(request, "A checklist was already submitted for this site and template today, or this submission could not be validated.")
        return redirect("dashboard")
    del request.session[draft_key]
    remaining = [item["template"] for item in _checklist_options_for_site(site, today) if item["available"]]
    if len(remaining) == 1:
        messages.success(request, f"'{template.name}' submitted. '{remaining[0].name}' is ready to fill.")
        return redirect("checklist-form")
    if len(remaining) > 1:
        messages.success(request, f"'{template.name}' submitted. Choose another available checklist from the dashboard.")
        return redirect("dashboard")
    messages.success(request, "All checklists locked and submitted successfully.")
    return redirect("submission-review", submission_id=submission.pk)


@login_required
def submission_review(request, submission_id):
    submission = get_object_or_404(Submission.objects.select_related("site", "template", "submitted_by").prefetch_related("answers__item"), pk=submission_id)
    if not (request.user.is_staff or (submission.submitted_by_id == request.user.id and not request.user.is_staff)):
        raise PermissionDenied
    answers = list(submission.answers.select_related("item").order_by("item__position", "item__checkpoint_no"))
    choice_summary = summarize_submission_choices(submission, answers)
    return render(request, "checklist/submission_review.html", {
        "submission": submission,
        "answers": answers,
        "choice_summary": choice_summary,
        "equipment_label": resolve_equipment_label(submission.site, submission.template) or submission.equipment_snapshot,
    })


@login_required
def operator_excel_review(request, submission_id):
    submission = get_object_or_404(Submission.objects.select_related("site", "template", "submitted_by"), pk=submission_id)
    if not (request.user.is_staff or submission.submitted_by_id == request.user.id):
        raise PermissionDenied
    return excel_response(submission)


def _resolve_export_date_range(request):
    today = timezone.localdate()
    month_str = request.GET.get("month", "").strip()
    start_str = request.GET.get("start_date", "").strip()
    end_str = request.GET.get("end_date", "").strip()

    if month_str:
        try:
            m_dt = datetime.strptime(month_str, "%Y-%m").date()
            start_date = m_dt.replace(day=1)
            end_date = m_dt.replace(day=calendar.monthrange(m_dt.year, m_dt.month)[1])
            return start_date, end_date
        except ValueError:
            pass

    if start_str or end_str:
        try:
            start_date = datetime.strptime(start_str, "%Y-%m-%d").date() if start_str else today.replace(day=1)
        except ValueError:
            start_date = today.replace(day=1)
        try:
            end_date = datetime.strptime(end_str, "%Y-%m-%d").date() if end_str else today.replace(day=calendar.monthrange(today.year, today.month)[1])
        except ValueError:
            end_date = today.replace(day=calendar.monthrange(today.year, today.month)[1])
        if start_date > end_date:
            start_date, end_date = end_date, start_date
        return start_date, end_date

    start_date = today.replace(day=1)
    end_date = today.replace(day=calendar.monthrange(today.year, today.month)[1])
    return start_date, end_date


def _run_monthly_export(request, file_type, *, site=None, fallback_redirect="history"):
    if file_type not in {"xlsx", "pdf"}:
        return HttpResponseForbidden("Unsupported export format")

    start_date, end_date = _resolve_export_date_range(request)
    template_id = request.GET.get("template", "").strip()
    category = request.GET.get("category", "").strip()

    qs = (
        Submission.objects.filter(checklist_date__range=[start_date, end_date])
        .select_related("site", "template", "submitted_by")
        .prefetch_related("answers__item")
        .order_by("checklist_date", "site__name", "template__name")
    )
    if site is not None:
        qs = qs.filter(site=site)
    if template_id:
        qs = qs.filter(template_id=template_id)

    submissions = _filter_submissions_by_category(list(qs), category, _site_category_map(site))
    if not submissions:
        messages.warning(
            request,
            f"No submissions found for {site.name if site else 'the selected sites'} between {start_date:%d %b %Y} and {end_date:%d %b %Y}.",
        )
        return redirect(request.META.get("HTTP_REFERER") or fallback_redirect)

    if file_type == "xlsx":
        return range_excel_response(submissions, site, start_date, end_date)
    return range_pdf_response(submissions, site, start_date, end_date)


@login_required
def history(request):
    today = timezone.localdate()
    selected_month = request.GET.get("month", "").strip()
    selected_template = request.GET.get("template", "").strip()
    selected_site = request.GET.get("site", "").strip()
    selected_category = request.GET.get("category", "").strip()

    if is_admin(request.user):
        submissions = Submission.objects.select_related("site", "template", "submitted_by").prefetch_related("answers__item").order_by("-checklist_date", "site__name")
        if selected_site:
            submissions = submissions.filter(site_id=selected_site)
        sites = Site.objects.order_by("name")
        templates = ChecklistTemplate.objects.order_by("name", "-version")
        operator_site = None
    else:
        operator_site = assigned_site(request.user)
        submissions = Submission.objects.filter(submitted_by=request.user).select_related("site", "template").prefetch_related("answers__item").order_by("-checklist_date")
        sites = []
        templates = (
            ChecklistTemplate.objects.filter(site_checklists__site=operator_site).distinct().order_by("name", "-version")
            if operator_site
            else ChecklistTemplate.objects.none()
        )

    if selected_template:
        submissions = submissions.filter(template_id=selected_template)
    if selected_month:
        try:
            m_dt = datetime.strptime(selected_month, "%Y-%m").date()
            m_start = m_dt.replace(day=1)
            m_end = m_dt.replace(day=calendar.monthrange(m_dt.year, m_dt.month)[1])
            submissions = submissions.filter(checklist_date__range=[m_start, m_end])
        except ValueError:
            selected_month = ""

    site_cat_map = _site_category_map(operator_site)
    submissions_list = _filter_submissions_by_category(list(submissions), selected_category, site_cat_map)
    for sub in submissions_list:
        sub.category = site_cat_map.get((sub.site_id, sub.template_id), sub.template.resolved_category())
        sub.choice_summary = summarize_submission_choices(sub, list(sub.answers.all()))

    return render(request, "checklist/history.html", {
        "submissions": submissions_list,
        "sites": sites,
        "templates": templates,
        "all_categories": [c[0] for c in CATEGORY_CHOICES],
        "operator_site": operator_site,
        "selected_site": selected_site,
        "selected_template": selected_template,
        "selected_category": selected_category,
        "selected_month": selected_month,
        "export_month": selected_month or today.strftime("%Y-%m"),
    })


@login_required
@user_passes_test(is_admin)
def export_submission(request, submission_id, file_type):
    submission = get_object_or_404(Submission.objects.select_related("site", "template", "submitted_by"), pk=submission_id)
    if file_type == "xlsx":
        return excel_response(submission)
    if file_type == "pdf":
        return pdf_response(submission)
    return HttpResponseForbidden("Unsupported export format")


@login_required
@user_passes_test(is_admin)
def export_monthly(request, file_type):
    site_id = request.GET.get("site", "").strip()
    site = get_object_or_404(Site, pk=site_id) if site_id else None
    return _run_monthly_export(request, file_type, site=site, fallback_redirect="dashboard")


@login_required
def operator_monthly_export(request, file_type):
    if is_admin(request.user):
        return export_monthly(request, file_type)
    site = assigned_site(request.user)
    if site is None:
        messages.error(request, "No site is assigned to your account.")
        return redirect("history")
    requested_site_id = request.GET.get("site", "").strip()
    if requested_site_id and str(site.pk) != requested_site_id:
        raise PermissionDenied
    return _run_monthly_export(request, file_type, site=site, fallback_redirect="history")


@login_required
@user_passes_test(is_admin)
def admin_settings(request):
    settings_obj = GlobalSettings.get_solo()
    deadline_form = DeadlineForm(instance=settings_obj)
    target_admin = request.user
    credential_form = AdminCredentialChangeForm(user=target_admin)

    if request.method == "POST":
        if "save_deadline" in request.POST:
            deadline_form = DeadlineForm(request.POST, instance=settings_obj)
            if deadline_form.is_valid():
                settings_obj = deadline_form.save(commit=False)
                settings_obj.updated_by = request.user
                settings_obj.save()
                messages.success(request, "The shared daily cutoff was updated.")
                return redirect("admin-settings")
        elif "change_credentials" in request.POST:
            credential_form = AdminCredentialChangeForm(request.POST, user=target_admin)
            if credential_form.is_valid():
                new_username = credential_form.cleaned_data["new_username"]
                new_password = credential_form.cleaned_data.get("new_password")
                target_admin.username = new_username
                if new_password:
                    target_admin.set_password(new_password)
                target_admin.save()
                update_session_auth_hash(request, target_admin)
                messages.success(request, f"Admin credentials updated successfully! UserID is now '{new_username}'.")
                return redirect("admin-settings")

    return render(request, "checklist/admin_settings.html", {
        "form": deadline_form,
        "credential_form": credential_form,
        "target_admin": target_admin,
    })


@login_required
@user_passes_test(is_admin)
def manage_users(request):
    if request.method == "POST":
        form = UserProvisionForm(request.POST)
        if form.is_valid():
            with transaction.atomic():
                user = form.save()
                UserSite.objects.create(user=user, site=form.cleaned_data["site"])
            messages.success(request, f"User {user.username} was created and assigned to {form.cleaned_data['site']}.")
            return redirect("manage-users")
    else:
        form = UserProvisionForm()
    users = User.objects.filter(is_staff=False).select_related("site_assignment__site").order_by("username")
    return render(request, "checklist/manage_users.html", {"form": form, "users": users})


@login_required
@user_passes_test(is_admin)
def edit_user(request, user_id):
    user = get_object_or_404(User.objects.filter(is_staff=False), pk=user_id)
    if request.method == "POST":
        form = UserEditForm(request.POST, instance=user)
        if form.is_valid():
            with transaction.atomic():
                form.save()
                site = form.cleaned_data.get("site")
                assignment = UserSite.objects.filter(user=user).first()
                if site:
                    if assignment:
                        assignment.site = site
                        assignment.save(update_fields=["site"])
                    else:
                        UserSite.objects.create(user=user, site=site)
                elif assignment:
                    assignment.delete()
            messages.success(request, f"User {user.username} was updated.")
            return redirect("manage-users")
    else:
        form = UserEditForm(instance=user)
    return render(request, "checklist/edit_user.html", {"form": form, "managed_user": user})


@login_required
@user_passes_test(is_admin)
def manage_sites(request):
    if request.method == "POST":
        form = SiteForm(request.POST)
        if form.is_valid():
            site = form.save()
            messages.success(request, "Site saved.")
            return redirect("manage-sites")
    else:
        form = SiteForm()
    sites = Site.objects.prefetch_related("user_assignment__user", "equipment").order_by("name")
    return render(request, "checklist/manage_sites.html", {"form": form, "sites": sites})


@login_required
@user_passes_test(is_admin)
def edit_site(request, site_id):
    site = get_object_or_404(Site.objects.prefetch_related("equipment", "user_assignment__user", "checklists__template"), pk=site_id)
    if request.method == "POST":
        form = SiteForm(request.POST, instance=site)
        if form.is_valid():
            form.save()
            messages.success(request, "Site details updated.")
            return redirect("edit-site", site_id=site.pk)
    else:
        form = SiteForm(instance=site)
    equipment_form = EquipmentForm()
    checklist_assign_form = SiteChecklistAssignForm(site=site)
    site_checklists = list(site.checklists.select_related("template").order_by("position", "template__name"))
    grouped_map = {}
    for sc in site_checklists:
        cat = sc.resolved_category()
        grouped_map.setdefault(cat, []).append(sc)
    grouped_site_checklists = [
        {"category": cat, "items": grouped_map[cat]}
        for cat in sorted(grouped_map.keys(), key=lambda c: (CATEGORY_ORDER.get(c, 99), c.lower()))
    ]
    return render(request, "checklist/edit_site.html", {
        "site": site,
        "form": form,
        "equipment_form": equipment_form,
        "checklist_assign_form": checklist_assign_form,
        "site_checklists": site_checklists,
        "grouped_site_checklists": grouped_site_checklists,
        "category_choices": CATEGORY_CHOICES,
    })


@login_required
@user_passes_test(is_admin)
def add_site_equipment(request, site_id):
    if request.method != "POST":
        return HttpResponseForbidden("Equipment changes must be submitted.")
    site = get_object_or_404(Site, pk=site_id)
    form = EquipmentForm(request.POST)
    if form.is_valid():
        equipment = form.save(commit=False)
        equipment.site = site
        try:
            equipment.save()
            messages.success(request, "Equipment added to the site.")
        except IntegrityError:
            messages.error(request, "That equipment is already listed for this site.")
    return redirect("edit-site", site_id=site.pk)


@login_required
@user_passes_test(is_admin)
def remove_site_equipment(request, site_id, equipment_id):
    if request.method != "POST":
        return HttpResponseForbidden("Equipment changes must be submitted.")
    equipment = get_object_or_404(SiteEquipment, pk=equipment_id, site_id=site_id)
    equipment.delete()
    messages.success(request, "Equipment removed from the site.")
    return redirect("edit-site", site_id=site_id)


@login_required
@user_passes_test(is_admin)
def assign_site_checklist(request, site_id):
    if request.method != "POST":
        return HttpResponseForbidden("Checklist assignment must be submitted.")
    site = get_object_or_404(Site, pk=site_id)
    form = SiteChecklistAssignForm(request.POST, site=site)
    if form.is_valid():
        sc = form.save(commit=False)
        sc.site = site
        try:
            sc.save()
            messages.success(request, f"Assigned '{sc.template.name}' to {site.name}.")
        except IntegrityError:
            messages.error(request, "That checklist is already assigned to this site.")
    else:
        for error in form.errors.values():
            messages.error(request, error.as_text())
    return redirect("edit-site", site_id=site.pk)


@login_required
@user_passes_test(is_admin)
def update_site_checklist_category(request, site_id, site_checklist_id):
    if request.method != "POST":
        return HttpResponseForbidden("Checklist changes must be submitted.")
    sc = get_object_or_404(SiteChecklist, pk=site_checklist_id, site_id=site_id)
    new_cat = request.POST.get("category", "").strip()[:60]
    sc.category = new_cat
    sc.save(update_fields=["category"])
    messages.success(request, f"Updated group for '{sc.template.name}' to '{sc.resolved_category()}'.")
    return redirect("edit-site", site_id=site_id)


@login_required
@user_passes_test(is_admin)
def remove_site_checklist(request, site_id, site_checklist_id):
    if request.method != "POST":
        return HttpResponseForbidden("Checklist changes must be submitted.")
    sc = get_object_or_404(SiteChecklist, pk=site_checklist_id, site_id=site_id)
    name = sc.template.name
    sc.delete()
    messages.success(request, f"Removed '{name}' from {sc.site.name}.")
    return redirect("edit-site", site_id=site_id)


@login_required
@user_passes_test(is_admin)
def manage_checklists(request):
    if request.method == "POST":
        form = ChecklistUploadForm(request.POST, request.FILES)
        if form.is_valid():
            template_name = form.cleaned_data["template_name"].strip()
            upload_file = form.cleaned_data["file"]
            file_name = upload_file.name.lower()
            try:
                if file_name.endswith((".xlsx", ".xls", ".xlx", ".xlsm", ".xlsb", ".xltx", ".xltm", ".ods")):
                    template, count = import_uploaded_excel(upload_file, template_name)
                elif file_name.endswith((".csv", ".tsv")):
                    template, count = import_uploaded_csv(upload_file, template_name)
                elif file_name.endswith(".pdf"):
                    template, count = import_uploaded_pdf(upload_file, template_name)
                else:
                    messages.error(request, "Unsupported file format. Please upload .xlsx, .xls, .xlx, .csv, or .pdf.")
                    return redirect("manage-checklists")

                # Apply explicit category if provided (otherwise importer auto-inferred it)
                chosen_category = form.cleaned_data.get("category", "").strip()
                if chosen_category:
                    template.category = chosen_category
                    template.save(update_fields=["category"])

                # If sites were selected, assign the new template to them
                sites_to_assign = form.cleaned_data.get("assign_to_sites")
                if sites_to_assign:
                    for s in sites_to_assign:
                        # find next position
                        max_pos = s.checklists.aggregate(models.Max("position"))["position__max"] or 0
                        SiteChecklist.objects.get_or_create(
                            site=s,
                            template=template,
                            defaults={"position": max_pos + 1, "active": True},
                        )

                messages.success(request, f"Checklist '{template.name}' (v{template.version}) uploaded with {count} checkpoints.")
                return redirect("preview-checklist-template", template_id=template.id)
            except Exception as e:
                messages.error(request, f"Error processing file: {e}")
                return redirect("manage-checklists")
        else:
            for field_errors in form.errors.values():
                for err in field_errors:
                    messages.error(request, err)
    else:
        form = ChecklistUploadForm()

    templates = ChecklistTemplate.objects.prefetch_related("items", "site_checklists__site").order_by("name", "-version")
    return render(request, "checklist/manage_checklists.html", {
        "form": form,
        "templates": templates,
    })


@login_required
@user_passes_test(is_admin)
def delete_checklist_template(request, template_id):
    if request.method != "POST":
        return HttpResponseForbidden("Action must be submitted via POST.")
    template = get_object_or_404(ChecklistTemplate, pk=template_id)
    if template.submissions.exists():
        template.active = False
        template.save(update_fields=["active"])
        messages.info(request, f"Checklist '{template.name}' is archived (has past submissions).")
    else:
        template.delete()
        messages.success(request, f"Checklist '{template.name}' deleted.")
    return redirect("manage-checklists")


@login_required
@user_passes_test(is_admin)
def preview_checklist_template(request, template_id):
    from collections import OrderedDict
    template = get_object_or_404(
        ChecklistTemplate.objects.prefetch_related("items", "site_checklists__site"),
        pk=template_id
    )
    items = list(template.items.order_by("position", "checkpoint_no"))

    sections_map = OrderedDict()
    for item in items:
        sec = item.section or "General Checks"
        sections_map.setdefault(sec, []).append(item)

    num_boolean = sum(1 for item in items if item.answer_type == ChecklistItem.AnswerType.BOOLEAN)
    num_numeric = sum(1 for item in items if item.answer_type == ChecklistItem.AnswerType.NUMBER)
    num_text = sum(1 for item in items if item.answer_type == ChecklistItem.AnswerType.TEXT)
    num_choices = sum(1 for item in items if item.answer_type == ChecklistItem.AnswerType.CHOICES)

    return render(request, "checklist/preview_template.html", {
        "template": template,
        "items": items,
        "sections_map": sections_map.items(),
        "total_items": len(items),
        "total_sections": len(sections_map),
        "num_boolean": num_boolean,
        "num_numeric": num_numeric,
        "num_text": num_text,
        "num_choices": num_choices,
        "assigned_sites": template.site_checklists.all(),
    })
