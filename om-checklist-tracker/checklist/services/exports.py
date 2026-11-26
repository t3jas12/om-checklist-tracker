from datetime import timedelta
from io import BytesIO
from pathlib import Path

from django.http import HttpResponse
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from .summary import summarize_submission_choices


def _equipment_label(submission):
    return (
        submission.equipment_snapshot
        or ", ".join(submission.site.equipment.values_list("name", flat=True))
        or submission.template.equipment_name
    )


def _split_section(section_text):
    """Split 'Section — Sub-section' into (section, subsection) when present."""
    if not section_text:
        return "", ""
    if " — " in section_text:
        sec, _, sub = section_text.partition(" — ")
        return sec.strip(), sub.strip()
    return section_text.strip(), ""


def _status_fill(value):
    low = (value or "").strip().lower()
    if low in {"yes", "ok", "operational", "pass", "true", "clean", "done", "working", "healthy"}:
        return PatternFill("solid", fgColor="E8F5EF")
    if low in {"no", "not ok", "non operational", "fail", "false", "not clean", "dirty", "faulty", "abnormal", "damaged"}:
        return PatternFill("solid", fgColor="FEE2E2")
    if low in {"spare", "n/a", "na", "pending"}:
        return PatternFill("solid", fgColor="FEF3C7")
    return None


def _attach_row_dropdown(ws, item, row_idx, start_col, end_col):
    """Attach an Excel DataValidation dropdown to the exported value cell(s) for this item."""
    if item.answer_type == "choices" and item.choices:
        opts_str = item.choices.replace('"', "").strip()
    elif item.answer_type == "boolean":
        opts_str = "YES,NO"
    else:
        return
    if not opts_str or len(opts_str) > 250:
        return
    dv = DataValidation(type="list", formula1=f'"{opts_str}"', allow_blank=True)
    start_ref = f"{get_column_letter(start_col)}{row_idx}"
    end_ref = f"{get_column_letter(end_col)}{row_idx}"
    dv.sqref = start_ref if start_col == end_col else f"{start_ref}:{end_ref}"
    ws.add_data_validation(dv)


def _unique_sheet_title(existing_titles, raw_title):
    base = "".join(ch for ch in raw_title if ch not in r'[]:*?/\\').strip() or "Checklist"
    candidate = base[:31]
    counter = 2
    while candidate in existing_titles:
        suffix = f" ({counter})"
        candidate = f"{base[:31 - len(suffix)]}{suffix}"
        counter += 1
    existing_titles.add(candidate)
    return candidate


# ---------------------------------------------------------------------------
# Single-submission exports
# ---------------------------------------------------------------------------

def excel_response(submission):
    template_path = Path(__file__).resolve().parents[2] / "Trafo_Checklist_Nov_2026_Populated.xlsx"
    use_trafo_sheet = template_path.exists() and "trafo" in submission.template.name.lower()
    workbook = load_workbook(template_path) if use_trafo_sheet else None
    if workbook is None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = _unique_sheet_title(set(), submission.template.name)

        HDR_FILL = PatternFill("solid", fgColor="123B32")
        HDR_FONT = Font(bold=True, color="FFFFFF", size=9)
        BODY_FONT = Font(size=9)
        BOLD_FONT = Font(bold=True, size=9)
        EVEN_FILL = PatternFill("solid", fgColor="F1F5F3")
        THIN = Side(style="thin", color="CBD5D1")
        BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
        LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
        CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

        items = list(submission.template.items.order_by("position", "checkpoint_no"))
        has_sub = any(" — " in (it.section or "") for it in items)

        sheet["A1"] = f"Document No. {submission.template.document_number}" if submission.template.document_number else submission.template.name
        sheet["A1"].font = BOLD_FONT
        sheet["D1"] = submission.site.company or submission.template.company or submission.site.name
        sheet["D1"].font = BOLD_FONT
        sheet["A2"] = f"Equipment Name : {_equipment_label(submission) or '—'}"
        sheet["A2"].font = BOLD_FONT
        sheet["D2"] = f"Date : {submission.checklist_date:%d %b %Y}   |   Approved By :"
        sheet["D2"].font = BOLD_FONT

        headers = ["Section"] + (["Panel / Sub-section"] if has_sub else []) + [
            "Check point no.", "Job Description", "Standard", "Method/Tool", "Frequency", "Reading / Status", "Comment"
        ]
        for ci, label in enumerate(headers, start=1):
            c = sheet.cell(4, ci, label)
            c.fill = HDR_FILL
            c.font = HDR_FONT
            c.alignment = CENTER
            c.border = BORDER
        sheet.row_dimensions[4].height = 24

        answer_by_item = {a.item_id: a for a in submission.answers.select_related("item")}
        prev_sec = None
        prev_sub = None
        row_idx = 5
        for item in items:
            ans = answer_by_item.get(item.id)
            val = ans.value if ans else ""
            com = ans.comment if ans else ""
            sec, sub = _split_section(item.section)
            sec_disp = sec if sec != prev_sec else ""
            sub_disp = sub if (sub != prev_sub or sec != prev_sec) else ""
            prev_sec = sec
            prev_sub = sub

            row_vals = ([sec_disp, sub_disp] if has_sub else [sec_disp]) + [
                item.checkpoint_no, item.description, item.standard or "—", item.method or "—", item.frequency or "Daily", val, com
            ]
            row_fill = EVEN_FILL if row_idx % 2 == 0 else PatternFill()
            val_col = len(row_vals) - 1
            for ci, v in enumerate(row_vals, start=1):
                cell = sheet.cell(row_idx, ci, v)
                cell.font = BODY_FONT
                cell.border = BORDER
                cell.alignment = CENTER if ci in ((3,) if has_sub else (2,)) or ci == val_col else LEFT
                cell.fill = (_status_fill(v) if ci == val_col else None) or row_fill
            _attach_row_dropdown(sheet, item, row_idx, val_col, val_col)
            row_idx += 1

        choice_summary = summarize_submission_choices(submission)
        row_idx += 1
        if choice_summary:
            summary_txt = "   |   ".join(f"{opt['label']}: {opt['count']}" for opt in choice_summary["option_counts"])
            if choice_summary["non_operational_labels"]:
                summary_txt += f"   (Non-Operational: {', '.join(choice_summary['non_operational_labels'])})"
            sheet.cell(row_idx, 1, "Status Summary:").font = BOLD_FONT
            sheet.cell(row_idx, 2, summary_txt).font = BODY_FONT
            row_idx += 1
        sheet.cell(row_idx, 1, "Remarks:").font = BOLD_FONT
        sheet.cell(row_idx, 2, submission.remarks or "—").font = BODY_FONT
        row_idx += 1
        sheet.cell(row_idx, 1, "Legends:").font = BOLD_FONT
        sheet.cell(row_idx, 2, "IF OK : YES   |   IF NOT OK : NO   |   D - Daily   |   W - Weekly").font = BODY_FONT
        row_idx += 1
        sheet.cell(row_idx, 1, "Checked & done by:").font = BOLD_FONT
        sheet.cell(row_idx, 2, submission.signer_name or submission.submitted_by.get_full_name() or submission.submitted_by.username).font = BODY_FONT
        sheet.cell(row_idx, 4, f"Submitted at: {submission.submitted_at:%Y-%m-%d %H:%M %Z}").font = BODY_FONT
        row_idx += 1
        sheet.cell(row_idx, 1, "Note:").font = BOLD_FONT
        sheet.cell(row_idx, 2, "If found abnormal inform Maintenance Department.").font = BODY_FONT

        widths = [18, 18, 10, 42, 34, 14, 11, 16, 26] if has_sub else [20, 10, 42, 34, 14, 11, 16, 26]
        for ci, w in enumerate(widths, start=1):
            sheet.column_dimensions[get_column_letter(ci)].width = w
    else:
        sheet = workbook["spv"] if "spv" in workbook.sheetnames else workbook.active
        # Remove populated sample daily values; exports contain only live submitted data.
        for row in range(7, sheet.max_row + 1):
            for column in range(8, sheet.max_column + 1):
                cell = sheet.cell(row, column)
                if cell.__class__.__name__ != "MergedCell":
                    cell.value = None
        day_column = 8 + submission.checklist_date.day - 1
        sheet["C3"] = f"Document No. {submission.template.document_number}"
        sheet["E3"] = submission.site.company or submission.template.company
        sheet["H4"] = f"Equipment Name: {_equipment_label(submission)}"
        sheet["H5"] = f"Month - {submission.checklist_date:%b-%Y}"
        for day in range(1, 32):
            sheet.cell(6, 7 + day, day)
        answer_by_item = {a.item_id: a for a in submission.answers.select_related("item")}
        items = list(submission.template.items.all())
        for item in items:
            row = item.position + 6
            answer = answer_by_item.get(item.id)
            if answer and row <= sheet.max_row:
                cell = sheet.cell(row, day_column, answer.value)
                cell.comment = None
        sheet["B35"] = f"Remarks: {submission.remarks}"
        sheet["C3"].comment = None
        sheet["B36"] = f"Submitted by: {submission.signer_name or submission.submitted_by.get_full_name() or submission.submitted_by.username}"
        sheet["E36"] = f"Submitted at: {submission.submitted_at.strftime('%Y-%m-%d %H:%M %Z')}"

    # Keep comments and remarks visible without relying on space in the imported workbook layout.
    review = workbook.create_sheet("Submission review")
    review.append(["Checkpoint", "Observation", "Reading / choice", "Comment"])
    for answer in submission.answers.select_related("item").order_by("item__position", "item__checkpoint_no"):
        review.append([answer.item.checkpoint_no, answer.item.description, answer.value, answer.comment])
    review.append([])
    review.append(["Overall remarks", submission.remarks])
    review.append(["Submitted by", submission.signer_name or submission.submitted_by.get_full_name() or submission.submitted_by.username])
    review.append(["Submitted at", submission.submitted_at.strftime("%Y-%m-%d %H:%M %Z")])

    output = BytesIO()
    workbook.save(output)
    response = HttpResponse(output.getvalue(), content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response["Content-Disposition"] = f'attachment; filename="{submission.site.name.replace(" ", "_")}_{submission.checklist_date}.xlsx"'
    return response


def pdf_response(submission):
    output = BytesIO()
    doc = SimpleDocTemplate(output, pagesize=landscape(A4), rightMargin=28, leftMargin=28, topMargin=28, bottomMargin=28)
    styles = getSampleStyleSheet()
    story = [Paragraph("O and M Checklist Tracker", styles["Title"]), Spacer(1, 8)]
    story.append(Paragraph(f"{submission.site.name} — {submission.checklist_date:%d %B %Y}", styles["Heading2"]))
    doc_part = f"Doc No.: {submission.template.document_number} | " if submission.template.document_number else ""
    story.append(Paragraph(f"{doc_part}Equipment: {_equipment_label(submission)} | Submitted by: {submission.signer_name or submission.submitted_by.username} | Submitted: {submission.submitted_at:%H:%M IST}", styles["Normal"]))
    choice_summary = summarize_submission_choices(submission)
    if choice_summary:
        parts = "  |  ".join(f"{opt['label']}: {opt['count']}" for opt in choice_summary["option_counts"])
        non_op_note = f"  —  Non-Operational: {', '.join(choice_summary['non_operational_labels'])}" if choice_summary["non_operational_labels"] else ""
        story.append(Spacer(1, 6))
        story.append(Paragraph(f"<b>Status Summary ({choice_summary['total_items']} total):</b> {parts}{non_op_note}", styles["Normal"]))
    story.append(Spacer(1, 12))
    data = [["#", "Section", "Observation", "Standard", "Reading / Status", "Comment"]]
    for answer in submission.answers.select_related("item").order_by("item__position"):
        item = answer.item
        data.append([str(item.checkpoint_no), item.section, Paragraph(item.description, styles["BodyText"]), Paragraph(item.standard or "—", styles["BodyText"]), answer.value or "—", Paragraph(answer.comment or "—", styles["BodyText"])])
    table = Table(data, colWidths=[28, 90, 220, 170, 100, 150], repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#123b32")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#cbd5d1")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f1f5f3")]),
        ("PADDING", (0, 0), (-1, -1), 5),
    ]))
    story.extend([
        table,
        Spacer(1, 10),
        Paragraph(f"<b>Remarks:</b> {submission.remarks or '—'}", styles["Normal"]),
        Spacer(1, 6),
        Paragraph(
            f"<b>Legends:</b> IF OK : YES  |  IF NOT OK : NO  |  D - Daily  |  W - Weekly  &nbsp;&nbsp;·&nbsp;&nbsp; "
            f"<b>Checked &amp; done by:</b> {submission.signer_name or submission.submitted_by.get_full_name() or submission.submitted_by.username}  &nbsp;&nbsp;·&nbsp;&nbsp; "
            f"<b>Verify by:</b> ____________  &nbsp;&nbsp;·&nbsp;&nbsp; "
            f"<b>Note:</b> If Found abnormal inform Maintenance Department.",
            styles["Normal"],
        ),
    ])
    doc.build(story)
    response = HttpResponse(output.getvalue(), content_type="application/pdf")
    response["Content-Disposition"] = f'attachment; filename="{submission.site.name.replace(" ", "_")}_{submission.checklist_date}.pdf"'
    return response


# ---------------------------------------------------------------------------
# Multi-day (monthly / date-range) exports
# ---------------------------------------------------------------------------

def _period_label(start_date, end_date):
    if start_date.year == end_date.year and start_date.month == end_date.month and start_date.day == 1:
        import calendar
        last_day = calendar.monthrange(start_date.year, start_date.month)[1]
        if end_date.day == last_day:
            return f"Month - {start_date:%b-%Y}"
    return f"{start_date:%d %b %Y} to {end_date:%d %b %Y}"


def _build_range_excel(submissions, site, start_date, end_date):
    """
    Build a full-month / date-range Excel workbook in O&M matrix format:
      • Top block: Document No., Company, Equipment Name, Month/Period
      • Table header: Section, (Panel/Sub-section), Check point no., Job Description,
        Standard, Method/Tool, Frequency, and 1 column per calendar day (1..30/31)
      • Footer: Remarks, Legends (IF OK/NOT OK, D-Daily, W-Weekly), Signature, Note
      • Extra sheet: Comments & Remarks log across the month
    """
    total_days = (end_date - start_date).days + 1
    days = [start_date + timedelta(days=i) for i in range(total_days)]
    multi_month = (start_date.year, start_date.month) != (end_date.year, end_date.month)
    period_str = _period_label(start_date, end_date)

    # Group submissions by (site_id, template_id) so both single-site and all-sites exports work cleanly
    grouped = {}
    for sub in submissions:
        grouped.setdefault((sub.site_id, sub.template_id), []).append(sub)

    wb = Workbook()
    wb.remove(wb.active)
    used_titles = set()

    if not grouped:
        ws = wb.create_sheet("No Data")
        ws["A1"] = f"No submissions found for {site.name if site else 'selected sites'} ({period_str})."
        output = BytesIO()
        wb.save(output)
        return output.getvalue()

    HDR_FILL = PatternFill("solid", fgColor="123B32")
    SUB_HDR_FILL = PatternFill("solid", fgColor="1B4D42")
    HDR_FONT = Font(bold=True, color="FFFFFF", size=9)
    META_FONT = Font(bold=True, size=9)
    BODY_FONT = Font(size=9)
    SMALL_BOLD = Font(bold=True, size=8, color="123B32")
    CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
    THIN = Side(style="thin", color="CBD5D1")
    BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    EVEN_FILL = PatternFill("solid", fgColor="F6F9F7")

    comments_rows = []

    for (_site_id, _tmpl_id), subs in grouped.items():
        subs = sorted(subs, key=lambda s: s.checklist_date)
        first_sub = subs[0]
        sub_site = first_sub.site
        template = first_sub.template
        items = list(template.items.order_by("position", "checkpoint_no"))
        has_sub = any(" — " in (it.section or "") for it in items)

        raw_title = f"{sub_site.name[:12]} - {template.name}" if site is None else template.name
        ws = wb.create_sheet(title=_unique_sheet_title(used_titles, raw_title))

        meta_cols = (
            ["Section", "Panel / Sub-section", "Check point no.", "Job Description", "Standard", "Method/Tool", "Frequency"]
            if has_sub
            else ["Section", "Check point no.", "Job Description", "Standard", "Method/Tool", "Frequency"]
        )
        n_meta = len(meta_cols)
        last_col = n_meta + len(days)
        equip = _equipment_label(first_sub) or "—"
        comp = sub_site.company or template.company or sub_site.name

        # --- Top Metadata Banner (Rows 1–3) ---
        ws.cell(1, 1, f"Document No. {template.document_number}" if template.document_number else template.name).font = SMALL_BOLD
        ws.cell(1, min(n_meta, 5), comp).font = SMALL_BOLD
        ws.cell(2, 1, f"Site: {sub_site.name}").font = SMALL_BOLD
        ws.cell(2, min(n_meta, 5), f"Equipment Name : {equip}").font = SMALL_BOLD
        if last_col > n_meta + 2:
            ws.cell(2, n_meta + 2, "Approved By :").font = SMALL_BOLD

        # --- Table Header (Rows 4 & 5) ---
        for ci, label in enumerate(meta_cols, start=1):
            ws.merge_cells(start_row=4, start_column=ci, end_row=5, end_column=ci)
            cell = ws.cell(4, ci, label)
            cell.fill = HDR_FILL
            cell.font = HDR_FONT
            cell.alignment = CENTER
            cell.border = BORDER
            ws.cell(5, ci).border = BORDER
            ws.cell(5, ci).fill = HDR_FILL

        ws.merge_cells(start_row=4, start_column=n_meta + 1, end_row=4, end_column=last_col)
        m_cell = ws.cell(4, n_meta + 1, period_str)
        m_cell.fill = HDR_FILL
        m_cell.font = HDR_FONT
        m_cell.alignment = CENTER
        for ci in range(n_meta + 1, last_col + 1):
            ws.cell(4, ci).border = BORDER
            ws.cell(4, ci).fill = HDR_FILL

        for di, day in enumerate(days):
            ci = n_meta + 1 + di
            day_label = f"{day:%d %b}" if multi_month else day.day
            d_cell = ws.cell(5, ci, day_label)
            d_cell.fill = SUB_HDR_FILL
            d_cell.font = HDR_FONT
            d_cell.alignment = CENTER
            d_cell.border = BORDER
            ws.column_dimensions[get_column_letter(ci)].width = 7 if not multi_month else 9

        ws.row_dimensions[4].height = 20
        ws.row_dimensions[5].height = 20

        # Build lookup: day -> {item_id -> answer}
        day_answers = {}
        submitters = []
        seen_submitters = set()
        for sub in subs:
            signer = sub.signer_name or sub.submitted_by.get_full_name() or sub.submitted_by.username
            if signer and signer not in seen_submitters:
                seen_submitters.add(signer)
                submitters.append(signer)

            ans_list = list(sub.answers.select_related("item").all())
            day_answers[sub.checklist_date] = {a.item_id: a for a in ans_list}
            if sub.remarks:
                comments_rows.append({
                    "date": sub.checklist_date,
                    "site": sub_site.name,
                    "template": template.name,
                    "type": "Overall remark",
                    "checkpoint": "",
                    "observation": "",
                    "text": sub.remarks,
                    "by": signer,
                })
            for a in ans_list:
                if a.comment:
                    comments_rows.append({
                        "date": sub.checklist_date,
                        "site": sub_site.name,
                        "template": template.name,
                        "type": "Item comment",
                        "checkpoint": a.item.checkpoint_no,
                        "observation": a.item.description,
                        "text": a.comment,
                        "by": signer,
                    })

        # --- Data Rows ---
        prev_sec = None
        prev_sub = None
        has_choices = False
        row_idx = 6
        for item in items:
            if item.answer_type == "choices":
                has_choices = True
            row_fill = EVEN_FILL if row_idx % 2 == 0 else PatternFill()
            sec, sub_sec = _split_section(item.section)
            sec_disp = sec if sec != prev_sec else ""
            sub_disp = sub_sec if (sub_sec != prev_sub or sec != prev_sec) else ""
            prev_sec = sec
            prev_sub = sub_sec

            meta_vals = (
                [sec_disp, sub_disp, item.checkpoint_no, item.description, item.standard or "—", item.method or "—", item.frequency or "Daily"]
                if has_sub
                else [sec_disp, item.checkpoint_no, item.description, item.standard or "—", item.method or "—", item.frequency or "Daily"]
            )
            desc_col = 4 if has_sub else 3
            cp_col = 3 if has_sub else 2
            freq_col = n_meta

            for ci, val in enumerate(meta_vals, start=1):
                cell = ws.cell(row_idx, ci, val)
                cell.font = META_FONT if ci == desc_col else BODY_FONT
                cell.alignment = CENTER if ci in (cp_col, freq_col) else LEFT
                cell.border = BORDER
                cell.fill = row_fill

            for di, day in enumerate(days):
                ci = n_meta + 1 + di
                ans = day_answers.get(day, {}).get(item.id)
                val = ans.value if ans else ""
                cell = ws.cell(row_idx, ci, val)
                cell.alignment = CENTER
                cell.border = BORDER
                cell.font = BODY_FONT
                cell.fill = _status_fill(val) or row_fill
                if has_choices and val and len(val) > 8:
                    ws.column_dimensions[get_column_letter(ci)].width = max(
                        ws.column_dimensions[get_column_letter(ci)].width or 7, 14
                    )

            _attach_row_dropdown(ws, item, row_idx, n_meta + 1, last_col)
            row_idx += 1

        # --- Dynamic Daily Option Counts (when template has choices items, e.g. inv3) ---
        if has_choices:
            day_summaries = {sub.checklist_date: summarize_submission_choices(sub) for sub in subs}
            ordered_labels = []
            label_kind = {}
            for summ in day_summaries.values():
                if not summ:
                    continue
                for opt in summ["option_counts"]:
                    if opt["label"] not in label_kind:
                        label_kind[opt["label"]] = opt["kind"]
                        ordered_labels.append(opt["label"])

            for opt_label in ordered_labels:
                kind = label_kind[opt_label]
                row_tint = (
                    PatternFill("solid", fgColor="FEE2E2") if kind == "fault"
                    else PatternFill("solid", fgColor="FEF3C7") if kind == "spare"
                    else PatternFill("solid", fgColor="E8F5EF") if kind == "ok"
                    else EVEN_FILL
                )
                for ci in range(1, n_meta + 1):
                    c = ws.cell(row_idx, ci, f"{opt_label} Count" if ci == desc_col else "")
                    c.font = SMALL_BOLD
                    c.alignment = LEFT
                    c.border = BORDER
                    c.fill = row_tint
                for di, day in enumerate(days):
                    ci = n_meta + 1 + di
                    summ = day_summaries.get(day)
                    cnt = ""
                    if summ:
                        cnt = next((o["count"] for o in summ["option_counts"] if o["label"] == opt_label), 0)
                    c = ws.cell(row_idx, ci, cnt)
                    c.font = SMALL_BOLD
                    c.alignment = CENTER
                    c.border = BORDER
                    c.fill = row_tint
                row_idx += 1

        # --- Footer: Remarks, Legends, Signature, Note (matches O&M sheet) ---
        sheet_has_comments = any(
            c["site"] == sub_site.name and c["template"] == template.name for c in comments_rows
        )
        row_idx += 1
        ws.cell(row_idx, 1, "Remarks:").font = SMALL_BOLD
        ws.cell(row_idx, 2, "See 'Comments & Remarks' sheet for daily remarks and item comments." if sheet_has_comments else "No abnormalities / remarks recorded.").font = BODY_FONT
        row_idx += 1
        ws.cell(row_idx, 1, "Legends:").font = SMALL_BOLD
        ws.cell(row_idx, 2, "IF OK : YES   |   IF NOT OK : NO   |   D - Daily   |   W - Weekly").font = BODY_FONT
        row_idx += 1
        ws.cell(row_idx, 1, "Signature").font = SMALL_BOLD
        ws.cell(row_idx, 2, f"Checked & done by: {', '.join(submitters) or '—'}").font = BODY_FONT
        ws.cell(row_idx, min(n_meta, 5), "Verify by:").font = SMALL_BOLD
        row_idx += 1
        ws.cell(row_idx, 1, "Note:").font = SMALL_BOLD
        ws.cell(row_idx, 2, "If Found abnormal inform Maintenance Department.").font = BODY_FONT

        col_widths = [16, 18, 8, 40, 32, 13, 10] if has_sub else [18, 8, 42, 34, 13, 10]
        for ci, w in enumerate(col_widths, start=1):
            ws.column_dimensions[get_column_letter(ci)].width = w

        ws.freeze_panes = f"{get_column_letter(n_meta + 1)}6"

    # --- Comments & Remarks sheet ---
    if comments_rows:
        cws = wb.create_sheet(_unique_sheet_title(used_titles, "Comments & Remarks"))
        hdr = ["Date", "Site", "Checklist", "Type", "Checkpoint #", "Observation", "Comment / Remark", "Submitted by"]
        for ci, h in enumerate(hdr, start=1):
            cell = cws.cell(1, ci, h)
            cell.fill = HDR_FILL
            cell.font = HDR_FONT
            cell.alignment = CENTER
            cell.border = BORDER
        for ri, row in enumerate(sorted(comments_rows, key=lambda r: (r["date"], r["site"])), start=2):
            vals = [row["date"].strftime("%d %b %Y"), row["site"], row["template"], row["type"], row["checkpoint"], row["observation"], row["text"], row["by"]]
            fill = EVEN_FILL if ri % 2 == 0 else PatternFill()
            for ci, v in enumerate(vals, start=1):
                cell = cws.cell(ri, ci, v)
                cell.font = BODY_FONT
                cell.alignment = LEFT
                cell.border = BORDER
                cell.fill = fill
        for col, width in zip(["A", "B", "C", "D", "E", "F", "G", "H"], [14, 18, 24, 14, 12, 34, 48, 18]):
            cws.column_dimensions[col].width = width

    output = BytesIO()
    wb.save(output)
    return output.getvalue()


def _build_range_pdf(submissions, site, start_date, end_date):
    """
    Build a Full-Month View landscape A4 PDF where days appear across columns
    (split cleanly into <=16-day chunks when the range exceeds 16 days so every
    reading remains crisp and legible), followed by Legends, Signatures, and a
    Daily Remarks & Comments log.
    """
    output = BytesIO()
    doc = SimpleDocTemplate(output, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    styles = getSampleStyleSheet()
    cell_style = ParagraphStyle("MatrixCell", parent=styles["BodyText"], fontSize=6.5, leading=8)
    day_cell_style = ParagraphStyle("DayCell", parent=styles["BodyText"], fontSize=6.0, leading=7.5, alignment=1)
    hdr_style = ParagraphStyle("MatrixHdr", parent=styles["BodyText"], fontSize=6.5, leading=8, textColor=colors.white, fontName="Helvetica-Bold")
    day_hdr_style = ParagraphStyle("DayHdr", parent=hdr_style, alignment=1)

    total_days = (end_date - start_date).days + 1
    days = [start_date + timedelta(days=i) for i in range(total_days)]
    multi_month = (start_date.year, start_date.month) != (end_date.year, end_date.month)
    period_str = _period_label(start_date, end_date)

    grouped = {}
    for sub in submissions:
        grouped.setdefault((sub.site_id, sub.template_id), []).append(sub)

    story = []
    if not grouped:
        story.append(Paragraph("O&M Monthly Checklist Report", styles["Title"]))
        story.append(Spacer(1, 10))
        story.append(Paragraph(f"No submissions found for {site.name if site else 'selected sites'} ({period_str}).", styles["Normal"]))
        doc.build(story)
        return output.getvalue()

    MAX_DAYS_PER_TABLE = 16
    day_chunks = [days[i:i + MAX_DAYS_PER_TABLE] for i in range(0, len(days), MAX_DAYS_PER_TABLE)]
    USABLE_WIDTH = 802.0  # landscape A4 (842pt) - 40pt horizontal margins

    for group_idx, ((_site_id, _tmpl_id), subs) in enumerate(grouped.items()):
        if group_idx > 0:
            story.append(PageBreak())

        subs = sorted(subs, key=lambda s: s.checklist_date)
        first_sub = subs[0]
        sub_site = first_sub.site
        template = first_sub.template
        items = list(template.items.order_by("position", "checkpoint_no"))
        equip = _equipment_label(first_sub) or "—"
        comp = sub_site.company or template.company or sub_site.name

        story.append(Paragraph(f"O&amp;M Monthly Checklist — {template.name}", styles["Heading2"]))
        doc_part = f"Doc No.: {template.document_number}  |  " if template.document_number else ""
        story.append(Paragraph(
            f"<b>Site:</b> {sub_site.name} ({comp})  |  {doc_part}<b>Equipment:</b> {equip}  |  "
            f"<b>Period:</b> {period_str}  |  <b>Submitted days:</b> {len(subs)}/{len(days)}",
            styles["Normal"],
        ))
        story.append(Spacer(1, 8))

        day_answers = {}
        submitters = []
        seen_submitters = set()
        notes_log = []
        for sub in subs:
            signer = sub.signer_name or sub.submitted_by.get_full_name() or sub.submitted_by.username
            if signer and signer not in seen_submitters:
                seen_submitters.add(signer)
                submitters.append(signer)
            ans_list = list(sub.answers.select_related("item").all())
            day_answers[sub.checklist_date] = {a.item_id: a for a in ans_list}
            if sub.remarks:
                notes_log.append((sub.checklist_date, "Overall remark", sub.remarks, signer))
            for a in ans_list:
                if a.comment:
                    notes_log.append((sub.checklist_date, f"CP {a.item.checkpoint_no}", a.comment, signer))

        for chunk_idx, chunk in enumerate(day_chunks):
            if chunk_idx > 0:
                story.append(Spacer(1, 10))
                story.append(Paragraph(
                    f"<b>Days {chunk[0]:%d %b} – {chunk[-1]:%d %b %Y}</b>",
                    styles["Normal"],
                ))
                story.append(Spacer(1, 4))

            # Allocate column widths cleanly to fill USABLE_WIDTH (802pt)
            day_col_w = min(38.0, max(26.0, (USABLE_WIDTH - 338.0) / len(chunk)))
            days_total_w = day_col_w * len(chunk)
            meta_budget = USABLE_WIDTH - days_total_w
            cp_w = 20.0
            sec_w = round(meta_budget * 0.20, 1)
            desc_w = round(meta_budget * 0.46, 1)
            std_w = round(meta_budget - cp_w - sec_w - desc_w, 1)

            header_row = [
                Paragraph("#", hdr_style),
                Paragraph("Section", hdr_style),
                Paragraph("Job Description", hdr_style),
                Paragraph("Standard", hdr_style),
            ] + [
                Paragraph(f"{d.day}<br/>{d:%b}" if multi_month else str(d.day), day_hdr_style)
                for d in chunk
            ]
            rows_data = [header_row]
            style_cmds = [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#123b32")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#cbd5d1")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f4f7f5")]),
                ("PADDING", (0, 0), (-1, -1), 3),
            ]

            prev_sec = None
            for ri, item in enumerate(items, start=1):
                sec_label = item.section if item.section != prev_sec else ""
                prev_sec = item.section
                row = [
                    Paragraph(str(item.checkpoint_no), day_cell_style),
                    Paragraph(sec_label or "", cell_style),
                    Paragraph(item.description, cell_style),
                    Paragraph(item.standard or "—", cell_style),
                ]
                for di, d in enumerate(chunk):
                    ans = day_answers.get(d, {}).get(item.id)
                    val = ans.value if ans else ""
                    low = val.strip().lower()
                    if low in {"no", "not ok", "non operational", "fail", "false", "not clean", "dirty", "faulty", "abnormal", "damaged"}:
                        style_cmds.append(("BACKGROUND", (4 + di, ri), (4 + di, ri), colors.HexColor("#fee2e2")))
                    elif low in {"spare", "n/a", "na", "pending"}:
                        style_cmds.append(("BACKGROUND", (4 + di, ri), (4 + di, ri), colors.HexColor("#fef3c7")))
                    row.append(Paragraph(val or "—", day_cell_style))
                rows_data.append(row)

            table = Table(rows_data, colWidths=[cp_w, sec_w, desc_w, std_w] + [day_col_w] * len(chunk), repeatRows=1)
            table.setStyle(TableStyle(style_cmds))
            story.append(table)

        day_non_op = [
            f"{sub.checklist_date:%d %b}: {summ['non_operational_count']} Non-Op"
            for sub in subs
            for summ in [summarize_submission_choices(sub)]
            if summ
        ]
        if day_non_op:
            story.append(Spacer(1, 6))
            story.append(Paragraph(f"<b>Daily Non-Operational Summary:</b> {'  |  '.join(day_non_op)}", cell_style))

        story.append(Spacer(1, 8))
        story.append(Paragraph(
            f"<b>Legends:</b> IF OK : YES  |  IF NOT OK : NO  |  D - Daily  |  W - Weekly  &nbsp;&nbsp;·&nbsp;&nbsp; "
            f"<b>Checked &amp; done by:</b> {', '.join(submitters) or '—'}  &nbsp;&nbsp;·&nbsp;&nbsp; "
            f"<b>Note:</b> If Found abnormal inform Maintenance Department.",
            cell_style,
        ))

        if notes_log:
            story.append(Spacer(1, 8))
            story.append(Paragraph("<b>Daily Remarks &amp; Item Comments</b>", styles["Normal"]))
            story.append(Spacer(1, 4))
            note_rows = [[
                Paragraph("Date", hdr_style),
                Paragraph("Ref", hdr_style),
                Paragraph("Comment / Remark", hdr_style),
                Paragraph("Submitted by", hdr_style),
            ]]
            for d, ref, txt, signer in sorted(notes_log, key=lambda x: x[0]):
                note_rows.append([
                    Paragraph(f"{d:%d %b %Y}", cell_style),
                    Paragraph(ref, cell_style),
                    Paragraph(txt, cell_style),
                    Paragraph(signer, cell_style),
                ])
            nt = Table(note_rows, colWidths=[75, 85, 512, 130], repeatRows=1)
            nt.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#123b32")),
                ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#cbd5d1")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f4f7f5")]),
                ("PADDING", (0, 0), (-1, -1), 4),
            ]))
            story.append(nt)

    doc.build(story)
    return output.getvalue()


def range_excel_response(submissions, site, start_date, end_date):
    """Django HttpResponse for a full-month / multi-day Excel export."""
    data = _build_range_excel(submissions, site, start_date, end_date)
    site_slug = site.name.replace(" ", "_") if site else "All_Sites"
    if start_date.year == end_date.year and start_date.month == end_date.month and start_date.day == 1:
        label = f"{site_slug}_{start_date:%Y-%m}"
    else:
        label = f"{site_slug}_{start_date:%Y-%m-%d}_to_{end_date:%Y-%m-%d}"
    response = HttpResponse(data, content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response["Content-Disposition"] = f'attachment; filename="{label}.xlsx"'
    return response


def range_pdf_response(submissions, site, start_date, end_date):
    """Django HttpResponse for a full-month / multi-day PDF export."""
    data = _build_range_pdf(submissions, site, start_date, end_date)
    site_slug = site.name.replace(" ", "_") if site else "All_Sites"
    if start_date.year == end_date.year and start_date.month == end_date.month and start_date.day == 1:
        label = f"{site_slug}_{start_date:%Y-%m}"
    else:
        label = f"{site_slug}_{start_date:%Y-%m-%d}_to_{end_date:%Y-%m-%d}"
    response = HttpResponse(data, content_type="application/pdf")
    response["Content-Disposition"] = f'attachment; filename="{label}.pdf"'
    return response
