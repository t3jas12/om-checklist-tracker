from datetime import date, datetime, time
from io import BytesIO
from unittest.mock import patch

from django.contrib.auth import get_user_model
from django.core.files.uploadedfile import SimpleUploadedFile
from django.test import TestCase
from django.utils import timezone
from openpyxl import Workbook, load_workbook

from .models import Answer, ChecklistItem, ChecklistTemplate, GlobalSettings, Site, SiteChecklist, SiteEquipment, Submission, UserSite

User = get_user_model()


class ChecklistWorkflowTests(TestCase):
    def setUp(self):
        self.site = Site.objects.create(name="North Solar")
        SiteEquipment.objects.create(site=self.site, name="Transformer 1")
        self.user = User.objects.create_user(username="operator", password="secure-test-password")
        UserSite.objects.create(user=self.user, site=self.site)
        self.template = ChecklistTemplate.objects.create(name="Test Trafo", version=1)
        self.number_item = ChecklistItem.objects.create(template=self.template, checkpoint_no=1, description="Winding temperature", standard="Below 80 C", answer_type="number", position=1)
        self.boolean_item = ChecklistItem.objects.create(template=self.template, checkpoint_no=2, description="Oil leak check", standard="No leakage", answer_type="boolean", position=2)
        settings = GlobalSettings.get_solo()
        settings.deadline = time(23, 59)
        settings.save()

    def form_data(self, **extra):
        values = {
            f"item_{self.number_item.id}": "62.5",
            f"comment_{self.number_item.id}": "Within expected range",
            f"item_{self.boolean_item.id}": "YES",
            f"comment_{self.boolean_item.id}": "No visible leak",
            "remarks": "Routine inspection completed",
        }
        values.update(extra)
        return values

    def post_preview(self, data=None):
        return self.client.post("/checklist/", data or self.form_data())

    def test_preview_does_not_persist_submission_and_shows_all_values(self):
        self.client.force_login(self.user)
        response = self.post_preview()
        self.assertEqual(response.status_code, 200)
        self.assertTemplateUsed(response, "checklist/checklist_preview.html")
        self.assertContains(response, "62.5")
        self.assertContains(response, "Within expected range")
        self.assertContains(response, "Routine inspection completed")
        self.assertFalse(Submission.objects.exists())
        self.assertFalse(Answer.objects.exists())

    def test_user_can_finalize_after_review_and_read_saved_submission(self):
        self.client.force_login(self.user)
        self.post_preview()
        lockin = self.client.get("/checklist/finalize/")
        self.assertEqual(lockin.status_code, 200)
        self.assertTemplateUsed(lockin, "checklist/checklist_lockin.html")
        response = self.client.post("/checklist/finalize/", {"confirm_lockin": "yes"})
        submission = Submission.objects.get(site=self.site, checklist_date=timezone.localdate())
        self.assertRedirects(response, f"/submissions/{submission.pk}/review/")
        self.assertEqual(submission.answers.count(), 2)
        self.assertEqual(submission.answers.get(item=self.number_item).comment, "Within expected range")
        self.assertEqual(submission.remarks, "Routine inspection completed")
        review = self.client.get(f"/submissions/{submission.pk}/review/")
        self.assertEqual(review.status_code, 200)
        self.assertContains(review, "No visible leak")
        self.assertContains(review, "Routine inspection completed")
        self.assertEqual(submission.equipment_snapshot, "Transformer 1")

    def test_invalid_preview_does_not_create_draft_or_submission(self):
        self.client.force_login(self.user)
        response = self.post_preview(self.form_data(**{f"item_{self.number_item.id}": "not a number"}))
        self.assertEqual(response.status_code, 200)
        self.assertFalse(Submission.objects.exists())
        self.assertFalse(self.client.session.get("checklist_draft_" + str(self.user.pk)))

    def test_cutoff_is_checked_for_preview_and_final_lockin(self):
        settings = GlobalSettings.get_solo()
        settings.deadline = time(17, 0)
        settings.save()
        self.client.force_login(self.user)
        with patch("checklist.models.timezone.localtime", return_value=timezone.make_aware(datetime(2026, 11, 20, 16, 59))):
            self.post_preview()
        with patch("checklist.models.timezone.localtime", return_value=timezone.make_aware(datetime(2026, 11, 20, 16, 59))):
            self.client.get("/checklist/finalize/")
        with patch("checklist.models.timezone.localtime", return_value=timezone.make_aware(datetime(2026, 11, 20, 17, 0))):
            response = self.client.post("/checklist/finalize/", {"confirm_lockin": "yes"})
        self.assertRedirects(response, "/")
        self.assertFalse(Submission.objects.exists())

    def test_finalize_requires_existing_server_side_preview(self):
        self.client.force_login(self.user)
        response = self.client.get("/checklist/finalize/")
        self.assertRedirects(response, "/checklist/")
        self.assertFalse(Submission.objects.exists())

    def test_lockin_requires_explicit_confirmation(self):
        self.client.force_login(self.user)
        self.post_preview()
        response = self.client.post("/checklist/finalize/", {})
        self.assertRedirects(response, "/checklist/finalize/")
        self.assertFalse(Submission.objects.exists())

    def test_regular_user_cannot_review_another_users_submission_or_download_pdf(self):
        other = User.objects.create_user(username="other", password="secure-test-password")
        other_site = Site.objects.create(name="South Solar")
        UserSite.objects.create(user=other, site=other_site)
        submission = Submission.objects.create(site=other_site, template=self.template, submitted_by=other, checklist_date=date(2026, 11, 19))
        self.client.force_login(self.user)
        self.assertEqual(self.client.get(f"/submissions/{submission.pk}/review/").status_code, 403)
        self.assertEqual(self.client.get(f"/admin/submissions/{submission.pk}/pdf/").status_code, 302)

    def test_operator_excel_review_contains_answers_and_comments(self):
        submission = Submission.objects.create(site=self.site, template=self.template, submitted_by=self.user, checklist_date=date(2026, 11, 19), remarks="All good")
        Answer.objects.create(submission=submission, item=self.number_item, value="65", comment="Checked")
        self.client.force_login(self.user)
        response = self.client.get(f"/submissions/{submission.pk}/excel/")
        self.assertEqual(response["Content-Type"], "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        workbook = load_workbook(BytesIO(response.content), data_only=True)
        self.assertIn("Submission review", workbook.sheetnames)
        rows = list(workbook["Submission review"].values)
        self.assertIn((1, "Winding temperature", "65", "Checked"), rows)
        self.assertIn(("Overall remarks", "All good", None, None), rows)

    def test_admin_can_export_pdf_and_excel(self):
        admin = User.objects.create_superuser(username="admin", password="strong-admin-password")
        submission = Submission.objects.create(site=self.site, template=self.template, submitted_by=admin, checklist_date=date(2026, 11, 19))
        Answer.objects.create(submission=submission, item=self.number_item, value="65", comment="Checked")
        self.client.force_login(admin)
        pdf = self.client.get(f"/admin/submissions/{submission.id}/pdf/")
        xlsx = self.client.get(f"/admin/submissions/{submission.id}/xlsx/")
        self.assertEqual(pdf["Content-Type"], "application/pdf")
        self.assertTrue(pdf.content.startswith(b"%PDF"))
        self.assertEqual(xlsx["Content-Type"], "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        workbook = load_workbook(BytesIO(xlsx.content), data_only=True)
        self.assertEqual(workbook["spv"]["Z7"].value, "65")
        self.assertIn("Submission review", workbook.sheetnames)

    def test_admin_can_edit_user_and_assignment(self):
        admin = User.objects.create_superuser(username="admin", password="strong-admin-password")
        second_site = Site.objects.create(name="East Solar")
        self.client.force_login(admin)
        response = self.client.post(f"/manage/users/{self.user.pk}/edit/", {
            "first_name": "Asha", "last_name": "Operator", "email": "asha@example.com",
            "is_active": "on", "site": str(second_site.pk),
        })
        self.assertRedirects(response, "/manage/users/")
        self.user.refresh_from_db()
        self.assertEqual(self.user.first_name, "Asha")
        self.assertEqual(self.user.email, "asha@example.com")
        self.assertEqual(self.user.site_assignment.site, second_site)
        self.assertEqual(self.user.username, "operator")

    def test_admin_cannot_assign_user_to_an_already_assigned_site(self):
        admin = User.objects.create_superuser(username="admin", password="strong-admin-password")
        other = User.objects.create_user(username="other", password="secure-test-password")
        UserSite.objects.create(user=other, site=Site.objects.create(name="West Solar"))
        self.client.force_login(admin)
        response = self.client.post(f"/manage/users/{self.user.pk}/edit/", {
            "first_name": "", "last_name": "", "email": "", "is_active": "on",
            "site": str(other.site_assignment.site_id),
        })
        self.assertEqual(response.status_code, 200)
        self.assertEqual(self.user.site_assignment.site, self.site)

    def test_admin_can_edit_site_and_add_remove_equipment(self):
        admin = User.objects.create_superuser(username="admin", password="strong-admin-password")
        self.client.force_login(admin)
        response = self.client.post(f"/manage/sites/{self.site.pk}/edit/", {"name": "North Solar", "company": "Sun Co", "location": "Rajasthan", "active": "on"})
        self.assertRedirects(response, f"/manage/sites/{self.site.pk}/edit/")
        add_response = self.client.post(f"/manage/sites/{self.site.pk}/equipment/add/", {"name": "Inverter 1"})
        self.assertRedirects(add_response, f"/manage/sites/{self.site.pk}/edit/")
        equipment = SiteEquipment.objects.get(site=self.site, name="Inverter 1")
        remove_response = self.client.post(f"/manage/sites/{self.site.pk}/equipment/{equipment.pk}/remove/")
        self.assertRedirects(remove_response, f"/manage/sites/{self.site.pk}/edit/")
        self.assertFalse(SiteEquipment.objects.filter(pk=equipment.pk).exists())
        self.site.refresh_from_db()
        self.assertEqual(self.site.company, "Sun Co")

    def test_non_admin_cannot_edit_sites(self):
        self.client.force_login(self.user)
        self.assertEqual(self.client.get(f"/manage/sites/{self.site.pk}/edit/").status_code, 302)

    def test_duplicate_site_day_submission_is_rejected_by_database(self):
        admin = User.objects.create_superuser(username="admin", password="strong-admin-password")
        Submission.objects.create(site=self.site, template=self.template, submitted_by=admin, checklist_date=date(2026, 11, 20))
        with self.assertRaises(Exception):
            Submission.objects.create(site=self.site, template=self.template, submitted_by=admin, checklist_date=date(2026, 11, 20))

    def test_multiple_checklists_sequential_flow_for_operator(self):
        # Create a second template and assign both to self.site
        second_template = ChecklistTemplate.objects.create(name="Inverter Checks", version=1)
        item2 = ChecklistItem.objects.create(template=second_template, checkpoint_no=1, description="Grid Voltage", answer_type="number", position=1)
        SiteChecklist.objects.create(site=self.site, template=self.template, position=1)
        SiteChecklist.objects.create(site=self.site, template=second_template, position=2)

        self.client.force_login(self.user)
        # Step 1: submit the first template
        self.post_preview()
        finalize_resp = self.client.post("/checklist/finalize/", {"confirm_lockin": "yes"})
        # Should redirect to checklist-form to fill the second template
        self.assertRedirects(finalize_resp, "/checklist/")
        self.assertTrue(Submission.objects.filter(site=self.site, template=self.template).exists())
        self.assertFalse(Submission.objects.filter(site=self.site, template=second_template).exists())

        # Step 2: fill second template
        resp2 = self.client.get("/checklist/")
        self.assertEqual(resp2.status_code, 200)
        self.assertContains(resp2, "Inverter Checks")
        self.assertContains(resp2, "Grid Voltage")

        # Submit preview for second template
        post_resp = self.client.post("/checklist/", {f"item_{item2.id}": "415", "remarks": "Inverter ok"})
        self.assertEqual(post_resp.status_code, 200)
        # Lock in second template
        finalize2 = self.client.post("/checklist/finalize/", {"confirm_lockin": "yes"})
        sub2 = Submission.objects.get(site=self.site, template=second_template)
        # Should redirect to review page since all are complete
        self.assertRedirects(finalize2, f"/submissions/{sub2.pk}/review/")

    def test_zero_sequence_checklists_can_be_filled_in_any_order(self):
        second_template = ChecklistTemplate.objects.create(name="Second checklist", version=1)
        second_item = ChecklistItem.objects.create(
            template=second_template, checkpoint_no=1, description="Second check", answer_type="text", position=1
        )
        SiteChecklist.objects.create(site=self.site, template=self.template, position=0)
        SiteChecklist.objects.create(site=self.site, template=second_template, position=0)
        self.client.force_login(self.user)

        dashboard = self.client.get("/")
        self.assertEqual(dashboard.status_code, 200)
        self.assertEqual(sum(item["available"] for item in dashboard.context["checklist_status_list"]), 2)
        self.assertContains(dashboard, "Any order")
        self.assertEqual(dashboard.content.decode().count("Fill now →"), 2)

        # Select and preview the second checklist first, despite its name/order.
        entry = self.client.get(f"/checklist/?template_id={second_template.pk}")
        self.assertEqual(entry.status_code, 200)
        self.assertContains(entry, "Second check")
        preview = self.client.post("/checklist/", {"template_id": second_template.pk, f"item_{second_item.pk}": "ok"})
        self.assertEqual(preview.status_code, 200)
        lockin = self.client.get(f"/checklist/finalize/?template_id={second_template.pk}")
        self.assertEqual(lockin.status_code, 200)
        finalized = self.client.post("/checklist/finalize/", {"template_id": second_template.pk, "confirm_lockin": "yes"})
        self.assertTrue(Submission.objects.filter(site=self.site, template=second_template).exists())
        self.assertFalse(Submission.objects.filter(site=self.site, template=self.template).exists())
        self.assertRedirects(finalized, "/checklist/")

    def test_only_next_positive_sequence_is_available_and_cannot_be_skipped(self):
        second_template = ChecklistTemplate.objects.create(name="Second sequence", version=1)
        SiteChecklist.objects.create(site=self.site, template=self.template, position=1)
        SiteChecklist.objects.create(site=self.site, template=second_template, position=2)
        self.client.force_login(self.user)

        dashboard = self.client.get("/")
        statuses = {item["template"].pk: item for item in dashboard.context["checklist_status_list"]}
        self.assertTrue(statuses[self.template.pk]["available"])
        self.assertFalse(statuses[second_template.pk]["available"])
        self.assertContains(dashboard, "Fill now →")
        self.assertContains(dashboard, "Upcoming")

        blocked = self.client.get(f"/checklist/?template_id={second_template.pk}")
        self.assertRedirects(blocked, "/")
        self.assertFalse(Submission.objects.filter(site=self.site, template=second_template).exists())

    def test_unsequenced_checklists_are_available_alongside_sequence_lane(self):
        optional_template = ChecklistTemplate.objects.create(name="Any-order checklist", version=1)
        sequenced_template = ChecklistTemplate.objects.create(name="Sequenced checklist", version=1)
        later_template = ChecklistTemplate.objects.create(name="Later checklist", version=1)
        SiteChecklist.objects.create(site=self.site, template=optional_template, position=0)
        SiteChecklist.objects.create(site=self.site, template=sequenced_template, position=1)
        SiteChecklist.objects.create(site=self.site, template=later_template, position=2)
        self.client.force_login(self.user)

        dashboard = self.client.get("/")
        statuses = {item["template"].pk: item for item in dashboard.context["checklist_status_list"]}
        self.assertTrue(statuses[optional_template.pk]["available"])
        self.assertTrue(statuses[sequenced_template.pk]["available"])
        self.assertFalse(statuses[later_template.pk]["available"])
        self.assertEqual(dashboard.content.decode().count("Fill now →"), 2)

    def test_previous_ist_day_submissions_do_not_block_new_day(self):
        from unittest.mock import patch

        first_template = ChecklistTemplate.objects.create(name="Daily first", version=1)
        second_template = ChecklistTemplate.objects.create(name="Daily second", version=1)
        SiteChecklist.objects.create(site=self.site, template=first_template, position=0)
        SiteChecklist.objects.create(site=self.site, template=second_template, position=0)
        previous_day = date(2026, 11, 20)
        today = date(2026, 11, 21)
        Submission.objects.create(
            site=self.site, template=first_template, submitted_by=self.user, checklist_date=previous_day
        )
        Submission.objects.create(
            site=self.site, template=second_template, submitted_by=self.user, checklist_date=previous_day
        )
        self.client.force_login(self.user)

        # At the next IST local date, prior-day submissions no longer mark either checklist complete.
        with patch("checklist.views.timezone.localdate", return_value=today), patch(
            "checklist.views.GlobalSettings.is_open", return_value=True
        ):
            dashboard = self.client.get("/")
            self.assertEqual(dashboard.status_code, 200)
            statuses = {item["template"].pk: item for item in dashboard.context["checklist_status_list"]}
            self.assertTrue(statuses[first_template.pk]["available"])
            self.assertTrue(statuses[second_template.pk]["available"])
            self.assertIsNone(statuses[first_template.pk]["submission"])
            self.assertIsNone(statuses[second_template.pk]["submission"])

            entry = self.client.get(f"/checklist/?template_id={second_template.pk}")
            self.assertEqual(entry.status_code, 200)
            self.assertContains(entry, second_template.name)

            # The database permits a new submission for the same checklist on the new date.
            new_day_submission = Submission.objects.create(
                site=self.site, template=second_template, submitted_by=self.user, checklist_date=today
            )
            self.assertNotEqual(new_day_submission.checklist_date, previous_day)


    def test_admin_upload_excel_checklist_template(self):
        admin = User.objects.create_superuser(username="admin", password="strong-admin-password")
        self.client.force_login(admin)

        wb = Workbook()
        ws = wb.active
        ws.title = "Checks"
        ws.append(["Checkpoint", "Description", "Standard", "Method", "Frequency"])
        ws.append([1, "HT Breaker Trip Test", "No trip failure", "Visual", "Daily"])
        ws.append([2, "HT Busbar Temperature", "Below 55 C", "IR Gun", "Daily"])
        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        upload_file = SimpleUploadedFile("ht_panel.xlsx", buf.read(), content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        response = self.client.post("/manage/checklists/", {
            "template_name": "HT Panel Daily Inspection",
            "file": upload_file,
            "assign_to_sites": [self.site.pk],
        })
        new_template = ChecklistTemplate.objects.get(name="HT Panel Daily Inspection")
        self.assertRedirects(response, f"/manage/checklists/{new_template.pk}/preview/")
        self.assertEqual(new_template.items.count(), 2)
        item = new_template.items.get(checkpoint_no=2)
        self.assertEqual(item.answer_type, "number")
        self.assertEqual(item.unit, "°C")
        self.assertTrue(SiteChecklist.objects.filter(site=self.site, template=new_template).exists())

    def test_admin_preview_checklist_template(self):
        admin = User.objects.create_superuser(username="admin", password="strong-admin-password")
        self.client.force_login(admin)

        # Admin can view preview
        resp = self.client.get(f"/manage/checklists/{self.template.pk}/preview/")
        self.assertEqual(resp.status_code, 200)
        self.assertContains(resp, self.template.name)
        self.assertContains(resp, "Winding temperature")
        self.assertContains(resp, "Yes / No")
        self.assertContains(resp, "Number")
        self.assertContains(resp, "TOTAL CHECKPOINTS")

        # Operator (non-staff) cannot view preview
        self.client.force_login(self.user)
        op_resp = self.client.get(f"/manage/checklists/{self.template.pk}/preview/")
        self.assertNotEqual(op_resp.status_code, 200)

    def test_import_uploaded_excel_multiline_headers_and_sections(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "SolarInverter"
        ws.append(["", "", "", "", "", ""])
        ws.append(["", "Section", "Check\nPoint No.", "Job Description", "Standard", "Method/\nTool", "Frequency"])
        ws.append(["", "AC Cabinet", 1, "AC Breaker Status", "Closed", "Visual", "Daily"])
        ws.append(["", "", 2, "Inverter Temperature", "Below 70 C", "Digital Display", "Daily"])
        ws.append(["", "DC Combiner", 3, "DC String Fuse Check", "Intact", "Multimeter", "Daily"])
        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)

        from checklist.services.workbook_import import import_uploaded_excel
        template, count = import_uploaded_excel(buf, "Solar Inverter Station")
        self.assertEqual(count, 3)
        self.assertEqual(template.items.count(), 3)
        item1 = template.items.get(checkpoint_no=1)
        self.assertEqual(item1.section, "AC Cabinet")
        self.assertEqual(item1.description, "AC Breaker Status")
        item2 = template.items.get(checkpoint_no=2)
        self.assertEqual(item2.section, "AC Cabinet")  # inherited section
        self.assertEqual(item2.answer_type, "number")
        item3 = template.items.get(checkpoint_no=3)
        self.assertEqual(item3.section, "DC Combiner")

    def test_admin_assign_and_remove_site_checklist(self):
        admin = User.objects.create_superuser(username="admin", password="strong-admin-password")
        self.client.force_login(admin)
        new_template = ChecklistTemplate.objects.create(name="Battery Bank Check", version=1)

        # Assign checklist
        assign_resp = self.client.post(f"/manage/sites/{self.site.pk}/checklists/assign/", {
            "template": str(new_template.pk),
            "position": "3",
            "active": "on",
        })
        self.assertRedirects(assign_resp, f"/manage/sites/{self.site.pk}/edit/")
        sc = SiteChecklist.objects.get(site=self.site, template=new_template)
        self.assertEqual(sc.position, 3)

        # Remove checklist
        remove_resp = self.client.post(f"/manage/sites/{self.site.pk}/checklists/{sc.pk}/remove/")
        self.assertRedirects(remove_resp, f"/manage/sites/{self.site.pk}/edit/")
        self.assertFalse(SiteChecklist.objects.filter(pk=sc.pk).exists())

    def test_admin_site_status_inline_review_tab(self):
        admin = User.objects.create_superuser(username="admin", password="strong-admin-password")
        sub = Submission.objects.create(site=self.site, template=self.template, submitted_by=self.user, checklist_date=date(2026, 11, 20), remarks="Routine checks fine")
        Answer.objects.create(submission=sub, item=self.number_item, value="70.5", comment="Stable")
        self.client.force_login(admin)

        # Access inline review tab for same day or date range
        response = self.client.get(f"/?tab=review&site={self.site.pk}&start_date=2026-11-20&end_date=2026-11-20")
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "North Solar")
        self.assertContains(response, "Test Trafo")
        self.assertContains(response, "70.5")
        self.assertContains(response, "Routine checks fine")
        self.assertContains(response, "Stable")

    def test_admin_credential_change_with_valid_current_password(self):
        admin = User.objects.create_superuser(username="admin", password="strong-admin-password")
        self.client.force_login(admin)

        response = self.client.post("/admin-settings/", {
            "change_credentials": "1",
            "new_username": "site_admin",
            "new_password": "new-strong-password",
            "confirm_password": "new-strong-password",
            "current_password": "strong-admin-password",
        })
        self.assertRedirects(response, "/admin-settings/")
        admin.refresh_from_db()
        self.assertEqual(admin.username, "site_admin")
        self.assertTrue(admin.check_password("new-strong-password"))

    def test_admin_credential_change_fails_with_invalid_current_password(self):
        admin = User.objects.create_superuser(username="admin", password="strong-admin-password")
        self.client.force_login(admin)

        response = self.client.post("/admin-settings/", {
            "change_credentials": "1",
            "new_username": "hacked_admin",
            "new_password": "hacked-password",
            "confirm_password": "hacked-password",
            "current_password": "wrong-password",
        })
        self.assertEqual(response.status_code, 200)
        admin.refresh_from_db()
        self.assertEqual(admin.username, "admin")
        self.assertFalse(admin.check_password("hacked-password"))

    def test_admin_can_update_username_only_with_current_password(self):
        admin = User.objects.create_superuser(username="admin", password="strong-admin-password")
        self.client.force_login(admin)

        response = self.client.post("/admin-settings/", {
            "change_credentials": "1",
            "new_username": "new_admin_user",
            "new_password": "",
            "confirm_password": "",
            "current_password": "strong-admin-password",
        })
        self.assertRedirects(response, "/admin-settings/")
        admin.refresh_from_db()
        self.assertEqual(admin.username, "new_admin_user")
        self.assertTrue(admin.check_password("strong-admin-password"))


class ElevenKVAndMonthlyExportTests(TestCase):
    SAMPLE_11KV_CSV = """\
,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
,,Document No. TEIPL/11KVEQUIPMENT,,,,,"Talf Energy India Pvt Ltd, Jhatikara DL9",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
,,,,,,,Equipment Name : 11 KV Equipment Control Panel & LT Panel,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
Section,,"Check 
point no.",Job Description,Standard,"Method/
Tool",Frequency,Month-,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
,,,,,,,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,31,31
Control room,HT panel(11Kv),1,Check current in all phases on 11kv HT panel,It should be balanced in all the phases,,D,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
,,2,Check voltage between all phases on 11kv HT Panel( RY/YB/BR),It should be balanced in all phases ( value notedown ),,D,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
,,3, Check annunciation on 11kv HT panel,Must be healthy,,D,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
,,4,Breaker healthy indication on 11kv HT panel,Must be healthy,,D,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
,,5, Check DC/AC supply status on 11kv HT panel,Must be healthy,,D,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
,,6,Check VCB spring charge indication on 11kv HT panel,Must be healthy,,D,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
,,7,Check MFM reading on 11kv HT panel,Reading must be visual/ appropriate,,D,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
,LT panel(800 Vac),8,Check AC breaker (MCCB) status,Must be healthy,,D,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
,,9,Check AC supply indication lamp,Must be glowing,,D,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
,,10,Check AC SPD status,Must be healthy,,D,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
,,11,Check panel physical healthness condition,Door handle/door look must be in working condition and rustfree.,,D,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
,,12,Check panel earthing connection,Confirm with match marking,,D,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
,,13,Panel surrounding area,Should be clean,,D,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
Legends,,,,Signature,Checked & done by,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
 IF OK :                   ,,,,,Verify by,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
 IF NOT OK : ,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
 D - Daily,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
 W - Weekly,,,,Note : If Found abnormal inform Maintenance Department.,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
"""

    def setUp(self):
        self.site = Site.objects.create(name="Jhatikara Site", company="Talf Energy India Pvt Ltd, Jhatikara DL9")
        self.admin = User.objects.create_superuser(username="admin", password="strong-admin-password")
        self.operator = User.objects.create_user(username="op1", password="secure-test-password")
        UserSite.objects.create(user=self.operator, site=self.site)

    def test_upload_11kv_csv_parses_metadata_subsections_and_answer_types(self):
        self.client.force_login(self.admin)
        upload = SimpleUploadedFile("11kv_equipment.csv", self.SAMPLE_11KV_CSV.encode("utf-8"), content_type="text/csv")
        response = self.client.post("/manage/checklists/", {
            "template_name": "11 KV Equipment Control Panel",
            "file": upload,
            "assign_to_sites": [self.site.pk],
        })
        template = ChecklistTemplate.objects.get(name="11 KV Equipment Control Panel")
        self.assertRedirects(response, f"/manage/checklists/{template.pk}/preview/")
        self.assertEqual(template.document_number, "TEIPL/11KVEQUIPMENT")
        self.assertEqual(template.company, "Talf Energy India Pvt Ltd, Jhatikara DL9")
        self.assertEqual(template.equipment_name, "11 KV Equipment Control Panel & LT Panel")
        self.assertEqual(template.items.count(), 13)

        cp1 = template.items.get(checkpoint_no=1)
        cp2 = template.items.get(checkpoint_no=2)
        cp8 = template.items.get(checkpoint_no=8)
        cp13 = template.items.get(checkpoint_no=13)
        self.assertEqual(cp1.section, "Control room — HT panel(11Kv)")
        self.assertEqual(cp1.answer_type, ChecklistItem.AnswerType.BOOLEAN)
        self.assertEqual(cp1.frequency, "Daily")
        self.assertEqual(cp2.section, "Control room — HT panel(11Kv)")
        self.assertEqual(cp2.answer_type, ChecklistItem.AnswerType.TEXT)
        self.assertEqual(cp8.section, "Control room — LT panel(800 Vac)")
        self.assertEqual(cp8.answer_type, ChecklistItem.AnswerType.BOOLEAN)
        self.assertEqual(cp13.answer_type, ChecklistItem.AnswerType.BOOLEAN)

    def test_monthly_excel_and_pdf_export_combines_all_days_in_month_matrix(self):
        from checklist.services.workbook_import import import_uploaded_csv

        template, count = import_uploaded_csv(
            SimpleUploadedFile("11kv.csv", self.SAMPLE_11KV_CSV.encode("utf-8")),
            "11 KV Equipment",
        )
        self.assertEqual(count, 13)
        SiteChecklist.objects.create(site=self.site, template=template, position=1)

        items = list(template.items.order_by("checkpoint_no"))
        for day_num, voltage_val, status_val in [(1, "11.1/11.0/11.1", "YES"), (15, "11.0/11.0/11.0", "YES"), (30, "10.9/11.0/11.1", "NO")]:
            sub = Submission.objects.create(
                site=self.site,
                template=template,
                submitted_by=self.operator,
                checklist_date=date(2026, 11, day_num),
                remarks=f"Day {day_num} check",
                signer_name="Operator One",
            )
            for item in items:
                val = voltage_val if item.checkpoint_no == 2 else (status_val if item.checkpoint_no == 1 else "YES")
                comment = "Phase imbalance noticed" if (day_num == 30 and item.checkpoint_no == 1) else ""
                Answer.objects.create(submission=sub, item=item, value=val, comment=comment)

        self.client.force_login(self.admin)
        xlsx_resp = self.client.get(f"/admin/export/monthly/xlsx/?site={self.site.pk}&month=2026-11")
        self.assertEqual(xlsx_resp.status_code, 200)
        self.assertEqual(xlsx_resp["Content-Type"], "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        wb = load_workbook(BytesIO(xlsx_resp.content), data_only=True)
        self.assertIn("11 KV Equipment", wb.sheetnames)
        self.assertIn("Comments & Remarks", wb.sheetnames)

        ws = wb["11 KV Equipment"]
        # Row 5 has day numbers starting at col 8 (H) because 7 metadata cols exist (Section + Sub-section + 5)
        self.assertEqual(ws.cell(5, 8).value, 1)
        self.assertEqual(ws.cell(5, 22).value, 15)
        self.assertEqual(ws.cell(5, 37).value, 30)
        # Row 6 = CP 1, Row 7 = CP 2
        self.assertEqual(ws.cell(6, 8).value, "YES")
        self.assertEqual(ws.cell(6, 37).value, "NO")
        self.assertEqual(ws.cell(7, 8).value, "11.1/11.0/11.1")
        self.assertEqual(ws.cell(7, 22).value, "11.0/11.0/11.0")
        self.assertEqual(ws.cell(7, 37).value, "10.9/11.0/11.1")

        pdf_resp = self.client.get(f"/admin/export/monthly/pdf/?site={self.site.pk}&month=2026-11")
        self.assertEqual(pdf_resp.status_code, 200)
        self.assertEqual(pdf_resp["Content-Type"], "application/pdf")
        self.assertTrue(pdf_resp.content.startswith(b"%PDF"))

    def test_operator_can_export_own_site_monthly_but_not_other_site(self):
        other_site = Site.objects.create(name="Other Site")
        template = ChecklistTemplate.objects.create(name="11 KV Test", version=1)
        item = ChecklistItem.objects.create(template=template, checkpoint_no=1, description="Status", answer_type="boolean", position=1)
        sub = Submission.objects.create(site=self.site, template=template, submitted_by=self.operator, checklist_date=date(2026, 11, 5))
        Answer.objects.create(submission=sub, item=item, value="YES")

        self.client.force_login(self.operator)
        ok_xlsx = self.client.get("/history/export/monthly/xlsx/?month=2026-11")
        self.assertEqual(ok_xlsx.status_code, 200)
        ok_pdf = self.client.get("/history/export/monthly/pdf/?month=2026-11")
        self.assertEqual(ok_pdf.status_code, 200)
        forbidden = self.client.get(f"/history/export/monthly/xlsx/?site={other_site.pk}&month=2026-11")
        self.assertEqual(forbidden.status_code, 403)
        admin_only = self.client.get(f"/admin/export/monthly/xlsx/?site={self.site.pk}&month=2026-11")
        self.assertEqual(admin_only.status_code, 302)

    def test_empty_monthly_export_redirects_with_warning(self):
        self.client.force_login(self.admin)
        resp = self.client.get(f"/admin/export/monthly/xlsx/?site={self.site.pk}&month=2026-01")
        self.assertEqual(resp.status_code, 302)


class WmsChecklistDynamicDropdownTests(TestCase):
    SAMPLE_WMS_CSV = """\
,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
,Document No. WMS/CHECKLIST,,"Talf Energy India Pvt Ltd, Jhatikara DL9",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
,,,,,"Equipment Name:  WMS, SCADA & CCTV ",,,,,,,,,,,,,,,,,,,,,,Approved By :  ,,,,,,,,
Section,"Check 
point no.",Job Description,Standard,"Method/
Tool",Month-,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
,,,,,1.0,2.0,3.0,4.0,5.0,6.0,7.0,8.0,9.0,10.0,11.0,12.0,13.0,14.0,15.0,16.0,17.0,18.0,19.0,20.0,21.0,22.0,23.0,24.0,25.0,26.0,27.0,28.0,29.0,30.0,31.0
WMS,1.0,Cleaning of Pyranometer ,It should be clean,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
,2.0,"Visual inspection of Pyranometer,Temperature sensor,SCADA system",no damage and dirt found,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
Cleaning of  Module,3.0,Cleaning of Module,It should be clean ,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
,4.0, Module   cleaning cycle ,clean date (check ok only when cleaning done),,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
SCADA System,5.0,Inverter AC/DC KW Communication,Values must be visible on SCADA Portal,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
,6.0,WMS to SCADA portal data communication,No sign of communication Failure,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
,7.0,All Data Report Should be recorded properly,No sign of communication Failure/ Malfunctioning,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
,8.0,Alarm/Event/Trend data recording of inverter,No sign of communication Failure/ Malfunctioning,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
Battary Bank,9.0,Visual inspection of Battery ,"No any sign of any leakage, cracks, corrosion",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
CCTV Camera,10.0,visual inspection of camera display screen,should be all CCTV camera funtional ok,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
,11.0,inspection of storage data of all CCTV cameras,Data must be store in DVR Storage drive,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
Auxiliary supply panel,12.0,Check auxiliary supply panel,should be in working condtion no any sign of failure,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
,13.0,Check supply to PV Area lighting pole,should be all light of PV area in working condition,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
,14.0,check indication lamp ,should be Glowing ,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
Remarks:,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
Legends:,,,Signature,check & done by,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
 IF OK :                   ,,,,verify by,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
 IF NOT OK : ,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
 D - Daily,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
 W - Weekly,,,Note : If Found abnormal inform Maintenance Department.,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
"""

    def setUp(self):
        self.site = Site.objects.create(name="Jhatikara Site", company="Talf Energy India Pvt Ltd, Jhatikara DL9")
        self.admin = User.objects.create_superuser(username="admin", password="strong-admin-password")
        self.operator = User.objects.create_user(username="op_wms", password="secure-test-password")
        UserSite.objects.create(user=self.operator, site=self.site)
        settings = GlobalSettings.get_solo()
        settings.deadline = time(23, 59)
        settings.save()

    def _build_wms_workbook_bytes(self):
        import csv
        import io
        from openpyxl.worksheet.datavalidation import DataValidation

        wb = Workbook()
        ws = wb.active
        ws.title = "WMS"
        for row in csv.reader(io.StringIO(self.SAMPLE_WMS_CSV)):
            ws.append(row)

        lists_ws = wb.create_sheet("Lists")
        lists_ws["A1"] = "Operational"
        lists_ws["A2"] = "Non Operational"
        lists_ws["A3"] = "Spare"

        # Default dropdown on all day cells F7:AJ20
        dv_default = DataValidation(type="list", formula1='"OK,Not OK"', allow_blank=True)
        dv_default.sqref = "F7:AJ20"
        ws.add_data_validation(dv_default)

        # Row-specific override on CP 4 (row 10): cleaning cycle dropdown
        dv_cycle = DataValidation(type="list", formula1='"Done,Pending,N/A"', allow_blank=True)
        dv_cycle.sqref = "F10:AJ10"
        ws.add_data_validation(dv_cycle)

        # Cross-sheet range reference on CP 9 (row 15): Battery bank status
        dv_battery = DataValidation(type="list", formula1="=Lists!$A$1:$A$3", allow_blank=True)
        dv_battery.sqref = "F15:AJ15"
        ws.add_data_validation(dv_battery)

        buf = BytesIO()
        wb.save(buf)
        return buf.getvalue()

    def test_wms_csv_import_handles_temperature_sensor_and_all_14_items(self):
        self.client.force_login(self.admin)
        upload = SimpleUploadedFile("wms.csv", self.SAMPLE_WMS_CSV.encode("utf-8"), content_type="text/csv")
        resp = self.client.post("/manage/checklists/", {
            "template_name": "WMS Checklist",
            "file": upload,
            "assign_to_sites": [self.site.pk],
        })
        tmpl = ChecklistTemplate.objects.get(name="WMS Checklist")
        self.assertRedirects(resp, f"/manage/checklists/{tmpl.pk}/preview/")
        self.assertEqual(tmpl.document_number, "WMS/CHECKLIST")
        self.assertEqual(tmpl.equipment_name, "WMS, SCADA & CCTV")
        self.assertEqual(tmpl.items.count(), 14)

        cp2 = tmpl.items.get(checkpoint_no=2)
        self.assertEqual(cp2.answer_type, ChecklistItem.AnswerType.BOOLEAN)
        self.assertEqual(cp2.unit, "")

    def test_wms_excel_detects_cell_and_cross_sheet_dropdowns_and_exports_daily_and_monthly(self):
        self.client.force_login(self.admin)
        xlsx_bytes = self._build_wms_workbook_bytes()
        upload = SimpleUploadedFile(
            "wms.xlsx",
            xlsx_bytes,
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        self.client.post("/manage/checklists/", {
            "template_name": "WMS Checklist Excel",
            "file": upload,
            "assign_to_sites": [self.site.pk],
        })
        tmpl = ChecklistTemplate.objects.get(name="WMS Checklist Excel")
        self.assertEqual(tmpl.items.count(), 14)

        cp1 = tmpl.items.get(checkpoint_no=1)
        cp2 = tmpl.items.get(checkpoint_no=2)
        cp4 = tmpl.items.get(checkpoint_no=4)
        cp9 = tmpl.items.get(checkpoint_no=9)
        self.assertEqual(cp1.answer_type, ChecklistItem.AnswerType.CHOICES)
        self.assertEqual(cp1.get_choices_list(), ["OK", "Not OK"])
        self.assertEqual(cp2.answer_type, ChecklistItem.AnswerType.CHOICES)
        self.assertEqual(cp4.get_choices_list(), ["Done", "Pending", "N/A"])
        self.assertEqual(cp9.get_choices_list(), ["Operational", "Non Operational", "Spare"])

        # Operator sees dynamic dropdown options and submits daily checklist
        self.client.force_login(self.operator)
        form_resp = self.client.get(f"/checklist/?template_id={tmpl.pk}")
        self.assertEqual(form_resp.status_code, 200)
        self.assertContains(form_resp, '<option value="Done"')
        self.assertContains(form_resp, '<option value="Operational"')
        self.assertContains(form_resp, "WMS/CHECKLIST")

        post_data = {"template_id": tmpl.pk, "remarks": "All WMS systems checked"}
        for it in tmpl.items.all():
            if it.checkpoint_no == 4:
                post_data[f"item_{it.id}"] = "Done"
            elif it.checkpoint_no == 9:
                post_data[f"item_{it.id}"] = "Operational"
            else:
                post_data[f"item_{it.id}"] = "OK"

        preview_resp = self.client.post("/checklist/", post_data)
        self.assertEqual(preview_resp.status_code, 200)
        lockin_resp = self.client.post("/checklist/finalize/", {"template_id": tmpl.pk, "confirm_lockin": "yes"})
        sub = Submission.objects.get(site=self.site, template=tmpl)
        self.assertRedirects(lockin_resp, f"/submissions/{sub.pk}/review/")

        # Daily Excel + PDF exports
        self.client.force_login(self.admin)
        daily_xlsx = self.client.get(f"/admin/submissions/{sub.pk}/xlsx/")
        self.assertEqual(daily_xlsx.status_code, 200)
        dwb = load_workbook(BytesIO(daily_xlsx.content), data_only=False)
        dws = dwb["WMS Checklist Excel"]
        self.assertGreaterEqual(len(dws.data_validations.dataValidation), 14)

        daily_pdf = self.client.get(f"/admin/submissions/{sub.pk}/pdf/")
        self.assertEqual(daily_pdf.status_code, 200)
        self.assertTrue(daily_pdf.content.startswith(b"%PDF"))

        # Monthly Excel + PDF exports
        month_str = sub.checklist_date.strftime("%Y-%m")
        monthly_xlsx = self.client.get(f"/admin/export/monthly/xlsx/?site={self.site.pk}&month={month_str}")
        self.assertEqual(monthly_xlsx.status_code, 200)
        mwb = load_workbook(BytesIO(monthly_xlsx.content), data_only=False)
        mws = mwb["WMS Checklist Excel"]
        day_col = 6 + sub.checklist_date.day  # 6 meta cols + day (1-indexed)
        self.assertEqual(mws.cell(6, day_col).value, "OK")
        self.assertEqual(mws.cell(9, day_col).value, "Done")
        self.assertEqual(mws.cell(14, day_col).value, "Operational")
        self.assertGreaterEqual(len(mws.data_validations.dataValidation), 14)

        monthly_pdf = self.client.get(f"/admin/export/monthly/pdf/?site={self.site.pk}&month={month_str}")
        self.assertEqual(monthly_pdf.status_code, 200)
        self.assertTrue(monthly_pdf.content.startswith(b"%PDF"))

    def test_upload_legacy_xls_and_misnamed_csv_do_not_raise_badzipfile(self):
        import subprocess
        import tempfile
        from pathlib import Path

        self.client.force_login(self.admin)

        # 1. CSV content uploaded with .xlsx extension (common cause of BadZipFile)
        misnamed = SimpleUploadedFile(
            "wms_misnamed.xlsx",
            self.SAMPLE_WMS_CSV.encode("utf-8"),
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        resp_csv = self.client.post("/manage/checklists/", {
            "template_name": "WMS Misnamed CSV",
            "file": misnamed,
        })
        tmpl_csv = ChecklistTemplate.objects.get(name="WMS Misnamed CSV")
        self.assertRedirects(resp_csv, f"/manage/checklists/{tmpl_csv.pk}/preview/")
        self.assertEqual(tmpl_csv.items.count(), 14)

        # 2. Genuine legacy binary .xls file (OLE2) converted via LibreOffice with dropdowns intact
        xlsx_bytes = self._build_wms_workbook_bytes()
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp)
            src = p / "wms.xlsx"
            src.write_bytes(xlsx_bytes)
            subprocess.run(
                ["soffice", "--headless", "--norestore", "--convert-to", "xls", "--outdir", str(p), str(src)],
                check=True,
                capture_output=True,
                timeout=30,
            )
            xls_bytes = (p / "wms.xls").read_bytes()

        xls_upload = SimpleUploadedFile("wms_legacy.xls", xls_bytes, content_type="application/vnd.ms-excel")
        resp_xls = self.client.post("/manage/checklists/", {
            "template_name": "WMS Legacy XLS",
            "file": xls_upload,
        })
        tmpl_xls = ChecklistTemplate.objects.get(name="WMS Legacy XLS")
        self.assertRedirects(resp_xls, f"/manage/checklists/{tmpl_xls.pk}/preview/")
        self.assertEqual(tmpl_xls.items.count(), 14)
        self.assertEqual(tmpl_xls.items.get(checkpoint_no=4).get_choices_list(), ["Done", "Pending", "N/A"])


class InverterStringSummaryTests(TestCase):
    def setUp(self):
        self.site = Site.objects.create(name="dhansa")
        self.admin = User.objects.create_superuser(username="admin", password="strong-admin-password")
        self.operator = User.objects.create_user(username="op_dhansa", password="secure-test-password")
        UserSite.objects.create(user=self.operator, site=self.site)
        self.inv3 = ChecklistTemplate.objects.create(name="inv3", version=1)
        self.items = [
            ChecklistItem.objects.create(
                template=self.inv3,
                section="Inverter 1",
                checkpoint_no=i,
                description=f"PV{i} - Status Check",
                answer_type=ChecklistItem.AnswerType.CHOICES,
                choices="Operational, Non-Operational, Spare",
                position=i,
            )
            for i in range(1, 7)
        ]

    def test_non_operational_string_count_for_day_inverter_and_site(self):
        target_day = date(2026, 11, 25)
        sub = Submission.objects.create(
            site=self.site,
            template=self.inv3,
            submitted_by=self.operator,
            checklist_date=target_day,
        )
        # PV1, PV4 -> Operational (2); PV3, PV6 -> Non-Operational (2); PV2, PV5 -> Spare (2)
        vals = ["Operational", "Spare", "Non-Operational", "Operational", "Spare", "Non-Operational"]
        for item, val in zip(self.items, vals):
            Answer.objects.create(submission=sub, item=item, value=val)

        # Submission review (operator/admin)
        self.client.force_login(self.operator)
        review_resp = self.client.get(f"/submissions/{sub.pk}/review/")
        self.assertEqual(review_resp.status_code, 200)
        self.assertContains(review_resp, "Non-Operational (2):")
        self.assertContains(review_resp, "PV3, PV6")

        # History
        hist_resp = self.client.get("/history/")
        self.assertEqual(hist_resp.status_code, 200)
        self.assertContains(hist_resp, "2 Non-Op / 6")
        self.assertContains(hist_resp, "2 Spare")

        # Admin Review tab filtered by site + inverter + day
        self.client.force_login(self.admin)
        dash_review = self.client.get(
            f"/?tab=review&site={self.site.pk}&template={self.inv3.pk}&start_date=2026-11-25&end_date=2026-11-25"
        )
        self.assertEqual(dash_review.status_code, 200)
        self.assertContains(dash_review, "Inverter / String Status Summary")
        self.assertContains(dash_review, "Total Non-Operational: 2")
        self.assertContains(dash_review, "2 / 6")
        self.assertContains(dash_review, "PV3, PV6")

        # Monthly Excel includes dynamic daily count rows
        xlsx_resp = self.client.get(f"/admin/export/monthly/xlsx/?site={self.site.pk}&template={self.inv3.pk}&month=2026-11")
        self.assertEqual(xlsx_resp.status_code, 200)
        wb = load_workbook(BytesIO(xlsx_resp.content), data_only=True)
        ws = wb["inv3"]
        day_col = 6 + 25  # 6 meta cols + day 25
        # Rows 6..11 = 6 PV items; Rows 12, 13, 14 = Operational Count, Non-Operational Count, Spare Count
        self.assertEqual(ws.cell(12, 3).value, "Operational Count")
        self.assertEqual(ws.cell(12, day_col).value, 2)
        self.assertEqual(ws.cell(13, 3).value, "Non-Operational Count")
        self.assertEqual(ws.cell(13, day_col).value, 2)
        self.assertEqual(ws.cell(14, 3).value, "Spare Count")
        self.assertEqual(ws.cell(14, day_col).value, 2)

    def test_executive_summary_shows_per_site_daily_non_operational_assets_in_red(self):
        other_site = Site.objects.create(name="Jhatikara")
        sub_day1 = Submission.objects.create(
            site=self.site,
            template=self.inv3,
            submitted_by=self.operator,
            checklist_date=date(2026, 11, 24),
        )
        for item, val in zip(self.items, ["Operational"] * 6):
            Answer.objects.create(submission=sub_day1, item=item, value=val)

        sub_day2 = Submission.objects.create(
            site=self.site,
            template=self.inv3,
            submitted_by=self.operator,
            checklist_date=date(2026, 11, 25),
        )
        for item, val in zip(
            self.items,
            ["Operational", "Spare", "Non-Operational", "Operational", "Spare", "Non-Operational"],
        ):
            Answer.objects.create(submission=sub_day2, item=item, value=val, comment="Faulty string" if val == "Non-Operational" else "")

        self.client.force_login(self.admin)

        # 1. Dedicated top-level /executive-summary/ page across all sites and days
        exec_resp = self.client.get("/executive-summary/?start_date=2026-11-24&end_date=2026-11-25")
        self.assertEqual(exec_resp.status_code, 200)
        self.assertContains(exec_resp, "dhansa")
        self.assertContains(exec_resp, "Jhatikara")
        self.assertContains(exec_resp, "2 Non-Operational Assets")
        self.assertContains(exec_resp, "PV3")
        self.assertContains(exec_resp, "PV6")
        self.assertContains(exec_resp, "#dc2626")  # red highlight
        self.assertContains(exec_resp, "All submitted assets operational")  # day 24 clean

        # 2. Dashboard stays clean (no Executive Summary section rendered on Dashboard) and top nav shows link
        with patch("checklist.views.timezone.localdate", return_value=date(2026, 11, 25)):
            status_resp = self.client.get("/")
        self.assertEqual(status_resp.status_code, 200)
        self.assertContains(status_resp, 'href="/executive-summary/"')
        self.assertNotContains(status_resp, "Executive Summary — Non-Operational Assets by Site")

        # 3. Non-staff operator & anonymous users cannot see or access /executive-summary/
        self.client.force_login(self.operator)
        op_dash = self.client.get("/")
        self.assertEqual(op_dash.status_code, 200)
        self.assertNotContains(op_dash, "Executive Summary")
        self.assertNotContains(op_dash, "/executive-summary/")
        op_hist = self.client.get("/history/")
        self.assertEqual(op_hist.status_code, 200)
        self.assertNotContains(op_hist, "/executive-summary/")
        op_resp = self.client.get("/executive-summary/")
        self.assertEqual(op_resp.status_code, 302)
        self.client.logout()
        anon_resp = self.client.get("/executive-summary/")
        self.assertEqual(anon_resp.status_code, 302)
        self.client.force_login(self.admin)

        # 4. Edit site page includes per-site Executive Summary
        site_resp = self.client.get(f"/manage/sites/{self.site.pk}/edit/")
        self.assertEqual(site_resp.status_code, 200)
        self.assertContains(site_resp, "Executive Summary — Daily Non-Operational Assets (dhansa)")
        self.assertContains(site_resp, "PV3")
        self.assertContains(site_resp, "PV6")








