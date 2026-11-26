import re
from datetime import time

from django.conf import settings
from django.core.exceptions import ValidationError
from django.db import models
from django.utils import timezone


class Site(models.Model):
    name = models.CharField(max_length=160, unique=True)
    company = models.CharField(max_length=200, blank=True)
    location = models.CharField(max_length=200, blank=True)
    active = models.BooleanField(default=True)

    def __str__(self):
        return self.name


class SiteEquipment(models.Model):
    site = models.ForeignKey(Site, on_delete=models.CASCADE, related_name="equipment")
    name = models.CharField(max_length=200)

    class Meta:
        ordering = ["name"]
        constraints = [models.UniqueConstraint(fields=["site", "name"], name="unique_equipment_per_site")]

    def __str__(self):
        return self.name


class UserSite(models.Model):
    user = models.OneToOneField(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, related_name="site_assignment")
    site = models.OneToOneField(Site, on_delete=models.CASCADE, related_name="user_assignment")

    def __str__(self):
        return f"{self.user} → {self.site}"


CATEGORY_CHOICES = [
    ("Inverters", "Inverters"),
    ("Transformers", "Transformers"),
    ("WMS & SCADA", "WMS & SCADA"),
    ("HT / LT Panels", "HT / LT Panels"),
    ("SMB / Combiner", "SMB / Combiner"),
    ("General", "General"),
]

CATEGORY_ORDER = {label: idx for idx, (label, _) in enumerate(CATEGORY_CHOICES)}


def infer_checklist_category(name="", equipment_name="", document_number=""):
    """Infer an equipment category group from template metadata."""
    text = f"{name} {equipment_name} {document_number}".lower()
    if re.search(r"\binv(?:erter)?\s*\d*\b", text) or "string monitoring" in text:
        return "Inverters"
    if "trafo" in text or "transformer" in text:
        return "Transformers"
    if any(tok in text for tok in ("wms", "scada", "cctv", "weather", "pyranometer")):
        return "WMS & SCADA"
    if any(tok in text for tok in ("11kv", "11 kv", "ht panel", "lt panel", "control panel", "switchgear")):
        return "HT / LT Panels"
    if any(tok in text for tok in ("smb", "combiner", "scb")):
        return "SMB / Combiner"
    return "General"


def resolve_equipment_label(site, template, site_category=None):
    """
    Resolve the equipment name specific to a checklist template on a site
    so Inverter sheets show the inverter, Transformer sheets show the transformer,
    and unrelated site equipment is not mixed in.
    """
    if template and (template.equipment_name or "").strip():
        return template.equipment_name.strip()

    tmpl_name = (template.name if template else "") or ""
    if site_category:
        tmpl_cat = site_category.strip()
    elif template and site is not None:
        sc_cat = (
            SiteChecklist.objects.filter(site=site, template=template)
            .values_list("category", flat=True)
            .first()
        )
        tmpl_cat = (sc_cat or "").strip() or template.resolved_category()
    elif template:
        tmpl_cat = template.resolved_category()
    else:
        tmpl_cat = ""

    if site is not None:
        all_equip = list(site.equipment.values_list("name", flat=True))
        if all_equip:
            m_inv = re.search(r"\binv(?:erter)?[\s_-]*(\d+)\b", tmpl_name, re.IGNORECASE)
            if m_inv:
                num = m_inv.group(1)
                numbered_matches = [
                    eq for eq in all_equip
                    if re.search(rf"\binv(?:erter)?[\s_-]*0*{re.escape(num)}\b", eq, re.IGNORECASE)
                ]
                if numbered_matches:
                    return ", ".join(numbered_matches)

            if tmpl_cat and tmpl_cat != "General":
                cat_matches = [eq for eq in all_equip if infer_checklist_category(eq) == tmpl_cat]
                if cat_matches:
                    return ", ".join(cat_matches)

    m_inv = re.search(r"\binv(?:erter)?[\s_-]*(\d+)\b", tmpl_name, re.IGNORECASE)
    if m_inv:
        return f"Inverter {m_inv.group(1)}"
    return tmpl_name


class ChecklistTemplate(models.Model):
    name = models.CharField(max_length=160)
    category = models.CharField(max_length=60, blank=True, default="")
    document_number = models.CharField(max_length=120, blank=True)
    company = models.CharField(max_length=200, blank=True)
    equipment_name = models.CharField(max_length=200, blank=True)
    version = models.PositiveIntegerField(default=1)
    active = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["name", "-version"]
        constraints = [models.UniqueConstraint(fields=["name", "version"], name="unique_template_version")]

    def resolved_category(self):
        return (self.category or "").strip() or infer_checklist_category(self.name, self.equipment_name, self.document_number)

    def __str__(self):
        return f"{self.name} v{self.version}"


class ChecklistItem(models.Model):
    class AnswerType(models.TextChoices):
        NUMBER = "number", "Numeric reading"
        BOOLEAN = "boolean", "Yes / No"
        TEXT = "text", "Short value / text"
        CHOICES = "choices", "Dropdown choices"

    template = models.ForeignKey(ChecklistTemplate, on_delete=models.CASCADE, related_name="items")
    section = models.CharField(max_length=160, blank=True)
    checkpoint_no = models.PositiveIntegerField()
    description = models.TextField()
    standard = models.TextField(blank=True)
    method = models.CharField(max_length=120, blank=True)
    frequency = models.CharField(max_length=80, blank=True)
    answer_type = models.CharField(max_length=12, choices=AnswerType.choices, default=AnswerType.TEXT)
    # Comma-separated list of allowed option values when answer_type == "choices".
    # E.g. "Operational,Non Operational,Spare"
    choices = models.CharField(max_length=500, blank=True)
    unit = models.CharField(max_length=40, blank=True)
    required = models.BooleanField(default=True)
    position = models.PositiveIntegerField(default=0)

    def get_choices_list(self):
        """Return a list of option strings for a 'choices' type item."""
        if self.answer_type == self.AnswerType.CHOICES and self.choices:
            return [c.strip() for c in self.choices.split(",") if c.strip()]
        return []

    class Meta:
        ordering = ["position", "checkpoint_no"]
        constraints = [models.UniqueConstraint(fields=["template", "checkpoint_no"], name="unique_checkpoint_per_template")]

    def __str__(self):
        return f"{self.checkpoint_no}. {self.description[:70]}"


class SiteChecklist(models.Model):
    site = models.ForeignKey(Site, on_delete=models.CASCADE, related_name="checklists")
    template = models.ForeignKey(ChecklistTemplate, on_delete=models.CASCADE, related_name="site_checklists")
    category = models.CharField(max_length=60, blank=True, default="")
    position = models.PositiveIntegerField(default=0)
    active = models.BooleanField(default=True)

    class Meta:
        ordering = ["position", "template__name"]
        constraints = [models.UniqueConstraint(fields=["site", "template"], name="unique_template_per_site")]

    def resolved_category(self):
        return (self.category or "").strip() or self.template.resolved_category()

    def __str__(self):
        return f"{self.site} — {self.template}"


class Submission(models.Model):
    site = models.ForeignKey(Site, on_delete=models.PROTECT, related_name="submissions")
    template = models.ForeignKey(ChecklistTemplate, on_delete=models.PROTECT, related_name="submissions")
    submitted_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.PROTECT, related_name="checklist_submissions")
    checklist_date = models.DateField()
    submitted_at = models.DateTimeField(default=timezone.now)
    remarks = models.TextField(blank=True)
    signer_name = models.CharField(max_length=160, blank=True)
    equipment_snapshot = models.TextField(blank=True)

    class Meta:
        ordering = ["-checklist_date", "site__name"]
        constraints = [models.UniqueConstraint(fields=["site", "template", "checklist_date"], name="one_submission_per_site_template_per_day")]

    def clean(self):
        if self.submitted_by_id and not self.submitted_by.is_superuser:
            assignment = getattr(self.submitted_by, "site_assignment", None)
            if assignment is None or assignment.site_id != self.site_id:
                raise ValidationError("This user is not assigned to the selected site.")
        if self.submitted_by_id:
            assignment = getattr(self.submitted_by, "site_assignment", None)
            if not self.submitted_by.is_superuser and assignment and UserSite.objects.filter(site_id=assignment.site_id).exclude(user_id=self.submitted_by_id).exists():
                raise ValidationError("This site is assigned to a different user.")

    def __str__(self):
        return f"{self.site} — {self.checklist_date}"


class Answer(models.Model):
    submission = models.ForeignKey(Submission, on_delete=models.CASCADE, related_name="answers")
    item = models.ForeignKey(ChecklistItem, on_delete=models.PROTECT, related_name="answers")
    value = models.CharField(max_length=300, blank=True)
    comment = models.TextField(blank=True)

    class Meta:
        constraints = [models.UniqueConstraint(fields=["submission", "item"], name="one_answer_per_item")]

    def __str__(self):
        return f"{self.item.checkpoint_no}: {self.value}"


class GlobalSettings(models.Model):
    deadline = models.TimeField(default=time(17, 0))
    updated_at = models.DateTimeField(auto_now=True)
    updated_by = models.ForeignKey(settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.SET_NULL)

    def save(self, *args, **kwargs):
        self.pk = 1
        super().save(*args, **kwargs)

    @classmethod
    def get_solo(cls):
        obj, _ = cls.objects.get_or_create(pk=1)
        return obj

    @classmethod
    def is_open(cls, now=None):
        now = timezone.localtime(now) if now is not None else timezone.localtime()
        return now.time().replace(tzinfo=None) < cls.get_solo().deadline

    def __str__(self):
        return f"Daily deadline: {self.deadline:%H:%M} IST"
