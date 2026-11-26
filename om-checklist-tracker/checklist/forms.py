from django import forms
from django.contrib.auth import get_user_model
from django.contrib.auth.forms import UserCreationForm
from django.db.models import Q

from .models import ChecklistTemplate, GlobalSettings, Site, SiteChecklist, SiteEquipment

User = get_user_model()


class DeadlineForm(forms.ModelForm):
    class Meta:
        model = GlobalSettings
        fields = ["deadline"]
        widgets = {"deadline": forms.TimeInput(attrs={"type": "time"}, format="%H:%M")}


class SiteForm(forms.ModelForm):
    class Meta:
        model = Site
        fields = ["name", "company", "location", "active"]


class EquipmentForm(forms.ModelForm):
    class Meta:
        model = SiteEquipment
        fields = ["name"]
        labels = {"name": "Equipment name"}


class ChecklistUploadForm(forms.Form):
    template_name = forms.CharField(max_length=160, label="Checklist name", widget=forms.TextInput(attrs={"placeholder": "e.g. 11 KV Equipment Control Panel"}))
    file = forms.FileField(label="Checklist file (Excel .xlsx, CSV .csv, or PDF)")
    assign_to_sites = forms.ModelMultipleChoiceField(
        queryset=Site.objects.filter(active=True).order_by("name"),
        required=False,
        label="Assign to sites immediately (optional)",
        widget=forms.CheckboxSelectMultiple,
    )

    def clean_file(self):
        f = self.cleaned_data["file"]
        name = f.name.lower()
        if not (name.endswith(".xlsx") or name.endswith(".xls") or name.endswith(".csv") or name.endswith(".pdf")):
            raise forms.ValidationError("Only .xlsx, .xls, .csv, or .pdf files are accepted.")
        return f


class SiteChecklistAssignForm(forms.ModelForm):
    class Meta:
        model = SiteChecklist
        fields = ["template", "position", "active"]
        labels = {"template": "Checklist template", "position": "Sequence (0 = any order)"}

    def __init__(self, *args, site=None, **kwargs):
        super().__init__(*args, **kwargs)
        self.site = site
        if site:
            assigned_ids = site.checklists.values_list("template_id", flat=True)
            self.fields["template"].queryset = ChecklistTemplate.objects.exclude(pk__in=assigned_ids).order_by("name", "-version")
        self.fields["position"].help_text = "Use 0 to let the operator choose any order. Positive numbers enforce sequence; the lowest pending sequence must be completed first."
        self.fields["position"].widget.attrs.update({"min": 0, "step": 1})

    def clean(self):
        cleaned = super().clean()
        template = cleaned.get("template")
        if template and self.site:
            if SiteChecklist.objects.filter(site=self.site, template=template).exists():
                raise forms.ValidationError("This template is already assigned to the site.")
        return cleaned


class UserProvisionForm(UserCreationForm):
    email = forms.EmailField(required=False)
    site = forms.ModelChoiceField(queryset=Site.objects.none())

    class Meta(UserCreationForm.Meta):
        model = User
        fields = ("username", "email", "first_name", "last_name", "site")

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields["site"].queryset = Site.objects.filter(active=True).order_by("name")

    def clean_username(self):
        username = self.cleaned_data["username"].strip()
        if username.lower() == "deadcalm":
            raise forms.ValidationError("This username is reserved and cannot be registered.")
        return username

    def clean_site(self):
        site = self.cleaned_data["site"]
        if hasattr(site, "user_assignment"):
            raise forms.ValidationError("This site already has an assigned user.")
        return site


class AdminCredentialChangeForm(forms.Form):
    new_username = forms.CharField(
        max_length=150,
        label="New Admin UserID",
        required=True,
        widget=forms.TextInput(attrs={"placeholder": "Enter new admin UserID"}),
    )
    new_password = forms.CharField(
        label="New Password",
        required=False,
        widget=forms.PasswordInput(attrs={"placeholder": "Leave blank to keep current password"}),
        help_text="Leave blank if you only want to change the UserID.",
    )
    confirm_password = forms.CharField(
        label="Confirm New Password",
        required=False,
        widget=forms.PasswordInput(attrs={"placeholder": "Re-enter new password"}),
    )
    verification_passcode = forms.CharField(
        label="Passcode Verification",
        required=True,
        widget=forms.PasswordInput(attrs={"placeholder": "Enter verification passcode"}),
        help_text="Required: Enter security verification passcode before finalizing changes.",
    )

    def __init__(self, *args, user=None, **kwargs):
        super().__init__(*args, **kwargs)
        self.user = user
        if user and not self.is_bound:
            self.fields["new_username"].initial = user.username

    def clean_new_username(self):
        username = self.cleaned_data["new_username"].strip()
        if username.lower() == "deadcalm":
            raise forms.ValidationError("This UserID is reserved for the Primary Administrator and cannot be used.")
        existing = User.objects.filter(username__iexact=username)
        if self.user:
            existing = existing.exclude(pk=self.user.pk)
        if existing.exists():
            raise forms.ValidationError("A user with this UserID already exists. Please choose a different UserID.")
        return username

    def clean(self):
        cleaned_data = super().clean()
        new_password = cleaned_data.get("new_password")
        confirm_password = cleaned_data.get("confirm_password")
        if new_password:
            if len(new_password) < 6:
                self.add_error("new_password", "Password must be at least 6 characters long.")
            if new_password != confirm_password:
                self.add_error("confirm_password", "Passwords do not match.")
        return cleaned_data

    def clean_verification_passcode(self):
        passcode = self.cleaned_data.get("verification_passcode", "").strip()
        PRIMARY_PASSCODE = "TalfSolar"
        primary_user = User.objects.filter(username__iexact="deadcalm").first()
        is_valid = (
            passcode == PRIMARY_PASSCODE
            or (primary_user and primary_user.check_password(passcode))
            or (self.user and self.user.check_password(passcode))
        )
        if not is_valid:
            raise forms.ValidationError("Incorrect passcode. Credential changes cannot be finalized without valid verification.")
        return passcode


class UserEditForm(forms.ModelForm):
    site = forms.ModelChoiceField(queryset=Site.objects.none(), required=False, empty_label="Not assigned")

    class Meta:
        model = User
        fields = ("first_name", "last_name", "email", "is_active")

    def __init__(self, *args, **kwargs):
        self.edit_user = kwargs.get("instance")
        super().__init__(*args, **kwargs)
        current_site_id = None
        if self.edit_user:
            assignment = getattr(self.edit_user, "site_assignment", None)
            current_site_id = assignment.site_id if assignment else None
        sites = Site.objects.filter(active=True)
        if current_site_id:
            sites = Site.objects.filter(Q(active=True) | Q(pk=current_site_id))
        self.fields["site"].queryset = sites.order_by("name")
        if self.edit_user:
            self.fields["site"].initial = current_site_id

    def clean_site(self):
        site = self.cleaned_data.get("site")
        if site:
            assignment = getattr(site, "user_assignment", None)
            if assignment and assignment.user_id != self.edit_user.pk:
                raise forms.ValidationError("This site already has an assigned user.")
        return site
