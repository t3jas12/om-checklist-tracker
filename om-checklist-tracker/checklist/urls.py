from django.contrib.auth import views as auth_views
from django.urls import path

from . import views

urlpatterns = [
    path("", views.dashboard, name="dashboard"),
    path("login/", auth_views.LoginView.as_view(template_name="registration/login.html"), name="login"),
    path("logout/", auth_views.LogoutView.as_view(), name="logout"),
    path("checklist/", views.checklist_form, name="checklist-form"),
    path("checklist/finalize/", views.finalize_checklist, name="finalize-checklist"),
    path("submissions/<int:submission_id>/review/", views.submission_review, name="submission-review"),
    path("submissions/<int:submission_id>/excel/", views.operator_excel_review, name="operator-excel-review"),
    path("history/", views.history, name="history"),
    path("executive-summary/", views.executive_summary, name="executive-summary"),
    path("history/export/monthly/<str:file_type>/", views.operator_monthly_export, name="operator-monthly-export"),
    path("admin-settings/", views.admin_settings, name="admin-settings"),
    path("manage/users/", views.manage_users, name="manage-users"),
    path("manage/users/<int:user_id>/edit/", views.edit_user, name="edit-user"),
    path("manage/sites/", views.manage_sites, name="manage-sites"),
    path("manage/sites/<int:site_id>/edit/", views.edit_site, name="edit-site"),
    path("manage/sites/<int:site_id>/equipment/add/", views.add_site_equipment, name="add-site-equipment"),
    path("manage/sites/<int:site_id>/equipment/<int:equipment_id>/remove/", views.remove_site_equipment, name="remove-site-equipment"),
    path("manage/sites/<int:site_id>/checklists/assign/", views.assign_site_checklist, name="assign-site-checklist"),
    path("manage/sites/<int:site_id>/checklists/<int:site_checklist_id>/category/", views.update_site_checklist_category, name="update-site-checklist-category"),
    path("manage/sites/<int:site_id>/checklists/<int:site_checklist_id>/remove/", views.remove_site_checklist, name="remove-site-checklist"),
    path("manage/checklists/", views.manage_checklists, name="manage-checklists"),
    path("manage/checklists/<int:template_id>/preview/", views.preview_checklist_template, name="preview-checklist-template"),
    path("manage/checklists/<int:template_id>/delete/", views.delete_checklist_template, name="delete-checklist-template"),
    path("admin/export/monthly/<str:file_type>/", views.export_monthly, name="export-monthly"),
    path("admin/submissions/<int:submission_id>/<str:file_type>/", views.export_submission, name="export-submission"),
]
