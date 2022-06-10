from django.contrib import admin
from .models import User,Signs,Event_details
from import_export.admin import ImportExportModelAdmin
# Register your models here.
@admin.register(User)
class ViewAdmin(ImportExportModelAdmin):
    list_display = ['Name']

admin.site.register(Signs)
admin.site.register(Event_details)