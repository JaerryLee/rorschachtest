from django.contrib import admin
from django.contrib.auth.admin import UserAdmin as BaseUserAdmin
from django.utils.translation import gettext_lazy as _
from .models import User

@admin.register(User)
class UserAdmin(BaseUserAdmin):
    list_display = ("username", "email", "phone", "group", "group_label",
                    "is_active", "is_staff", "is_superuser", "last_login")
    def group_label(self, obj):
        return obj.get_group_display()
    group_label.short_description = "권한(라벨)"
    
    list_filter  = ("group", "is_active", "is_staff", "is_superuser")
    search_fields = ("username", "email", "phone")
    ordering = ("-date_joined",)

    fieldsets = (
        (None, {"fields": ("username", "password")}),
        (_("Personal info"), {"fields": ("first_name", "last_name", "email", "phone")}),
        (_("Permissions"), {"fields": ("group", "consent", "is_active", "is_staff", "is_superuser", "groups", "user_permissions")}),
        (_("Important dates"), {"fields": ("last_login", "date_joined")}),
    )

    add_fieldsets = (
        (None, {
            "classes": ("wide",),
            "fields": ("username", "email", "phone", "group","consent",
                       "password1", "password2", "is_active", "is_staff", "is_superuser"),
        }),
    )
