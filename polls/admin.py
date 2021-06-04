from django.contrib import admin
from .models import Question, FileUploadModel

# Register your models here.
class QuestionAdmin(admin.ModelAdmin):
    fieldsets = [
        (None,               {'fields': ['question_text']}),
        ('Date information', {'fields': ['pub_date']}),
    ]

admin.site.register(Question)
admin.site.register(FileUploadModel)
