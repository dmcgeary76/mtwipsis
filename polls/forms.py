from django import forms
from .models import FileUploadModel

class FileUploadForm(forms.ModelForm):
    class Meta:
        model = FileUploadModel
        fields = [
            'name',
            'file_upload'
        ]
