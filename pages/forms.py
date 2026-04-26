from django import forms

class ScribeForm(forms.Form):
    raw_text = forms.CharField(
        widget=forms.Textarea(attrs={
            'placeholder': 'Paste your interview transcription here...',
            'class': 'form-input',
            'rows': 10
        }),
        required=False,
        label=''
    )
    
    audio_file = forms.FileField(
        required=False,
        widget=forms.FileInput(attrs={'class': 'form-input', 'accept': 'audio/*'}),
        label='Upload Interview Audio'
    )
    
    TRANSCRIBE_OPTIONS = [
        ('verbatim', 'Verbatim (Keep Hinglish: "achha", "theek hai")'),
        ('english', 'Full English Translation'),
    ]
    
    transcribe_type = forms.ChoiceField(
        choices=TRANSCRIBE_OPTIONS,
        initial='verbatim',
        widget=forms.RadioSelect(attrs={'class': 'radio-group'}),
        label='Transcription Mode'
    )
