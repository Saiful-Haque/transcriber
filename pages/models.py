from django.db import models

class TranscriptionJob(models.Model):
    TRANSCRIBE_OPTIONS = [
        ('verbatim', 'Verbatim'),
        ('english', 'English Translation'),
    ]
    
    raw_text = models.TextField(null=True, blank=True)
    audio_file = models.FileField(upload_to='audios/', null=True, blank=True)
    formatted_content = models.TextField()
    transcribe_type = models.CharField(max_length=20, choices=TRANSCRIBE_OPTIONS)
    
    # Timing Stats
    created_at = models.DateTimeField(auto_now_add=True)
    finish_time = models.DateTimeField(null=True, blank=True)
    upload_duration = models.FloatField(default=0) # in seconds
    transcription_duration = models.FloatField(default=0) # in seconds
    total_duration = models.FloatField(default=0)
    
    class Meta:
        ordering = ['-created_at']

    def __str__(self):
        return f"Job {self.id} - {self.created_at.strftime('%Y-%m-%d %H:%M')}"
