from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('board', '0004_post_file'),
    ]

    operations = [
        migrations.AlterField(
            model_name='post',
            name='group',
            field=models.CharField(choices=[('beginner', 'Beginner'), ('intermediate', 'Intermediate'), ('advanced', 'Advanced')], default='beginner', max_length=20),
        ),
    ]
