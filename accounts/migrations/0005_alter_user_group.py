from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('accounts', '0004_alter_user_consent'),
    ]

    operations = [
        migrations.AlterField(
            model_name='user',
            name='group',
            field=models.CharField(choices=[('beginner', '초급'), ('intermediate', '중급'), ('advanced', '고급')], default='beginner', max_length=20),
        ),
    ]
