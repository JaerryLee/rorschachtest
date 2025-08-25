from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('scoring', '0014_client_consent'),
    ]

    operations = [
        migrations.AlterModelOptions(
            name='responsecode',
            options={'ordering': ['client_id', 'card', 'response_num']},
        ),
        migrations.AlterField(
            model_name='cardimages',
            name='detail_img',
            field=models.FileField(blank=True, null=True, upload_to='images/location'),
        ),
        migrations.AlterField(
            model_name='cardimages',
            name='img_file',
            field=models.FileField(blank=True, null=True, upload_to='images/card'),
        ),
        migrations.AlterField(
            model_name='datatable',
            name='column2',
            field=models.TextField(),
        ),
        migrations.AlterField(
            model_name='structuralsummary',
            name='fqx_none',
            field=models.PositiveIntegerField(default=0, verbose_name='fqx_no'),
        ),
        migrations.AlterField(
            model_name='structuralsummary',
            name='mq_none',
            field=models.PositiveIntegerField(default=0, verbose_name='mq_no'),
        ),
        migrations.AlterField(
            model_name='structuralsummary',
            name='wd_none',
            field=models.PositiveIntegerField(default=0, verbose_name='wd_no'),
        ),
        migrations.AddIndex(
            model_name='responsecode',
            index=models.Index(fields=['client', 'card'], name='scoring_res_client__f53718_idx'),
        ),
    ]
