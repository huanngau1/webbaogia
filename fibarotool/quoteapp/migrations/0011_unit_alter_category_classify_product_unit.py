# Generated by Django 5.0.7 on 2024-08-11 05:14

import django.db.models.deletion
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('quoteapp', '0010_category_classify_alter_product_category_and_more'),
    ]

    operations = [
        migrations.CreateModel(
            name='Unit',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('name', models.CharField(blank=True, max_length=255, null=True)),
            ],
        ),
        migrations.AlterField(
            model_name='category',
            name='classify',
            field=models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name='categories', to='quoteapp.classify'),
        ),
        migrations.AddField(
            model_name='product',
            name='unit',
            field=models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.CASCADE, to='quoteapp.unit'),
        ),
    ]
