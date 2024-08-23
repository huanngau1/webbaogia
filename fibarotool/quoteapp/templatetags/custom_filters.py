from django import template

register = template.Library()

@register.filter
def format_currency(value):
    try:
        value = float(value)
    except (ValueError, TypeError):
        return value
    return "{:,.0f}".format(value)
