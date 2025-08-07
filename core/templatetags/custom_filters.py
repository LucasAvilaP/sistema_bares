from django import template

register = template.Library()

@register.filter
def until(value, end):
    """Gera um range de value até end (ex: 1|until:13 → 1 a 12)"""
    return range(value, end)


@register.filter
def to_range(start, end):
    return range(start, end)



@register.filter
def mul(value, arg):
    try:
        return float(value) * float(arg)
    except (ValueError, TypeError):
        return 0

@register.filter
def floatval(value):
    try:
        return float(value)
    except (ValueError, TypeError):
        return 0
    


@register.filter
def sum_quantidade_solicitada(requisicoes):
    return sum([r.quantidade_solicitada for r in requisicoes])
