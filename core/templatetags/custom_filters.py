from django import template
from decimal import Decimal
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


@register.filter
def as_dot(value, places=2):
    try:
        q = Decimal(str(value))
    except Exception:
        return ''
    fmt = f'{{0:.{int(places)}f}}'
    return fmt.format(q)  # sempre com ponto


@register.filter
def get_item(d, k):
    try:
        return d.get(k, "")
    except Exception:
        return ""