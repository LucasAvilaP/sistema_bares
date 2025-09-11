from .models import EstoqueBar

def atualizar_estoque(bar, produto, quantidade_delta):
    estoque, _ = EstoqueBar.objects.get_or_create(bar=bar, produto=produto)
    estoque.quantidade += quantidade_delta
    estoque.save()



from decimal import Decimal

def calcular_totais_ml_e_doses(produto, garrafas, doses_avulsas):
    """
    total_ml = (garrafas * volume_garrafa_ml) + (doses_avulsas * dose_padrao_ml)
    doses_equivalentes = total_ml / dose_padrao_ml
    Retorna: total_ml (Decimal), doses_equivalentes (Decimal), doses_por_garrafa_inferidas (int)
    """
    garrafas = Decimal(garrafas or 0)
    doses_avulsas = Decimal(doses_avulsas or 0)

    # Defaults seguros
    dose_ml = Decimal(getattr(produto, 'dose_padrao_ml', 50) or 50)
    volume_ml = Decimal(getattr(produto, 'volume_garrafa_ml', 0) or 0)

    total_ml = (garrafas * volume_ml) + (doses_avulsas * dose_ml)
    doses_equivalentes = (total_ml / dose_ml) if dose_ml else doses_avulsas

    dpg = 0
    if volume_ml and dose_ml:
        try:
            dpg = int(round(volume_ml / dose_ml))
        except Exception:
            dpg = 0

    return total_ml, doses_equivalentes, dpg
