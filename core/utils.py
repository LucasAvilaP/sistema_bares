from .models import EstoqueBar

def atualizar_estoque(bar, produto, quantidade_delta):
    estoque, _ = EstoqueBar.objects.get_or_create(bar=bar, produto=produto)
    estoque.quantidade += quantidade_delta
    estoque.save()
