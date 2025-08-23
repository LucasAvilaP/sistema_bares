# core/signals.py
from django.db.models.signals import post_save
from django.dispatch import receiver
from django.db import transaction
from core.models import Produto, Bar, EstoqueBar

def _criar_estoques_para_bares(produto_id: int):
    """Cria EstoqueBar(produto=produto_id) para todos os bares que ainda não têm."""
    qs_bares = Bar.objects.only('id')
    existentes = set(
        EstoqueBar.objects.filter(produto_id=produto_id).values_list('bar_id', flat=True)
    )
    novos = [
        EstoqueBar(bar_id=b.id, produto_id=produto_id, quantidade_garrafas=0, quantidade_doses=0)
        for b in qs_bares if b.id not in existentes
    ]
    if novos:
        EstoqueBar.objects.bulk_create(novos, ignore_conflicts=True, batch_size=1000)

def _criar_estoques_para_produtos(bar_id: int):
    """Cria EstoqueBar(bar=bar_id) para todos os produtos que ainda não têm."""
    qs_produtos = Produto.objects.filter(ativo=True).only('id')
    existentes = set(
        EstoqueBar.objects.filter(bar_id=bar_id).values_list('produto_id', flat=True)
    )
    novos = [
        EstoqueBar(bar_id=bar_id, produto_id=p.id, quantidade_garrafas=0, quantidade_doses=0)
        for p in qs_produtos if p.id not in existentes
    ]
    if novos:
        EstoqueBar.objects.bulk_create(novos, ignore_conflicts=True, batch_size=1000)

@receiver(post_save, sender=Produto)
def produto_pos_save(sender, instance: Produto, created, **kwargs):
    # quando um novo produto for criado, cria estoques em todos os bares
    if created and instance.ativo:
        transaction.on_commit(lambda: _criar_estoques_para_bares(instance.id))

@receiver(post_save, sender=Bar)
def bar_pos_save(sender, instance: Bar, created, **kwargs):
    # quando um novo bar for criado, cria estoques para todos os produtos ativos
    if created:
        transaction.on_commit(lambda: _criar_estoques_para_produtos(instance.id))
