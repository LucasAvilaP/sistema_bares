# core/management/commands/criar_estoques_faltantes.py
from django.core.management.base import BaseCommand
from django.db import transaction
from core.models import Bar, Produto, EstoqueBar

class Command(BaseCommand):
    help = "Cria EstoqueBar faltante para todas as combinações Bar × Produto (apenas onde não existe)."

    def add_arguments(self, parser):
        parser.add_argument('--dry-run', action='store_true', help='Mostra o que seria criado, sem gravar.')

    def handle(self, *args, **opts):
        dry = opts['dry_run']

        bars = list(Bar.objects.only('id'))
        produtos = list(Produto.objects.filter(ativo=True).only('id'))

        existentes = set(EstoqueBar.objects.values_list('bar_id', 'produto_id'))
        a_criar = []
        for b in bars:
            for p in produtos:
                if (b.id, p.id) not in existentes:
                    a_criar.append(EstoqueBar(bar_id=b.id, produto_id=p.id))

        if not a_criar:
            self.stdout.write(self.style.SUCCESS("Nada a criar. Todos os estoques existem."))
            return

        self.stdout.write(f"Faltantes: {len(a_criar)} registros.")
        if dry:
            self.stdout.write(self.style.WARNING("Dry-run: nenhuma linha gravada."))
            return

        with transaction.atomic():
            EstoqueBar.objects.bulk_create(a_criar, ignore_conflicts=True, batch_size=2000)

        self.stdout.write(self.style.SUCCESS(f"Criados {len(a_criar)} registros de EstoqueBar."))
