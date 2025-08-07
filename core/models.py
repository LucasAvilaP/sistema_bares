# core/models.py

from django.db import models
from django.utils import timezone
from django.contrib.auth.models import User
from decimal import Decimal
from django.db import transaction







class Restaurante(models.Model):
    nome = models.CharField(max_length=100)

    def __str__(self):
        return self.nome

class Bar(models.Model):
    nome = models.CharField(max_length=100)
    restaurante = models.ForeignKey(Restaurante, on_delete=models.CASCADE, related_name='bares')
    is_estoque_central = models.BooleanField(default=False)  # Este bar será o "estoque" do restaurante

    def __str__(self):
        return f"{self.nome} ({self.restaurante.nome})"



class Produto(models.Model):
    CATEGORIAS = (
        ('DESTILADO', 'Destilado'),
        ('CERVEJA', 'Cerveja'),
        ('VINHO', 'Vinho'),
        ('OUTRO', 'Outro'),
    )

    nome = models.CharField(max_length=100)
    unidade_medida = models.CharField(max_length=20, default='un')  # 'un', 'dose'
    categoria = models.CharField(max_length=20, choices=CATEGORIAS)
    doses_por_garrafa = models.PositiveIntegerField(null=True, blank=True)  # Ex: 16 doses por garrafa

    ativo = models.BooleanField(default=True)

    def __str__(self):
        return self.nome



class RecebimentoEstoque(models.Model):
    restaurante = models.ForeignKey(Restaurante, on_delete=models.CASCADE, related_name='recebimentos')
    bar = models.ForeignKey(Bar, on_delete=models.CASCADE, related_name='recebimentos')
    produto = models.ForeignKey(Produto, on_delete=models.CASCADE)
    quantidade = models.DecimalField(max_digits=10, decimal_places=2)
    data_recebimento = models.DateTimeField(default=timezone.now)

    def __str__(self):
        return f"{self.produto.nome} - {self.quantidade} un - {self.bar.nome}"
    



class TransferenciaBar(models.Model):
    restaurante = models.ForeignKey(Restaurante, on_delete=models.CASCADE, related_name='transferencias')
    origem = models.ForeignKey(Bar, on_delete=models.CASCADE, related_name='transferencias_saida')
    destino = models.ForeignKey(Bar, on_delete=models.CASCADE, related_name='transferencias_entrada')
    produto = models.ForeignKey(Produto, on_delete=models.CASCADE)
    quantidade = models.DecimalField(max_digits=10, decimal_places=2)
    usuario = models.ForeignKey(User, on_delete=models.SET_NULL, null=True)
    data_transferencia = models.DateTimeField(default=timezone.now)

    def __str__(self):
        return f"{self.produto.nome} | {self.quantidade} un | {self.origem.nome} → {self.destino.nome}"



class ContagemBar(models.Model):
    bar = models.ForeignKey(Bar, on_delete=models.CASCADE, related_name='contagens')
    produto = models.ForeignKey(Produto, on_delete=models.CASCADE)
    
    quantidade_garrafas_cheias = models.PositiveIntegerField(default=0)
    quantidade_doses_restantes = models.DecimalField(max_digits=6, decimal_places=2, default=0)

    data_contagem = models.DateTimeField(default=timezone.now)
    usuario = models.ForeignKey(User, on_delete=models.SET_NULL, null=True)
    observacao = models.TextField(blank=True, null=True)

    def __str__(self):
        return f"{self.bar.nome} - {self.produto.nome}"



from django.contrib.auth.models import User

class RequisicaoProduto(models.Model):
    STATUS_CHOICES = (
        ('PENDENTE', 'Pendente'),
        ('APROVADA', 'Aprovada'),
        ('NEGADA', 'Negada'),
        ('FALHA_ESTOQUE', 'Falha no Estoque'),
    )

    restaurante = models.ForeignKey(Restaurante, on_delete=models.CASCADE)
    bar = models.ForeignKey(Bar, on_delete=models.CASCADE)
    produto = models.ForeignKey(Produto, on_delete=models.CASCADE)
    quantidade_solicitada = models.DecimalField(max_digits=10, decimal_places=2)
    status = models.CharField(max_length=13, choices=STATUS_CHOICES, default='PENDENTE')
    data_solicitacao = models.DateTimeField(default=timezone.now)
    usuario = models.ForeignKey(User, on_delete=models.SET_NULL, null=True)
    observacao = models.TextField(blank=True, null=True)

    # ✅ NOVO CAMPO
    usuario_aprovador = models.ForeignKey(
        User,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='requisicoes_aprovadas'
    )

    def __str__(self):
        return f"{self.produto.nome} - {self.quantidade_solicitada} un ({self.status})"





class EstoqueBar(models.Model):
    bar = models.ForeignKey('Bar', on_delete=models.CASCADE)
    produto = models.ForeignKey('Produto', on_delete=models.CASCADE)
    quantidade_garrafas = models.DecimalField(max_digits=10, decimal_places=2, default=0)
    quantidade_doses = models.DecimalField(max_digits=10, decimal_places=2, default=0)

    class Meta:
        unique_together = ('bar', 'produto')

    def __str__(self):
        return f"{self.bar.nome} - {self.produto.nome}: {self.quantidade_garrafas} garrafas | {self.quantidade_doses} doses"

    @classmethod
    @transaction.atomic
    def retirar(cls, bar, produto, garrafas: Decimal, doses: Decimal = Decimal(0)) -> bool:
        estoque, _ = cls.objects.get_or_create(
            bar=bar,
            produto=produto,
            defaults={'quantidade_garrafas': 0, 'quantidade_doses': 0}
        )
        if estoque.quantidade_garrafas >= garrafas and estoque.quantidade_doses >= doses:
            estoque.quantidade_garrafas -= garrafas
            estoque.quantidade_doses -= doses
            estoque.save()
            return True
        return False

    @classmethod
    @transaction.atomic
    def adicionar(cls, bar, produto, garrafas: Decimal, doses: Decimal = Decimal(0)):
        estoque, _ = cls.objects.get_or_create(
            bar=bar,
            produto=produto,
            defaults={'quantidade_garrafas': 0, 'quantidade_doses': 0}
        )
        estoque.quantidade_garrafas += garrafas
        estoque.quantidade_doses += doses
        estoque.save()




class AcessoUsuarioBar(models.Model):
    user = models.ForeignKey(User, on_delete=models.CASCADE, related_name='acessos_bar')
    restaurante = models.ForeignKey(Restaurante, on_delete=models.CASCADE)
    bares = models.ManyToManyField(Bar)

    def __str__(self):
        return f'{self.user.username} - {self.restaurante.nome}'




class Evento(models.Model):
    nome = models.CharField(max_length=100)
    data_criacao = models.DateTimeField(auto_now_add=True)
    responsavel = models.ForeignKey(User, on_delete=models.SET_NULL, null=True)

    def __str__(self):
        return f"{self.nome} - {self.data_criacao.date()}"

class EventoProduto(models.Model):
    evento = models.ForeignKey(Evento, on_delete=models.CASCADE, related_name='produtos')
    produto = models.ForeignKey('Produto', on_delete=models.PROTECT)
    garrafas = models.PositiveIntegerField(default=0)
    doses = models.PositiveIntegerField(default=0)

    def __str__(self):
        return f"{self.produto.nome} - {self.garrafas} garrafas, {self.doses} doses"



class PermissaoPagina(models.Model):
    PAGINAS_CHOICES = [
        ('contagem', 'Contagem'),
        ('transferencias', 'Transferências'),
        ('eventos', 'Eventos'),
        ('requisicoes', 'Requisições'),
        ('aprovacao', 'Aprovacao'),
        ('historico_cont', 'Histórico_cont'),
        ('historico_requi', 'Historico_requi'),
        ('historico_transf', 'Historico_transf'),
        ('relatorios', 'Relatórios'),
        ('entrada_mercadoria', 'Entrada_mercadoria'),
    ]

    user = models.ForeignKey(User, on_delete=models.CASCADE)
    nome_pagina = models.CharField(max_length=50, choices=PAGINAS_CHOICES)

    def __str__(self):
        return f"{self.user.username} - {self.get_nome_pagina_display()}"