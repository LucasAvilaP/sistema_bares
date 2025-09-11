# core/models.py

from django.db import models
from django.utils import timezone
from django.contrib.auth.models import User
from decimal import Decimal
from django.db import transaction
from django.db.models import F, Q



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
    codigo = models.CharField(max_length=30, unique=True, db_index=True, blank=True, null=True)
    unidade_medida = models.CharField(max_length=20, default='un')
    categoria = models.CharField(max_length=20, choices=CATEGORIAS)

    # Existente: opcional, continua útil
    doses_por_garrafa = models.PositiveIntegerField(null=True, blank=True)

    # NOVOS:
    volume_garrafa_ml = models.PositiveIntegerField(null=True, blank=True, help_text="Ex.: 700, 750, 770, 1000, 1750")
    dose_padrao_ml = models.PositiveIntegerField(default=50, help_text="Tamanho da dose padrão em mL (padrão 50)")

    ativo = models.BooleanField(default=True)

    class Meta:
        ordering = ['nome']

    def __str__(self):
        return self.nome

    # Helpers (opcionais, mas práticos):
    def get_dose_ml(self) -> int:
        return self.dose_padrao_ml or 50

    def get_doses_por_garrafa(self):
        """
        Prioriza volume_garrafa_ml, senão cai para doses_por_garrafa.
        """
        if self.volume_garrafa_ml and self.get_dose_ml():
            return round(self.volume_garrafa_ml / self.get_dose_ml())
        return self.doses_por_garrafa or 0




class RecebimentoEstoque(models.Model):
    restaurante = models.ForeignKey(Restaurante, on_delete=models.CASCADE, related_name='recebimentos')
    usuario = models.ForeignKey(User, on_delete=models.SET_NULL, null=True)
    bar = models.ForeignKey(Bar, on_delete=models.CASCADE, related_name='recebimentos')
    produto = models.ForeignKey(Produto, on_delete=models.CASCADE)
    quantidade = models.DecimalField(max_digits=10, decimal_places=2)
    data_recebimento = models.DateTimeField(default=timezone.now)


    class Meta:
        indexes = [
            models.Index(fields=['restaurante', 'data_recebimento'], name='ix_receb_rest_data'),
            # se preferir created_at:
            # models.Index(fields=['restaurante', 'created_at'], name='ix_receb_rest_created')
        ]


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

    # quem decidiu
    usuario_aprovador = models.ForeignKey(
        User, on_delete=models.SET_NULL, null=True, blank=True,
        related_name='requisicoes_aprovadas'
    )

    # 🔹 NOVOS CAMPOS
    motivo_negativa = models.TextField(blank=True, null=True)
    data_decisao = models.DateTimeField(blank=True, null=True)

    def __str__(self):
        return f"{self.produto.nome} - {self.quantidade_solicitada} un ({self.status})"






class EstoqueBar(models.Model):
    bar = models.ForeignKey('Bar', on_delete=models.CASCADE)
    produto = models.ForeignKey('Produto', on_delete=models.CASCADE)
    quantidade_garrafas = models.DecimalField(max_digits=10, decimal_places=2, default=0)
    quantidade_doses    = models.DecimalField(max_digits=10, decimal_places=2, default=0)

    class Meta:
        constraints = [
            models.UniqueConstraint(fields=['bar', 'produto'], name='uniq_estoque_bar_produto'),
            # evita ficar negativo em qualquer via
            models.CheckConstraint(
                check=Q(quantidade_garrafas__gte=0) & Q(quantidade_doses__gte=0),
                name='chk_estoque_nao_negativo'
            ),
        ]
        indexes = [models.Index(fields=['bar', 'produto'])]
        ordering = ['bar__nome', 'produto__nome']

    def __str__(self):
        return f"{self.bar.nome} - {self.produto.nome}: {self.quantidade_garrafas} garrafas | {self.quantidade_doses} doses"

    @classmethod
    @transaction.atomic
    def retirar(cls, bar, produto, garrafas: Decimal = Decimal('0'), doses: Decimal = Decimal('0')) -> bool:
        """Debita SOMENTE os campos com quantidade > 0.
           Usa lock de linha + F() para evitar race/lost update."""
        garrafas = Decimal(garrafas or 0)
        doses    = Decimal(doses or 0)

        est, _ = (cls.objects
                  .select_for_update()
                  .get_or_create(bar=bar, produto=produto,
                                 defaults={'quantidade_garrafas': Decimal('0'),
                                           'quantidade_doses': Decimal('0')}))

        # validações por campo (independentes)
        if garrafas > 0 and est.quantidade_garrafas < garrafas:
            return False
        if doses > 0 and est.quantidade_doses < doses:
            return False

        # atualizações atômicas
        if garrafas > 0:
            cls.objects.filter(pk=est.pk).update(
                quantidade_garrafas=F('quantidade_garrafas') - garrafas
            )
        if doses > 0:
            cls.objects.filter(pk=est.pk).update(
                quantidade_doses=F('quantidade_doses') - doses
            )
        return True

    @classmethod
    @transaction.atomic
    def adicionar(cls, bar, produto, garrafas: Decimal = Decimal('0'), doses: Decimal = Decimal('0')) -> None:
        garrafas = Decimal(garrafas or 0)
        doses    = Decimal(doses or 0)

        est, _ = (cls.objects
                  .select_for_update()
                  .get_or_create(bar=bar, produto=produto,
                                 defaults={'quantidade_garrafas': Decimal('0'),
                                           'quantidade_doses': Decimal('0')}))
        if garrafas > 0:
            cls.objects.filter(pk=est.pk).update(
                quantidade_garrafas=F('quantidade_garrafas') + garrafas
            )
        if doses > 0:
            cls.objects.filter(pk=est.pk).update(
                quantidade_doses=F('quantidade_doses') + doses
            )

    @classmethod
    @transaction.atomic
    def transferir(cls, origem, destino, produto,
                   garrafas: Decimal = Decimal('0'), doses: Decimal = Decimal('0')) -> bool:
        """Débito e crédito na MESMA transação e com as duas linhas travadas."""
        garrafas = Decimal(garrafas or 0)
        doses    = Decimal(doses or 0)

        # trava as duas linhas envolvidas
        est_origem, _ = (cls.objects.select_for_update()
                         .get_or_create(bar=origem, produto=produto,
                                        defaults={'quantidade_garrafas': Decimal('0'),
                                                  'quantidade_doses': Decimal('0')}))
        est_dest, _   = (cls.objects.select_for_update()
                         .get_or_create(bar=destino, produto=produto,
                                        defaults={'quantidade_garrafas': Decimal('0'),
                                                  'quantidade_doses': Decimal('0')}))

        # checa disponibilidade independente
        if garrafas > 0 and est_origem.quantidade_garrafas < garrafas:
            return False
        if doses > 0 and est_origem.quantidade_doses < doses:
            return False

        # faz as duas pernas com F() (atômico)
        if garrafas > 0:
            cls.objects.filter(pk=est_origem.pk).update(
                quantidade_garrafas=F('quantidade_garrafas') - garrafas
            )
            cls.objects.filter(pk=est_dest.pk).update(
                quantidade_garrafas=F('quantidade_garrafas') + garrafas
            )
        if doses > 0:
            cls.objects.filter(pk=est_origem.pk).update(
                quantidade_doses=F('quantidade_doses') - doses
            )
            cls.objects.filter(pk=est_dest.pk).update(
                quantidade_doses=F('quantidade_doses') + doses
            )
        return True




class AcessoUsuarioBar(models.Model):
    user = models.ForeignKey(User, on_delete=models.CASCADE, related_name='acessos_bar')
    restaurante = models.ForeignKey(Restaurante, on_delete=models.CASCADE)
    bares = models.ManyToManyField(Bar)

    def __str__(self):
        return f'{self.user.username} - {self.restaurante.nome}'




class Evento(models.Model):
    STATUS_CHOICES = (
        ('ABERTO', 'Aberto'),
        ('FINALIZADO', 'Finalizado'),
    )
    restaurante = models.ForeignKey(
        Restaurante,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='eventos'
    )
    nome = models.CharField(max_length=100)
    data_criacao = models.DateTimeField(auto_now_add=True)
    responsavel = models.ForeignKey(User, on_delete=models.SET_NULL, null=True)

    # ✅ já adicionados antes
    numero_pessoas = models.PositiveIntegerField(null=True, blank=True)
    horas = models.DecimalField(max_digits=4, decimal_places=1, null=True, blank=True)

    # ✅ NOVOS CAMPOS (fluxo)
    status = models.CharField(max_length=12, choices=STATUS_CHOICES, default='ABERTO')
    data_evento = models.DateField(default=timezone.localdate)  # dia do evento
    finalizado_em = models.DateTimeField(null=True, blank=True)
    supervisor_finalizou = models.ForeignKey(
        User, on_delete=models.SET_NULL, null=True, blank=True, related_name='eventos_finalizados'
    )

     # ✅ NOVOS CAMPOS: controle de baixa do estoque
    baixado_estoque = models.BooleanField(default=False)
    baixado_por = models.ForeignKey(
        User, on_delete=models.SET_NULL, null=True, blank=True, related_name='eventos_baixados'
    )
    baixado_em = models.DateTimeField(null=True, blank=True)
    baixado_obs = models.CharField(max_length=255, blank=True, default="")

    class Meta:
        indexes = [
            models.Index(fields=['data_evento', 'baixado_estoque']),
        ]


    def __str__(self):
        return f"{self.nome} ({self.get_status_display()}) - {self.data_evento}"

class EventoProduto(models.Model):
    evento = models.ForeignKey(Evento, on_delete=models.CASCADE, related_name='produtos')
    produto = models.ForeignKey('Produto', on_delete=models.PROTECT)
    garrafas = models.PositiveIntegerField(default=0)
    doses = models.PositiveIntegerField(default=0)

    def __str__(self):
        return f"{self.produto.nome} - {self.garrafas} garrafas, {self.doses} doses"


# ✅ Catálogo simples de alimentos (somente para eventos)
class Alimento(models.Model):
    UNIDADE_CHOICES = (
        ('un', 'Unidade'),
        ('kg', 'Quilo'),
        ('g', 'Grama'),
        ('porcao', 'Porção'),
        ('l', 'Litro'),
        ('ml', 'Mililitro'),
    )
    nome = models.CharField(max_length=255)
    codigo = models.CharField(max_length=30, unique=True)  # usado no autocomplete
    unidade = models.CharField(max_length=10, choices=UNIDADE_CHOICES, default='un')
    ativo = models.BooleanField(default=True)

    def __str__(self):
        return f"[{self.codigo}] {self.nome}"


# Itens de alimento consumidos no evento (não mexe em estoque de bar!)
class EventoAlimento(models.Model):
    evento = models.ForeignKey(Evento, on_delete=models.CASCADE, related_name='alimentos')
    alimento = models.ForeignKey(Alimento, on_delete=models.PROTECT)
    quantidade = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal('0.00'))

    def __str__(self):
        return f"{self.alimento.nome} - {self.quantidade} {self.alimento.unidade}"


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
        ('historico_entrada', 'Historico_entrada'),
        ('relatorios', 'Relatórios'),
        ('entrada_mercadoria', 'Entrada_mercadoria'),
        ('importacao', 'Importacao'),
    ]

    user = models.ForeignKey(User, on_delete=models.CASCADE)
    nome_pagina = models.CharField(max_length=50, choices=PAGINAS_CHOICES)

    def __str__(self):
        return f"{self.user.username} - {self.get_nome_pagina_display()}"
    



class PerdaProduto(models.Model):
    MOTIVOS = (
        ('QUEBRA', 'Quebra de garrafa'),
        ('DERRAMAMENTO', 'Derramamento'),
        ('SOBRA', 'Descarte de sobra'),
        ('QUALIDADE', 'Produto impróprio'),
        ('OUTRO', 'Outro'),
    )

    restaurante = models.ForeignKey(Restaurante, on_delete=models.CASCADE, related_name='perdas')
    bar         = models.ForeignKey(Bar, on_delete=models.CASCADE, related_name='perdas')
    produto     = models.ForeignKey(Produto, on_delete=models.PROTECT, related_name='perdas')

    garrafas = models.PositiveIntegerField(default=0)
    doses    = models.PositiveIntegerField(default=0)

    motivo      = models.CharField(max_length=20, choices=MOTIVOS)
    observacao  = models.TextField(blank=True, null=True)

    usuario        = models.ForeignKey(User, on_delete=models.SET_NULL, null=True)
    data_registro  = models.DateTimeField(default=timezone.now)

    # auditoria do estoque no momento da perda
    estoque_antes_garrafas = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal('0.00'))
    estoque_antes_doses    = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal('0.00'))
    estoque_depois_garrafas = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal('0.00'))
    estoque_depois_doses    = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal('0.00'))

    # ✅ marcação de baixa (conciliado)
    baixado = models.BooleanField(default=False)
    baixado_em = models.DateTimeField(blank=True, null=True)
    baixado_por = models.ForeignKey(
        User, on_delete=models.SET_NULL, null=True, blank=True, related_name='perdas_baixadas'
    )
    baixado_obs = models.CharField(max_length=255, blank=True, null=True)

    class Meta:
        ordering = ['-data_registro']
        indexes = [
            models.Index(fields=['baixado', 'data_registro']),
        ]

    def __str__(self):
        return f"{self.bar.nome} | {self.produto.nome} (-{self.garrafas} garrafas, -{self.doses} doses)"