from django import forms
from django.contrib import admin, messages
from django.contrib.admin.helpers import ActionForm
from django.contrib.auth.admin import UserAdmin
from django.contrib.auth.models import User
from django.db import transaction
from django.urls import reverse
from django.http import HttpResponseRedirect

from .models import (
    Restaurante, Bar, Produto, RecebimentoEstoque,
    TransferenciaBar, ContagemBar, RequisicaoProduto,
    EstoqueBar, AcessoUsuarioBar, Evento, EventoProduto,
    PermissaoPagina, Alimento, EventoAlimento, PerdaProduto
)

# -------------------------------------------------------------------
# Registros simples
# -------------------------------------------------------------------
admin.site.register(Alimento)
admin.site.register(EventoAlimento)

# -------------------------------------------------------------------
# Inlines (mover para cima para serem visíveis no CustomUserAdmin)
# -------------------------------------------------------------------
class AcessoUsuarioBarInline(admin.StackedInline):
    model = AcessoUsuarioBar
    extra = 1
    verbose_name = "Acesso a Restaurantes e Bares"
    verbose_name_plural = "Acessos a Restaurantes e Bares"
    filter_horizontal = ('bares',)

class PermissaoPaginaInline(admin.TabularInline):
    model = PermissaoPagina
    extra = 1

# -------------------------------------------------------------------
# Admins básicos
# -------------------------------------------------------------------
@admin.register(Restaurante)
class RestauranteAdmin(admin.ModelAdmin):
    list_display = ('nome',)

@admin.register(Bar)
class BarAdmin(admin.ModelAdmin):
    list_display  = ('nome', 'restaurante', 'is_estoque_central')
    list_filter   = ('restaurante', 'is_estoque_central')
    search_fields = ('nome', 'restaurante__nome')  # necessário p/ autocomplete em EstoqueBar

# -------------------------------------------------------------------
# Action form (precisa herdar de ActionForm)
# -------------------------------------------------------------------
class AddProdutosToBarsActionForm(ActionForm):
    bares = forms.ModelMultipleChoiceField(
        queryset=Bar.objects.order_by('restaurante__nome', 'nome'),
        required=True,
        label='Adicionar aos bares'
    )
    quantidade_garrafas = forms.DecimalField(
        max_digits=10, decimal_places=2, required=False, initial=0, label='Qtd. garrafas (opcional)'
    )
    quantidade_doses = forms.DecimalField(
        max_digits=10, decimal_places=2, required=False, initial=0, label='Qtd. doses (opcional)'
    )

# -------------------------------------------------------------------
# ProdutoAdmin: inclui action para criar EstoqueBar em massa
# -------------------------------------------------------------------
@admin.register(Produto)
class ProdutoAdmin(admin.ModelAdmin):
    list_display  = ('nome', 'unidade_medida', 'categoria', 'ativo')
    list_filter   = ('categoria', 'ativo')
    search_fields = ('nome', 'codigo')
    ordering      = ('nome',)

    actions = ['adicionar_aos_bares']
    action_form = AddProdutosToBarsActionForm  # formulário acima do select de ações

    @admin.action(description='Adicionar produtos selecionados aos bares…')
    def adicionar_aos_bares(self, request, queryset):
        bar_ids = request.POST.getlist('bares')
        if not bar_ids:
            messages.error(request, "Selecione pelo menos um bar no formulário acima.")
            return

        try:
            q_g = float(request.POST.get('quantidade_garrafas') or 0)
            q_d = float(request.POST.get('quantidade_doses') or 0)
        except ValueError:
            messages.error(request, "Quantidades inválidas.")
            return

        bares = Bar.objects.filter(pk__in=bar_ids)
        produtos = list(queryset)

        existentes = set(
            EstoqueBar.objects
            .filter(bar__in=bares, produto__in=produtos)
            .values_list('bar_id', 'produto_id')
        )

        criar = []
        for b in bares:
            for p in produtos:
                if (b.id, p.id) not in existentes:
                    criar.append(EstoqueBar(
                        bar=b, produto=p,
                        quantidade_garrafas=q_g or 0,
                        quantidade_doses=q_d or 0
                    ))

        with transaction.atomic():
            if criar:
                EstoqueBar.objects.bulk_create(criar, ignore_conflicts=True, batch_size=2000)

        messages.success(
            request,
            f"Criados {len(criar)} estoques (Produtos: {len(produtos)} | Bares: {bares.count()})."
        )

# -------------------------------------------------------------------
# Recebimento de estoque
# -------------------------------------------------------------------
@admin.register(RecebimentoEstoque)
class RecebimentoEstoqueAdmin(admin.ModelAdmin):
    list_display = ('produto', 'quantidade', 'bar', 'restaurante', 'data_recebimento')
    list_filter = ('restaurante', 'bar', 'produto')
    search_fields = ('produto__nome', 'bar__nome', 'restaurante__nome')

# -------------------------------------------------------------------
# Transferência entre bares
# -------------------------------------------------------------------
@admin.register(TransferenciaBar)
class TransferenciaBarAdmin(admin.ModelAdmin):
    list_display = ('produto', 'quantidade', 'origem', 'destino', 'restaurante', 'usuario', 'data_transferencia')
    list_filter = ('restaurante', 'produto', 'origem', 'destino', 'usuario')
    search_fields = ('produto__nome', 'origem__nome', 'destino__nome', 'usuario__username')

# -------------------------------------------------------------------
# Contagem de bar
# -------------------------------------------------------------------
@admin.register(ContagemBar)
class ContagemBarAdmin(admin.ModelAdmin):
    list_display = (
        'bar', 'produto', 'quantidade_garrafas_cheias',
        'quantidade_doses_restantes', 'usuario', 'data_contagem'
    )
    list_filter = ('bar', 'produto', 'usuario')
    search_fields = ('bar__nome', 'produto__nome', 'usuario__username')

# -------------------------------------------------------------------
# Requisição de produto
# -------------------------------------------------------------------
@admin.register(RequisicaoProduto)
class RequisicaoProdutoAdmin(admin.ModelAdmin):
    list_display = ('produto', 'quantidade_solicitada', 'bar', 'restaurante', 'usuario', 'status', 'data_solicitacao')
    list_filter  = ('restaurante', 'bar', 'status')
    search_fields = ('produto__nome', 'bar__nome', 'usuario__username')
    actions = ['aprovar_requisicao', 'negar_requisicao']

    @admin.action(description="Aprovar requisições selecionadas")
    def aprovar_requisicao(self, request, queryset):
        queryset.update(status='APROVADA')

    @admin.action(description="Negar requisições selecionadas")
    def negar_requisicao(self, request, queryset):
        queryset.update(status='NEGADA')

# -------------------------------------------------------------------
# Filtro lateral "Por produto" em ordem alfabética garantida
# -------------------------------------------------------------------
class ProdutoOrdenadoListFilter(admin.RelatedFieldListFilter):
    def field_choices(self, field, request, model_admin):
        qs = Produto.objects.order_by('nome').only('id', 'nome')
        return [(obj.pk, str(obj)) for obj in qs]

# -------------------------------------------------------------------
# Estoque de bar
# -------------------------------------------------------------------
@admin.register(EstoqueBar)
class EstoqueBarAdmin(admin.ModelAdmin):
    list_display = ('bar', 'produto', 'quantidade_garrafas', 'quantidade_doses')
    list_filter  = (
        'bar',
        ('produto', ProdutoOrdenadoListFilter),  # filtro “Por produto” A→Z
    )
    search_fields = ('bar__nome', 'produto__nome')
    actions = ['criar_estoques_faltantes']
    autocomplete_fields = ('bar', 'produto')  # formulário com busca

    @admin.action(description="Criar estoques faltantes para todos os bares/produtos")
    def criar_estoques_faltantes(self, request, queryset):
        existentes = set(EstoqueBar.objects.values_list('bar_id', 'produto_id'))
        a_criar = []
        for b in Bar.objects.only('id'):
            for p in Produto.objects.filter(ativo=True).only('id'):
                if (b.id, p.id) not in existentes:
                    a_criar.append(EstoqueBar(bar_id=b.id, produto_id=p.id))
        if not a_criar:
            messages.info(request, "Nada a criar.")
            return
        with transaction.atomic():
            EstoqueBar.objects.bulk_create(a_criar, ignore_conflicts=True, batch_size=2000)
        messages.success(request, f"Criados {len(a_criar)} registros.")

# -------------------------------------------------------------------
# Eventos
# -------------------------------------------------------------------
class EventoProdutoInline(admin.TabularInline):
    model = EventoProduto
    extra = 0

@admin.register(Evento)
class EventoAdmin(admin.ModelAdmin):
    list_display = ['nome', 'data_criacao', 'responsavel']
    inlines = [EventoProdutoInline]

# -------------------------------------------------------------------
# User com inlines de acesso e permissões
#   - Inlines aparecem somente no change_view por padrão
#   - response_add redireciona para change_view assim que o usuário é criado
# -------------------------------------------------------------------
admin.site.unregister(User)

class CustomUserAdmin(UserAdmin):
    inlines = [AcessoUsuarioBarInline, PermissaoPaginaInline]

    def response_add(self, request, obj, post_url_continue=None):
        """
        Após criar o usuário, redireciona para a tela de edição (change_view)
        para já exibir os inlines de Acesso e Permissão de Páginas.
        """
        change_url = reverse('admin:auth_user_change', args=[obj.pk])
        messages.info(request, "Usuário criado. Agora você pode definir acessos e permissões.")
        return HttpResponseRedirect(change_url)

admin.site.register(User, CustomUserAdmin)

# -------------------------------------------------------------------
# Perdas de produto
# -------------------------------------------------------------------
@admin.register(PerdaProduto)
class PerdaProdutoAdmin(admin.ModelAdmin):
    list_display = ('data_registro', 'bar', 'produto', 'garrafas', 'doses', 'motivo', 'usuario')
    list_filter  = ('bar', 'motivo', 'data_registro')
    search_fields = ('produto__nome', 'produto__codigo', 'bar__nome', 'usuario__username')
