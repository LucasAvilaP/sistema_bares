from django.contrib import admin, messages
from .models import (
    Restaurante, Bar, Produto, RecebimentoEstoque,
    TransferenciaBar, ContagemBar, RequisicaoProduto,
    EstoqueBar, AcessoUsuarioBar, Evento, EventoProduto, PermissaoPagina, Alimento, EventoAlimento, PerdaProduto
)
from django.contrib.auth.models import User
from django.contrib.auth.admin import UserAdmin
from django.db import transaction

admin.site.register(Alimento)
admin.site.register(EventoAlimento)

# Inline para associar bares ao usuário
class AcessoUsuarioBarInline(admin.StackedInline):
    model = AcessoUsuarioBar
    extra = 1
    verbose_name = "Acesso a Restaurantes e Bares"
    verbose_name_plural = "Acessos a Restaurantes e Bares"
    filter_horizontal = ('bares',)


# Admin Restaurante
@admin.register(Restaurante)
class RestauranteAdmin(admin.ModelAdmin):
    list_display = ('nome',)

# Admin Bar
@admin.register(Bar)
class BarAdmin(admin.ModelAdmin):
    list_display = ('nome', 'restaurante', 'is_estoque_central')
    list_filter = ('restaurante', 'is_estoque_central')

# Admin Produto
@admin.register(Produto)
class ProdutoAdmin(admin.ModelAdmin):
    list_display = ('nome', 'unidade_medida', 'categoria', 'ativo')
    list_filter = ('categoria', 'ativo')
    search_fields = ('nome',)

# Admin Recebimento de Estoque
@admin.register(RecebimentoEstoque)
class RecebimentoEstoqueAdmin(admin.ModelAdmin):
    list_display = ('produto', 'quantidade', 'bar', 'restaurante', 'data_recebimento')
    list_filter = ('restaurante', 'bar', 'produto')
    search_fields = ('produto__nome', 'bar__nome', 'restaurante__nome')

# Admin Transferência entre Bares
@admin.register(TransferenciaBar)
class TransferenciaBarAdmin(admin.ModelAdmin):
    list_display = ('produto', 'quantidade', 'origem', 'destino', 'restaurante', 'usuario', 'data_transferencia')
    list_filter = ('restaurante', 'produto', 'origem', 'destino', 'usuario')
    search_fields = ('produto__nome', 'origem__nome', 'destino__nome', 'usuario__username')

# Admin Contagem de Bar
@admin.register(ContagemBar)
class ContagemBarAdmin(admin.ModelAdmin):
    list_display = (
        'bar', 'produto', 'quantidade_garrafas_cheias',
        'quantidade_doses_restantes', 'usuario', 'data_contagem'
    )
    list_filter = ('bar', 'produto', 'usuario')
    search_fields = ('bar__nome', 'produto__nome', 'usuario__username')

# Admin Requisição de Produto
@admin.register(RequisicaoProduto)
class RequisicaoProdutoAdmin(admin.ModelAdmin):
    list_display = ('produto', 'quantidade_solicitada', 'bar', 'restaurante', 'usuario', 'status', 'data_solicitacao')
    list_filter = ('restaurante', 'bar', 'status')
    search_fields = ('produto__nome', 'bar__nome', 'usuario__username')
    actions = ['aprovar_requisicao', 'negar_requisicao']

    def aprovar_requisicao(self, request, queryset):
        queryset.update(status='APROVADA')
    aprovar_requisicao.short_description = "Aprovar requisições selecionadas"

    def negar_requisicao(self, request, queryset):
        queryset.update(status='NEGADA')
    negar_requisicao.short_description = "Negar requisições selecionadas"

# Admin Estoque de Bar
@admin.register(EstoqueBar)
class EstoqueBarAdmin(admin.ModelAdmin):
    list_display = ('bar', 'produto', 'quantidade_garrafas', 'quantidade_doses')
    list_filter  = ('bar', 'produto')
    search_fields = ('bar__nome', 'produto__nome')
    actions = ['criar_estoques_faltantes']

    @admin.action(description="Criar estoques faltantes para todos os bares/produtos")
    def criar_estoques_faltantes(self, request, queryset):
        existentes = set(EstoqueBar.objects.values_list('bar_id', 'produto_id'))
        a_criar = []
        for b in Bar.objects.only('id'):
            for p in Produto.objects.filter(ativo=True).only('id'):
                if (b.id, p.id) not in existentes:
                    a_criar.append(EstoqueBar(bar_id=b.id, produto_id=p.id))
        if not a_criar:
            self.message_user(request, "Nada a criar.", level=messages.INFO)
            return
        with transaction.atomic():
            EstoqueBar.objects.bulk_create(a_criar, ignore_conflicts=True, batch_size=2000)
        self.message_user(request, f"Criados {len(a_criar)} registros.", level=messages.SUCCESS)




class EventoProdutoInline(admin.TabularInline):
    model = EventoProduto
    extra = 0

@admin.register(Evento)
class EventoAdmin(admin.ModelAdmin):
    list_display = ['nome', 'data_criacao', 'responsavel']
    inlines = [EventoProdutoInline]


class PermissaoPaginaInline(admin.TabularInline):
    model = PermissaoPagina
    extra = 1

# Desregistrar o admin original
admin.site.unregister(User)

class CustomUserAdmin(UserAdmin):
    inlines = [AcessoUsuarioBarInline, PermissaoPaginaInline]

# Registre novamente o User com os dois inlines
admin.site.register(User, CustomUserAdmin)



@admin.register(PerdaProduto)
class PerdaProdutoAdmin(admin.ModelAdmin):
    list_display = ('data_registro','bar','produto','garrafas','doses','motivo','usuario')
    list_filter = ('bar','motivo','data_registro')
    search_fields = ('produto__nome','produto__codigo','bar__nome','usuario__username')