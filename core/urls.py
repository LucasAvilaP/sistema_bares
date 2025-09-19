from django.urls import path
from . import views

urlpatterns = [
    # Auth / Dashboard
    path('login/', views.login_view, name='login'),
    path('logout/', views.logout_view, name='logout'),
    path('dashboard/', views.dashboard, name='dashboard'),

    # Seleção / Contagens
    path('selecionar-bar/', views.selecionar_bar_view, name='selecionar-bar'),
    path('contagem/', views.contagem_view, name='contagem'),
    path('contagens/historico/', views.historico_contagens_view, name='historico-contagens'),

    # Perdas
    path('perdas/', views.pagina_perdas, name='pagina_perdas'),
    path('perdas/registrar', views.registrar_perda, name='registrar_perda'),
    path('perdas/excluir/<int:perda_id>', views.excluir_perda, name='excluir_perda'),

    # Requisições
    path('requisicao/', views.requisicao_produtos_view, name='requisicao'),
    path('requisicoes/aprovacao/', views.aprovar_requisicoes_view, name='aprovar-requisicoes'),
    path('requisicoes/historico/', views.historico_requisicoes_view, name='historico-requisicoes'),

    # Entrada / Transferências
    path('entrada-mercadorias/', views.entrada_mercadorias_view, name='entrada-mercadorias'),
    path('historico/entradas/', views.historico_entradas_view, name='historico-entradas'),
    path('transferencias/entre-bares/', views.transferencia_entre_bares_view, name='transferencia-bares'),
    path('transferencias/historico/', views.historico_transferencias_view, name='historico_transferencias'),

    #Importação
    path('importacao/', views.assistente_importacao, name="assistente_importacao"),


    # -------------------- Eventos (novo fluxo) --------------------
    path('eventos/', views.pagina_eventos, name='pagina_eventos'),
    path('eventos/criar/', views.criar_evento, name='criar_evento'),
    path('eventos/<int:evento_id>/editar/', views.editar_evento, name='editar_evento'),
    path('eventos/<int:evento_id>/excluir/', views.excluir_evento, name='excluir_evento'),
    path('eventos/exportar_excel/', views.exportar_consolidado_eventos_excel, name='exportar_consolidado_eventos_excel'),

    # (legado) manter compatibilidade com formulários antigos
    path('eventos/salvar/', views.criar_evento, name='salvar_evento'),

    # -------------------- Relatórios --------------------
    path('relatorios/', views.relatorios_view, name='relatorios'),
    path('relatorios/saida-estoque/', views.relatorio_saida_estoque, name='relatorio_saida_estoque'),
    path('relatorios/saida-estoque/exportar/', views.exportar_saida_estoque_excel, name='exportar_saida_estoque_excel'),
    path('relatorios/consolidado-diferenca/', views.relatorio_consolidado_view, name='relatorio-consolidado'),
    path('relatorios/consolidado-diferenca/excel/', views.relatorio_consolidado_excel_view, name='relatorio_consolidado_excel'),
    path('relatorios/contagem-atual/', views.relatorio_contagem_atual, name='relatorio_contagem_atual'),
    path('relatorios/contagem-atual/excel/', views.exportar_contagem_atual_excel, name='exportar_contagem_atual_excel'),
    path('relatorios/diferenca-contagens/', views.relatorio_diferenca_contagens, name='relatorio_diferenca_contagens'),
    path('relatorios/diferenca-contagens/exportar-excel/', views.exportar_diferenca_contagens_excel, name='exportar_diferenca_contagens_excel'),
    path('relatorios/perdas/', views.relatorio_perdas, name='relatorio_perdas'),
    path("relatorios/perdas/exportar/", views.exportar_relatorio_perdas_excel, name="exportar_relatorio_perdas_excel"),
    path('relatorios/perdas/marcar/<int:perda_id>/', views.marcar_perda_baixada, name='marcar_perda_baixada'),
    path('relatorios/perdas/desmarcar/<int:perda_id>/', views.desmarcar_perda_baixada, name='desmarcar_perda_baixada'),
    path(
        "relatorios/consolidado/periodo/",
        views.relatorio_consolidado_periodo,
        name="relatorio_consolidado_periodo",
    ),
    # Exportação Excel
    path(
        "relatorios/consolidado/periodo/exportar/",
        views.exportar_consolidado_periodo_excel,
        name="exportar_consolidado_periodo_excel",
    ),
    # (opcional) Consolidado atual
    path(
        "relatorios/consolidado/atual/",
        views.consolidado_atual_view,
        name="consolidado_atual",
    ),
    path(
    "relatorios/consolidado/atual/exportar/",
    views.exportar_consolidado_atual_excel,
    name="exportar_consolidado_atual_excel",
),
    


    # Relatório específico de eventos (permanece)
    path('relatorio_eventos/', views.relatorio_eventos, name='relatorio_eventos'),
    path('relatorio_eventos/exportar_excel/', views.exportar_relatorio_eventos_excel, name='exportar_relatorio_eventos_excel'),
    path('relatorios/eventos/<int:evento_id>/baixar/', views.marcar_evento_baixado, name='marcar_evento_baixado'),
    path('relatorios/eventos/<int:evento_id>/desmarcar/', views.desmarcar_evento_baixado, name='desmarcar_evento_baixado'),
]
