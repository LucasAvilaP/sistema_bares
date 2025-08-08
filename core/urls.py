from django.urls import path
from . import views




urlpatterns = [
    path('login/', views.login_view, name='login'),
    path('logout/', views.logout_view, name='logout'),
    path('selecionar-bar/', views.selecionar_bar_view, name='selecionar-bar'),
    path('contagem/', views.contagem_view, name='contagem'),
    path('contagens/historico/', views.historico_contagens_view, name='historico-contagens'),
    path('requisicao/', views.requisicao_produtos_view, name='requisicao'),
    path('entrada-mercadorias/', views.entrada_mercadorias_view, name='entrada-mercadorias'),
    path('requisicoes/aprovacao/', views.aprovar_requisicoes_view, name='aprovar-requisicoes'),
    path('requisicoes/historico/', views.historico_requisicoes_view, name='historico-requisicoes'),
    path('transferencias/entre-bares/', views.transferencia_entre_bares_view, name='transferencia-bares'),
    path('transferencias/historico/', views.historico_transferencias_view, name='historico_transferencias'),
    path('eventos/', views.pagina_eventos, name='pagina_eventos'),
    path('eventos/salvar/', views.salvar_evento, name='salvar_evento'),
    path('eventos/excluir/<int:evento_id>/', views.excluir_evento, name='excluir_evento'),
    path('eventos/exportar_excel/', views.exportar_consolidado_eventos_excel, name='exportar_consolidado_eventos_excel'),
    path('relatorio_eventos/exportar_excel/', views.exportar_relatorio_eventos_excel, name='exportar_relatorio_eventos_excel'),
    path('relatorio_eventos/', views.relatorio_eventos, name='relatorio_eventos'),
    path('relatorios/', views.relatorios_view, name='relatorios'),
    path('relatorios/saida-estoque/', views.relatorio_saida_estoque, name='relatorio_saida_estoque'),
    path('relatorios/saida-estoque/exportar/', views.exportar_saida_estoque_excel, name='exportar_saida_estoque_excel'),
    path('relatorios/consolidado-diferenca/', views.relatorio_consolidado_view, name='relatorio-consolidado'),
    path('relatorios/consolidado-diferenca/excel/', views.relatorio_consolidado_excel_view, name='relatorio_consolidado_excel'),
    path('relatorios/contagem-atual/', views.relatorio_contagem_atual, name='relatorio_contagem_atual'),
    path('relatorios/contagem-atual/excel/', views.exportar_contagem_atual_excel, name='exportar_contagem_atual_excel'),
    path('relatorios/diferenca-contagens/', views.relatorio_diferenca_contagens, name='relatorio_diferenca_contagens'),
    path('relatorios/diferenca-contagens/exportar-excel/', views.exportar_diferenca_contagens_excel, name='exportar_diferenca_contagens_excel'),
    path('dashboard/', views.dashboard, name='dashboard'),
]
