from django.shortcuts import redirect
from django.urls import resolve

# paths que NÃO exigem contexto restaurante/bar:
ALLOWLIST = {
    'login', 'logout',
    'pagina_eventos', 'criar_evento', 'editar_evento', 'excluir_evento',
    'exportar_consolidado_eventos_excel',
    'relatorio_eventos', 'exportar_relatorio_eventos_excel',
    'marcar_evento_baixado', 'desmarcar_evento_baixado',
    # estáticos / dashboard neutro (se optar)
    'relatorios',
}

# prefixes que você também quer liberar (ex.: /admin/ ou /static/) – opcional
ALLOW_PREFIXES = ('/admin/', '/static/', '/media/')

class ContextGuardMiddleware:
    """
    Se a view NÃO estiver na allowlist e exigir restaurante/bar,
    verifica sessão: se faltar, redireciona para 'selecionar-bar'.
    """
    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        path = request.path or ''

        # libera prefixos
        if path.startswith(ALLOW_PREFIXES):
            return self.get_response(request)

        # resolve nome da URL (se não resolver, segue normalmente)
        try:
            match = resolve(path)
            view_name = match.url_name
        except Exception:
            return self.get_response(request)

        # Se a rota está liberada, segue.
        if view_name in ALLOWLIST:
            return self.get_response(request)

        # Daqui pra baixo: rotas protegidas por contexto
        restaurante_id = request.session.get('restaurante_id')
        bar_id = request.session.get('bar_id')

        # Ajuste: algumas rotas exigem só restaurante, outras restaurante+bar.
        # Para simplificar, exigimos pelo menos restaurante,
        # e para páginas de contagem/transferência exigimos também bar.
        requires_bar = view_name in {
            'contagem', 'historico-contagens', 'historico-entradas',
            'transferencia-bares', 'historico_transferencias',
            'entrada_mercadorias', 'requisicao', 'aprovar-requisicoes',
            'relatorio_saida_estoque', 'exportar_saida_estoque_excel',
            'relatorio_contagem_atual', 'exportar_contagem_atual_excel',
            'relatorio_diferenca_contagens', 'exportar_diferenca_contagens_excel',
            'relatorio_perdas', 'exportar_relatorio_perdas_excel',
            'pagina_perdas', 'registrar_perda', 'excluir_perda',
            'dashboard',
        }

        if not restaurante_id or (requires_bar and not bar_id):
            # guarda a rota pretendida para voltar depois
            request.session['next_after_select'] = path
            return redirect('selecionar-bar')

        return self.get_response(request)
