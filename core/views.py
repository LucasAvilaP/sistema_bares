from django.contrib.auth import authenticate, login, logout
from django.shortcuts import render, redirect, get_object_or_404
from .models import Produto, Bar, Restaurante, RequisicaoProduto, TransferenciaBar, ContagemBar, EstoqueBar, models, AcessoUsuarioBar, EventoProduto, Evento, PermissaoPagina, RecebimentoEstoque
from decimal import Decimal, InvalidOperation
from django.contrib import messages
from django.contrib.auth.decorators import login_required, user_passes_test
import openpyxl
import io
from io import BytesIO
import xlsxwriter
from django.utils.timezone import is_aware
from openpyxl.styles import Font
from django.utils import timezone
from django.db.models import Sum, Max, Q, Count
from openpyxl.utils import get_column_letter
from django.http import HttpResponse
from openpyxl import Workbook
from django.contrib.auth.decorators import user_passes_test
from django.db.models import DateField
from django.db.models.functions import TruncDate
from collections import defaultdict, OrderedDict
from django.utils.timezone import now
from django.utils.dateparse import parse_date
from babel.dates import parse_date
from datetime import datetime, date

def login_view(request):
    context = {}

    if request.method == 'POST':
        username = request.POST.get('username')
        password = request.POST.get('password')
        restaurante_id = request.POST.get('restaurante')

        user = authenticate(request, username=username, password=password)
        if user is not None:
            login(request, user)

            # Verifica se esse restaurante é permitido para esse usuário
            acesso_valido = AcessoUsuarioBar.objects.filter(user=user, restaurante_id=restaurante_id).exists()
            if not acesso_valido:
                messages.error(request, 'Você não tem acesso a este restaurante.')
                return redirect('login')

            request.session['restaurante_id'] = restaurante_id
            return redirect('selecionar-bar')

        else:
            context['erro'] = 'Usuário ou senha inválidos'

    context['restaurantes'] = Restaurante.objects.all()  # Mostrar todos os restaurantes para dropdown
    return render(request, 'core/login.html', context)


def logout_view(request):
    logout(request)

    # Limpa manualmente dados da sessão, se necessário
    request.session.flush()

    return redirect('login')


@login_required
def selecionar_bar_view(request):
    restaurante_id = request.session.get('restaurante_id')

    if not restaurante_id:
        messages.error(request, "Restaurante não selecionado.")
        return redirect('login')

    restaurante = get_object_or_404(Restaurante, id=restaurante_id)

    acessos = AcessoUsuarioBar.objects.filter(user=request.user, restaurante_id=restaurante_id)

    # Correção aqui: ManyToMany → bares__id
    bares_ids = acessos.values_list('bares__id', flat=True)
    bares = Bar.objects.filter(id__in=bares_ids)

    if request.method == 'POST':
        bar_id = request.POST.get('bar')
        if bar_id and bares.filter(id=bar_id).exists():
            bar = Bar.objects.get(id=bar_id)
            request.session['bar_id'] = bar.id
            request.session['bar_nome'] = bar.nome
            return redirect('dashboard')

        else:
            messages.error(request, "Bar inválido ou sem permissão.")

    return render(request, 'core/selecionar_bar.html', {
        'bares': bares,
        'restaurante': restaurante
    })


@login_required
def requisicao_produtos_view(request):
    if not PermissaoPagina.objects.filter(user=request.user, nome_pagina='requisicoes').exists():
        messages.error(request, 'Você não tem permissão para acessar essa página.')
        return redirect('dashboard')

    if request.method == 'POST':
        produtos = request.POST.getlist('produto[]')
        quantidades = request.POST.getlist('quantidade[]')

        restaurante_id = request.session.get('restaurante_id')
        bar_id = request.session.get('bar_id')

        restaurante = Restaurante.objects.get(id=restaurante_id)
        bar = Bar.objects.get(id=bar_id)

        try:
            estoque_central = Bar.objects.get(restaurante=restaurante, is_estoque_central=True)
        except Bar.DoesNotExist:
            messages.error(request, 'Estoque central não encontrado para o restaurante.')
            return redirect('requisicao')

        produtos_requisitados = []
        erros = []

        for prod_id, qtd in zip(produtos, quantidades):
            produto = Produto.objects.get(id=prod_id)

            try:
                qtd_solicitada = Decimal(qtd.replace(",", ".")) if qtd else Decimal("0")
            except:
                qtd_solicitada = Decimal("0")

            estoque = EstoqueBar.objects.filter(bar=estoque_central, produto=produto).first()

            if not estoque or estoque.quantidade_garrafas < qtd_solicitada:
                erros.append(f"❌ {produto.nome} - Estoque insuficiente no estoque central.")
                continue

            # Cria a requisição
            RequisicaoProduto.objects.create(
                restaurante=restaurante,
                bar=bar,  # ← bar solicitante
                produto=produto,
                quantidade_solicitada=qtd_solicitada,
                usuario=request.user
            )
            produtos_requisitados.append(produto.nome)

        if produtos_requisitados:
            messages.success(
                request,
                f"Requisição enviada para: {', '.join(produtos_requisitados)}"
            )

        if erros:
            for erro in erros:
                messages.error(request, erro)

        return redirect('requisicao')

    produtos = Produto.objects.filter(ativo=True).order_by('nome')
    return render(request, 'core/requisicao.html', {'produtos': produtos})




@login_required
def entrada_mercadorias_view(request):
    if not PermissaoPagina.objects.filter(user=request.user, nome_pagina='entrada_mercadoria').exists():
        messages.error(request, 'Você não tem permissão para acessar essa página.')
        return redirect('dashboard')
    restaurante_id = request.session.get('restaurante_id')
    restaurante = Restaurante.objects.get(id=restaurante_id)

    try:
        estoque_central = Bar.objects.get(restaurante=restaurante, is_estoque_central=True)
    except Bar.DoesNotExist:
        messages.error(request, "Estoque central não definido para este restaurante.")
        return redirect('/dashboard/')

    if request.method == 'POST':
        produtos = request.POST.getlist('produto[]')
        quantidades = request.POST.getlist('quantidade[]')

        for prod_id, qtd in zip(produtos, quantidades):
            produto = Produto.objects.get(id=prod_id)
            quantidade = Decimal(qtd.replace(',', '.')) if qtd else Decimal(0)

            # Salva o registro da entrada
            RecebimentoEstoque.objects.create(
                restaurante=restaurante,
                bar=estoque_central,
                produto=produto,
                quantidade=quantidade
            )

            # Atualiza o estoque
            EstoqueBar.adicionar(estoque_central, produto, quantidade)

        messages.success(request, "Entrada de mercadorias realizada com sucesso!")
        return redirect('entrada-mercadorias')

    produtos = Produto.objects.filter(ativo=True).order_by('nome')
    return render(request, 'core/entrada_mercadorias.html', {'produtos': produtos})


@login_required
def contagem_view(request):
    # 🔒 Verificação de permissão para "contagem"
    if not PermissaoPagina.objects.filter(user=request.user, nome_pagina='contagem').exists():
        messages.error(request, "Você não tem permissão para acessar a página de contagem.")
        return redirect('dashboard')

    bar_id = request.session.get('bar_id')
    restaurante_id = request.session.get('restaurante_id')
    bar = Bar.objects.get(id=bar_id)
    restaurante = Restaurante.objects.get(id=restaurante_id)

    if request.method == 'POST':
        produtos = request.POST.getlist('produto[]')
        garrafas = request.POST.getlist('garrafas[]')
        doses = request.POST.getlist('doses[]')

        for prod_id, qtd_garrafas, qtd_doses in zip(produtos, garrafas, doses):
            produto = Produto.objects.get(id=prod_id)

            try:
                garrafas_valor = int(qtd_garrafas) if qtd_garrafas else 0
            except ValueError:
                garrafas_valor = 0

            try:
                doses_valor = Decimal(qtd_doses.replace(",", ".")) if qtd_doses else Decimal("0")
            except (InvalidOperation, ValueError):
                doses_valor = Decimal("0")

            # 1️⃣ Cria o registro da contagem (como já fazia)
            ContagemBar.objects.create(
                bar=bar,
                produto=produto,
                quantidade_garrafas_cheias=garrafas_valor,
                quantidade_doses_restantes=doses_valor,
                usuario=request.user
            )

            # 2️⃣ Atualiza o EstoqueBar com os valores da contagem
            estoque, _ = EstoqueBar.objects.get_or_create(bar=bar, produto=produto)
            estoque.quantidade_garrafas = garrafas_valor
            estoque.quantidade_doses = doses_valor
            estoque.save()

        messages.success(request, "Contagem registrada e estoque atualizado com sucesso!")
        return redirect('contagem')

    produtos = Produto.objects.filter(ativo=True).order_by('nome')
    return render(request, 'core/contagem.html', {'produtos': produtos})




@login_required
def historico_contagens_view(request):
    # 🔒 Verificação de permissão para "historico_contagens"
    if not PermissaoPagina.objects.filter(user=request.user, nome_pagina='historico_cont').exists():
        messages.error(request, "Você não tem permissão para acessar a página de Historico de contagens.")
        return redirect('dashboard')
    
    bar_id = request.session.get('bar_id')
    bar = Bar.objects.get(id=bar_id)

    mes = request.GET.get('mes')
    ano = request.GET.get('ano')

    contagens = ContagemBar.objects.filter(bar=bar)

    if mes and ano:
        contagens = contagens.filter(data_contagem__month=mes, data_contagem__year=ano)
    else:
        contagens = contagens.annotate(data=TruncDate('data_contagem')).order_by('-data_contagem')

    agrupado = {}
    for c in contagens:
        data = c.data_contagem.date()
        if mes and ano:
            agrupado.setdefault(data, []).append(c)
        else:
            if data not in agrupado:
                if len(agrupado) >= 10:
                    break
                agrupado[data] = []
            agrupado[data].append(c)

    return render(request, 'core/historico_contagens.html', {
        'agrupado': agrupado,
        'now': now(),
        'meses': list(range(1, 13))  # de 1 a 12
    })





@login_required
def aprovar_requisicoes_view(request):
    # 🔒 Verificação de permissão para "aprovacao"
    if not PermissaoPagina.objects.filter(user=request.user, nome_pagina='aprovacao').exists():
        messages.error(request, "Você não tem permissão para acessar a página de Aprovar/Rejeitar.")
        return redirect('dashboard')  # ou outra página padrão
    restaurante_id = request.session.get('restaurante_id')
    requisicoes = RequisicaoProduto.objects.filter(restaurante_id=restaurante_id, status='PENDENTE')
    restaurante = get_object_or_404(Restaurante, id=restaurante_id)
    bar_central = get_object_or_404(Bar, restaurante=restaurante, is_estoque_central=True)

    if request.method == 'POST':
        for key in request.POST:
            if key.startswith('aprovacao_'):
                req_id = key.split('_')[1]
                decisao = request.POST[key]
                requisicao = get_object_or_404(RequisicaoProduto, id=req_id)

                if decisao == 'aprovar':
                    qtd = Decimal(requisicao.quantidade_solicitada)

                    sucesso = EstoqueBar.retirar(bar_central, requisicao.produto, qtd)

                    if sucesso:
                        EstoqueBar.adicionar(requisicao.bar, requisicao.produto, qtd)

                        TransferenciaBar.objects.create(
                            restaurante=requisicao.restaurante,
                            origem=bar_central,
                            destino=requisicao.bar,
                            produto=requisicao.produto,
                            quantidade=qtd,
                            usuario=request.user
                        )

                        requisicao.status = 'APROVADA'
                        requisicao.usuario_aprovador = request.user
                        messages.success(request, f"Requisição aprovada com sucesso.")
                    else:
                        requisicao.status = 'FALHA_ESTOQUE'
                        requisicao.usuario_aprovador = request.user
                        messages.warning(request, f"Produto '{requisicao.produto.nome}' insuficiente no estoque central. Requisição {req_id} não aprovada.")

                    requisicao.save()

                elif decisao == 'negar':
                    requisicao.status = 'NEGADA'
                    requisicao.usuario_aprovador = request.user
                    requisicao.save()
                    messages.info(request, f"Requisição foi negada.")

        return redirect('aprovar-requisicoes')

    return render(request, 'core/aprovar_requisicoes.html', {'requisicoes': requisicoes})






def atualizar_estoque(bar, produto, quantidade_delta):
    estoque, _ = EstoqueBar.objects.get_or_create(bar=bar, produto=produto)
    estoque.quantidade += quantidade_delta
    estoque.save()



@login_required
def historico_requisicoes_view(request):
    # 🔒 Verificação de permissão para "historico_requisicoes"
    if not PermissaoPagina.objects.filter(user=request.user, nome_pagina='historico_requi').exists():
        messages.error(request, "Você não tem permissão para acessar a página de Historico de requisições.")
        return redirect('dashboard')
    
    bar_id = request.session.get('bar_id')
    restaurante_id = request.session.get('restaurante_id')

    filtro_ativo = False
    agrupado = defaultdict(list)

    if not bar_id or not restaurante_id:
        return render(request, 'core/historico_requisicoes.html', {
            'agrupado': agrupado,
            'now': datetime.now(),
            'filtro_ativo': filtro_ativo
        })

    mes = request.GET.get('mes')
    ano = request.GET.get('ano')

    requisicoes = RequisicaoProduto.objects.filter(
        restaurante_id=restaurante_id,
        bar_id=bar_id
    )

    if mes and ano:
        try:
            mes = int(mes)
            ano = int(ano)
            requisicoes = requisicoes.filter(
                data_solicitacao__month=mes,
                data_solicitacao__year=ano
            )
            filtro_ativo = True
        except ValueError:
            pass
    else:
        requisicoes = requisicoes.order_by('-data_solicitacao')[:20]

    requisicoes = requisicoes.annotate(
        data_truncada=TruncDate('data_solicitacao')
    )

    for r in requisicoes:
        agrupado[r.data_truncada].append(r)

    return render(request, 'core/historico_requisicoes.html', {
        'agrupado': dict(agrupado),
        'now': datetime.now(),
        'filtro_ativo': filtro_ativo
    })





from decimal import Decimal, InvalidOperation

@login_required
def transferencia_entre_bares_view(request):
    # 🔒 Verificação de permissão
    if not PermissaoPagina.objects.filter(user=request.user, nome_pagina='transferencias').exists():
        messages.error(request, "Você não tem permissão para acessar a página de transferências.")
        return redirect('dashboard')

    restaurante_id = request.session.get('restaurante_id')
    bar_origem_id = request.session.get('bar_id')

    restaurante = Restaurante.objects.get(id=restaurante_id)
    bar_origem = Bar.objects.get(id=bar_origem_id)
    bares_destino = Bar.objects.filter(restaurante=restaurante).exclude(id=bar_origem_id)

    if request.method == 'POST':
        produto_id = request.POST.get('produto')
        bar_destino_id = request.POST.get('bar_destino')
        quantidade_str = request.POST.get('quantidade')

        # 🔎 Validação dos campos
        if not produto_id or not bar_destino_id or not quantidade_str:
            messages.error(request, "Todos os campos são obrigatórios.")
            return redirect('transferencia-bares')

        # 🔢 Tentar converter a quantidade
        try:
            quantidade = Decimal(quantidade_str.replace(',', '.'))
        except (InvalidOperation, AttributeError):
            messages.error(request, "Quantidade inválida.")
            return redirect('transferencia-bares')

        # 🔎 Verifica se produto e bar de destino existem
        try:
            produto = Produto.objects.get(id=produto_id)
            bar_destino = Bar.objects.get(id=bar_destino_id)
        except Produto.DoesNotExist:
            messages.error(request, "Produto inválido.")
            return redirect('transferencia-bares')
        except Bar.DoesNotExist:
            messages.error(request, "Bar de destino inválido.")
            return redirect('transferencia-bares')

        # 🧮 Verifica se há estoque suficiente no bar de origem
        sucesso = EstoqueBar.retirar(bar_origem, produto, quantidade)

        if not sucesso:
            messages.error(request, "Estoque insuficiente para essa transferência.")
            return redirect('transferencia-bares')

        # ➕ Adiciona ao bar de destino
        EstoqueBar.adicionar(bar_destino, produto, quantidade)

        # 📝 Registra a transferência
        TransferenciaBar.objects.create(
            restaurante=restaurante,
            origem=bar_origem,
            destino=bar_destino,
            produto=produto,
            quantidade=quantidade,
            usuario=request.user
        )

        messages.success(request, "Transferência realizada com sucesso!")
        return redirect('transferencia-bares')

    # 🧾 Carrega os produtos com estoque no bar de origem
    produtos_disponiveis = EstoqueBar.objects.filter(bar=bar_origem, quantidade_garrafas__gt=0)

    return render(request, 'core/transferencia_bares.html', {
        'bares_destino': bares_destino,
        'produtos': [e.produto for e in produtos_disponiveis]
    })






@login_required
def historico_transferencias_view(request):
    # 🔒 Verificação de permissão para "historico_transferencias"
    if not PermissaoPagina.objects.filter(user=request.user, nome_pagina='historico_transf').exists():
        messages.error(request, "Você não tem permissão para acessar a página de Historico de transferencias.")
        return redirect('dashboard')
    
    bar_id = request.session.get('bar_id')
    restaurante_id = request.session.get('restaurante_id')

    filtro_ativo = False
    agrupado = defaultdict(list)  # ✅ evita KeyError automaticamente

    if not bar_id or not restaurante_id:
        return render(request, 'core/historico_transferencias.html', {
            'agrupado': agrupado,
            'now': datetime.now(),
            'filtro_ativo': filtro_ativo
        })

    bar = Bar.objects.get(id=bar_id)
    mes = request.GET.get('mes')
    ano = request.GET.get('ano')

    # 🔍 Base query
    transferencias = TransferenciaBar.objects.filter(
        restaurante_id=restaurante_id
    ).filter(
        models.Q(origem=bar) | models.Q(destino=bar)
    )

    if mes and ano:
        try:
            mes = int(mes)
            ano = int(ano)
            transferencias = transferencias.filter(
                data_transferencia__month=mes,
                data_transferencia__year=ano
            )
            filtro_ativo = True
        except ValueError:
            pass
    else:
        transferencias = transferencias.order_by('-data_transferencia')[:10]

    # ✅ Anota data truncada para agrupar
    transferencias = transferencias.annotate(
        data_truncada=TruncDate('data_transferencia')
    )

    # ✅ Agora agrupamos sem risco de erro
    for t in transferencias:
        agrupado[t.data_truncada].append(t)

    return render(request, 'core/historico_transferencias.html', {
        'agrupado': dict(agrupado),
        'now': datetime.now(),
        'filtro_ativo': filtro_ativo
    })


@login_required
def pagina_eventos(request):
    # 🔒 Verificação de permissão
    if not PermissaoPagina.objects.filter(user=request.user, nome_pagina='eventos').exists():
        messages.error(request, "Você não tem permissão para acessar a página de eventos.")
        return redirect('dashboard')

    hoje = timezone.localdate()
    eventos = Evento.objects.filter(data_criacao__date=hoje)

    consolidado = defaultdict(lambda: {'garrafas': 0, 'doses': 0, 'ml': 0})
    for evento in eventos:
        for item in evento.produtos.all():
            consolidado[item.produto.nome]['garrafas'] += item.garrafas
            consolidado[item.produto.nome]['doses'] += item.doses
            consolidado[item.produto.nome]['ml'] += item.doses * 50  # 50 ml fixos por dose

    produtos = Produto.objects.all()

    context = {
        'produtos': produtos,
        'eventos': eventos,
        'consolidado': dict(consolidado),
    }
    return render(request, 'eventos/pagina_eventos.html', context)


@login_required
def salvar_evento(request):
    if request.method == 'POST':
        nome = request.POST.get('nome_evento')
        responsavel = request.user

        evento = Evento.objects.create(nome=nome, responsavel=responsavel)

        produtos_ids = request.POST.getlist('produto_id[]')
        garrafas_list = request.POST.getlist('garrafas[]')
        doses_list = request.POST.getlist('doses[]')

        for i in range(len(produtos_ids)):
            try:
                produto = Produto.objects.get(id=produtos_ids[i])
                garrafas = int(garrafas_list[i])
                doses = int(doses_list[i])

                if garrafas > 0 or doses > 0:
                    EventoProduto.objects.create(
                        evento=evento,
                        produto=produto,
                        garrafas=garrafas,
                        doses=doses
                    )
            except (Produto.DoesNotExist, ValueError):
                continue  # ignora entradas inválidas

        return redirect('pagina_eventos')


@login_required
def excluir_evento(request, evento_id):
    evento = get_object_or_404(Evento, id=evento_id)
    evento.delete()
    return redirect('pagina_eventos')




@login_required
def dashboard(request):
    bar_id = request.session.get('bar_id')
    bar = get_object_or_404(Bar, id=bar_id)

    # Estoque atual agrupado
    estoque_qs = EstoqueBar.objects.filter(bar=bar).select_related('produto')
    estoque_agrupado = [
        {
            'produto': e.produto.nome,
            'garrafas': e.quantidade_garrafas,
            'doses': e.quantidade_doses
        }
        for e in estoque_qs
    ]

    # Gráfico: Top 5 produtos com maior quantidade no estoque atual (garrafas + doses/10)
    estoque_top_qs = sorted(
    estoque_qs,
    key=lambda e: e.quantidade_garrafas + (e.quantidade_doses / Decimal('10')),
    reverse=True
    )[:5]

    estoque_labels = [e.produto.nome for e in estoque_top_qs]
    estoque_valores = [
        round(e.quantidade_garrafas + (e.quantidade_doses / Decimal('10')), 2)
        for e in estoque_top_qs
    ]

    # Últimas requisições e transferências
    ultimas_requisicoes = RequisicaoProduto.objects.filter(bar=bar).order_by('-data_solicitacao')[:5]
    ultimas_transferencias = TransferenciaBar.objects.filter(origem=bar).order_by('-data_transferencia')[:5]

    # 📊 Gráfico: Produtos mais requisitados
    ranking_qs = (
        RequisicaoProduto.objects
        .filter(bar=bar)
        .values('produto__nome')
        .annotate(total=Count('id'))
        .order_by('-total')[:5]
    )
    produtos_ranking = [r['produto__nome'] for r in ranking_qs]
    totais_ranking = [r['total'] for r in ranking_qs]

    return render(request, 'core/dashboard.html', {
        'bar': bar,
        'estoque': estoque_agrupado,
        'ultimas_requisicoes': ultimas_requisicoes,
        'ultimas_transferencias': ultimas_transferencias,
        'dias': estoque_labels,             # agora representa produtos do estoque
        'saidas': estoque_valores,          # agora representa quantidades (garrafas + doses/10)
        'ranking_produtos': produtos_ranking,
        'ranking_totais': totais_ranking,
    })


    



#                                                                                         SEÇÃO DE RELATÓRIOS

@login_required
def relatorios_view(request):
    # 🔒 Verificação de permissão
    if not PermissaoPagina.objects.filter(user=request.user, nome_pagina='relatorios').exists():
        messages.error(request, "Você não tem permissão para acessar a página de relatórios.")
        return redirect('dashboard')

    return render(request, 'core/relatorios.html')




@login_required
def relatorio_saida_estoque(request):
    bar_id = request.session.get('bar_id')
    if not bar_id:
        return render(request, 'erro.html', {'mensagem': 'Nenhum bar selecionado.'})

    mes = request.GET.get('mes')
    ano = request.GET.get('ano')
    produto_query = request.GET.get('produto', '').strip()

    requisicoes = RequisicaoProduto.objects.filter(
        bar_id=bar_id,
        status__in=['APROVADA', 'NEGADA', 'FALHA_ESTOQUE']
    )

    # Aplicar os filtros de produto antes de cortar
    if produto_query:
        requisicoes = requisicoes.filter(produto__nome__icontains=produto_query)

    if mes and ano:
        requisicoes = requisicoes.filter(data_solicitacao__month=int(mes), data_solicitacao__year=int(ano))
    elif mes:
        requisicoes = requisicoes.filter(data_solicitacao__month=int(mes))
    elif ano:
        requisicoes = requisicoes.filter(data_solicitacao__year=int(ano))

    # Se nenhum filtro de data foi usado, limita a 50 últimas
    if not mes and not ano:
        requisicoes = requisicoes.order_by('-data_solicitacao')[:50]
    else:
        requisicoes = requisicoes.order_by('-data_solicitacao')

    context = {
        'requisicoes': requisicoes,
        'mes': mes or '',
        'ano': ano or '',
        'produto': produto_query or '',
    }

    return render(request, 'core/relatorios/saida_estoque.html', context)






@login_required
def relatorio_consolidado_view(request):
    bar_id = request.session.get('bar_id')
    if not bar_id:
        return render(request, 'erro.html', {'mensagem': 'Nenhum bar selecionado.'})

    # Filtros de mês e ano, padrão para o mês atual
    hoje = timezone.now()
    mes = request.GET.get('mes', f"{hoje.month:02}")
    ano = request.GET.get('ano', str(hoje.year))

    # Buscar requisições aprovadas no período
    requisicoes = RequisicaoProduto.objects.filter(
        bar_id=bar_id,
        status='APROVADA',
        data_solicitacao__month=int(mes),
        data_solicitacao__year=int(ano)
    )

    dados_agrupados = defaultdict(lambda: defaultdict(dict))  # data -> produto_obj -> dados

    for req in requisicoes:
        data_str = req.data_solicitacao.date().strftime('%d/%m/%Y')
        produto = req.produto

        # Acumular quantidade requisitada
        dados_agrupados[data_str][produto]['quantidade_requisitada'] = \
            dados_agrupados[data_str][produto].get('quantidade_requisitada', 0) + float(req.quantidade_solicitada)

        # Buscar contagem do dia
        contagem = ContagemBar.objects.filter(
            bar_id=bar_id,
            produto=produto,
            data_contagem__date=req.data_solicitacao.date()
        ).order_by('-data_contagem').first()

        if contagem:
            dados_agrupados[data_str][produto]['cheias'] = contagem.quantidade_garrafas_cheias
            dados_agrupados[data_str][produto]['doses'] = float(contagem.quantidade_doses_restantes)
        else:
            dados_agrupados[data_str][produto]['cheias'] = 0
            dados_agrupados[data_str][produto]['doses'] = 0.0

        # Cálculo da diferença (apenas em garrafas)
        requisitado = dados_agrupados[data_str][produto]['quantidade_requisitada']
        cheias = dados_agrupados[data_str][produto]['cheias']
        dados_agrupados[data_str][produto]['diferenca'] = abs(cheias - requisitado)


    # Converter Produto -> produto.nome e ordenar datas de forma decrescente
    dados_agrupados_final = dict()
    for data in sorted(dados_agrupados.keys(), reverse=True):
        produtos_formatados = {
            produto_obj.nome: dados
            for produto_obj, dados in dados_agrupados[data].items()
        }
        dados_agrupados_final[data] = produtos_formatados

    context = {
        'dados_agrupados': dados_agrupados_final,
        'mes': mes,
        'ano': ano,
    }

    return render(request, 'core/relatorios/consolidado_diferenca.html', context)


@login_required
def relatorio_contagem_atual(request):
    bar_id = request.session.get('bar_id')
    if not bar_id:
        return render(request, 'erro.html', {'mensagem': 'Nenhum bar selecionado.'})

    bar_atual = Bar.objects.get(id=bar_id)
    restaurante = bar_atual.restaurante

    bares = Bar.objects.filter(restaurante=restaurante).order_by('nome')
    dados_por_bar = {}
    somatorio_total = defaultdict(lambda: {
        'garrafas': 0,
        'doses': 0.0,
        'produto': None
    })

    for bar in bares:
        contagens = ContagemBar.objects.filter(
            bar=bar
        ).order_by('-data_contagem')

        ultima_contagem_por_produto = {}
        for contagem in contagens:
            if contagem.produto_id not in ultima_contagem_por_produto:
                ultima_contagem_por_produto[contagem.produto_id] = contagem

        contagens_finais = list(ultima_contagem_por_produto.values())
        dados_por_bar[bar.nome] = contagens_finais

        for contagem in contagens_finais:
            pid = contagem.produto_id
            somatorio_total[pid]['produto'] = contagem.produto
            somatorio_total[pid]['garrafas'] += contagem.quantidade_garrafas_cheias or 0
            somatorio_total[pid]['doses'] += float(contagem.quantidade_doses_restantes or 0)

    context = {
        'dados_por_bar': dados_por_bar,
        'restaurante': restaurante,
        'somatorio_total': dict(somatorio_total),  # 🔁 converter para dict padrão
    }

    return render(request, 'core/relatorios/contagem_atual.html', context)


@login_required
def relatorio_eventos(request):
    data_inicio_str = request.GET.get('data_inicio')
    data_fim_str = request.GET.get('data_fim')
    nome_evento = request.GET.get('nome_evento', '').strip()

    eventos = Evento.objects.none()  # Começa vazio
    consolidado = {}

    filtros_aplicados = data_inicio_str or data_fim_str or nome_evento  # verifica se houve algum filtro

    if filtros_aplicados:
        eventos = Evento.objects.all()

        if nome_evento:
            eventos = eventos.filter(nome__icontains=nome_evento)

        if data_inicio_str and data_fim_str:
            try:
                data_inicio = datetime.strptime(data_inicio_str, "%Y-%m-%d").date()
                data_fim = datetime.strptime(data_fim_str, "%Y-%m-%d").date()
                eventos = eventos.filter(data_criacao__date__range=(data_inicio, data_fim))
            except (ValueError, TypeError):
                data_inicio = data_fim = None
        else:
            data_inicio = data_fim = None

        consolidado = defaultdict(lambda: {'garrafas': 0, 'doses': 0, 'ml': 0})
        for evento in eventos:
            for item in evento.produtos.all():
                consolidado[item.produto.nome]['garrafas'] += item.garrafas
                consolidado[item.produto.nome]['doses'] += item.doses
                consolidado[item.produto.nome]['ml'] += item.doses * 50

        consolidado = dict(consolidado)
    else:
        data_inicio = data_fim = None

    context = {
        'eventos': eventos,
        'consolidado': consolidado,
        'data_inicio': data_inicio_str,
        'data_fim': data_fim_str,
        'nome_evento': nome_evento,
    }
    return render(request, 'core/relatorios/relatorio_eventos.html', context)






#                                                                                     SEÇÃO DE EXPORTAÇÃO DE EXPORTAÇÃO EXCEL


@login_required
def exportar_relatorio_eventos_excel(request):
    # Recupera parâmetros GET
    data_inicio_str = request.GET.get('data_inicio')
    data_fim_str = request.GET.get('data_fim')
    nome_evento = request.GET.get('nome_evento', '').strip()

    # Tenta converter as datas, define defaults em caso de erro
    try:
        data_inicio = datetime.strptime(data_inicio_str, "%Y-%m-%d").date()
        data_fim = datetime.strptime(data_fim_str, "%Y-%m-%d").date()
    except (TypeError, ValueError):
        today = date.today()
        data_inicio = today.replace(day=1)
        data_fim = today

    # Filtra eventos por data
    eventos = Evento.objects.filter(data_criacao__date__range=(data_inicio, data_fim))

    # Aplica filtro por nome do evento se fornecido
    if nome_evento:
        eventos = eventos.filter(nome__icontains=nome_evento)

    # Consolidação dos dados
    consolidado = defaultdict(lambda: {'garrafas': 0, 'doses': 0})
    for evento in eventos:
        for item in evento.produtos.all():
            consolidado[item.produto.nome]['garrafas'] += item.garrafas
            consolidado[item.produto.nome]['doses'] += item.doses

    # Criação da planilha
    wb = Workbook()
    ws = wb.active
    ws.title = "Eventos"

    # Cabeçalho
    ws.append(["Produto", "Garrafas", "Doses", "Doses (ml)"])

    # Dados
    for produto, dados in consolidado.items():
        ml_total = dados['doses'] * 50  # 50ml fixos por dose
        ws.append([produto, dados['garrafas'], dados['doses'], ml_total])

    # Gera arquivo Excel em memória
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    # Resposta HTTP com download do Excel
    filename = f"relatorio_eventos_{data_inicio}_a_{data_fim}.xlsx"
    response = HttpResponse(
        output.getvalue(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = f'attachment; filename="{filename}"'
    return response





@login_required
def exportar_consolidado_eventos_excel(request):
    hoje = timezone.localdate()
    eventos = Evento.objects.filter(data_criacao__date=hoje)

    consolidado = defaultdict(lambda: {'garrafas': 0, 'doses': 0, 'ml': 0})  # <- inclui 'ml' aqui
    for evento in eventos:
        for item in evento.produtos.all():
            consolidado[item.produto.nome]['garrafas'] += item.garrafas
            consolidado[item.produto.nome]['doses'] += item.doses
            consolidado[item.produto.nome]['ml'] += item.doses * 50  # dose fixa de 50ml

    # Criar arquivo Excel
    wb = Workbook()
    ws = wb.active
    ws.title = f"Consolidado_{hoje}"

    # Cabeçalho
    ws.append(["Produto", "Garrafas", "Doses", "Doses (ml)"])

    # Conteúdo
    for produto, dados in consolidado.items():
        ws.append([produto, dados['garrafas'], dados['doses'], dados['ml']])

    # Resposta HTTP
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    filename = f"consolidado_eventos_{hoje}.xlsx"
    response['Content-Disposition'] = f'attachment; filename="{filename}"'
    wb.save(response)
    return response




@login_required
def exportar_saida_estoque_excel(request):
    bar_id = request.session.get('bar_id')
    if not bar_id:
        return render(request, 'erro.html', {'mensagem': 'Nenhum bar selecionado.'})

    mes = request.GET.get('mes')
    ano = request.GET.get('ano')
    produto_nome = request.GET.get('produto', '').strip().lower()

    requisicoes = RequisicaoProduto.objects.filter(
        bar_id=bar_id,
        status__in=['APROVADA', 'NEGADA', 'FALHA_ESTOQUE']
    )

    # Aplicar filtros
    if mes and ano:
        requisicoes = requisicoes.filter(data_solicitacao__month=int(mes), data_solicitacao__year=int(ano))
    elif mes:
        requisicoes = requisicoes.filter(data_solicitacao__month=int(mes))
    elif ano:
        requisicoes = requisicoes.filter(data_solicitacao__year=int(ano))

    if produto_nome:
        requisicoes = requisicoes.filter(produto__nome__icontains=produto_nome)

    # Ordenar por data (mais recentes primeiro)
    requisicoes = requisicoes.order_by('-data_solicitacao')

    # Criar planilha
    wb = Workbook()
    ws = wb.active
    ws.title = "Saída Estoque Bar"

    headers = ['Produto', 'Quantidade', 'Status', 'Data', 'Solicitado por', 'Aprovado/Negado por']
    ws.append(headers)

    total_quantidade = 0

    for req in requisicoes:
        quantidade = float(req.quantidade_solicitada or 0)
        total_quantidade += quantidade

        ws.append([
            req.produto.nome,
            f"{quantidade:.2f}",
            req.get_status_display(),
            req.data_solicitacao.strftime('%d/%m/%Y %H:%M'),
            req.usuario.username if req.usuario else '-',
            req.usuario_aprovador.username if req.usuario_aprovador else '-',
        ])

    # Linha extra de separação
    ws.append([])

    # Linha de total
    ws.append(['TOTAL', f"{total_quantidade:.2f}", '', '', '', ''])

    # Auto ajuste de colunas
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_length + 2

    # Gerar resposta HTTP
    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    response['Content-Disposition'] = 'attachment; filename=relatorio_saida_estoque.xlsx'
    wb.save(response)
    return response



@login_required
def relatorio_consolidado_excel_view(request):
    bar_id = request.session.get('bar_id')
    if not bar_id:
        return HttpResponse("Nenhum bar selecionado.", status=400)

    hoje = timezone.now()
    mes = request.GET.get('mes', f"{hoje.month:02}")
    ano = request.GET.get('ano', str(hoje.year))

    requisicoes = RequisicaoProduto.objects.filter(
        bar_id=bar_id,
        status='APROVADA',
        data_solicitacao__month=int(mes),
        data_solicitacao__year=int(ano)
    )

    dados_agrupados = defaultdict(lambda: defaultdict(dict))

    for req in requisicoes:
        data_str = req.data_solicitacao.date().strftime('%d/%m/%Y')
        produto = req.produto

        dados_agrupados[data_str][produto]['quantidade_requisitada'] = \
            dados_agrupados[data_str][produto].get('quantidade_requisitada', 0) + float(req.quantidade_solicitada)

        contagem = ContagemBar.objects.filter(
            bar_id=bar_id,
            produto=produto,
            data_contagem__date=req.data_solicitacao.date()
        ).order_by('-data_contagem').first()

        if contagem:
            dados_agrupados[data_str][produto]['cheias'] = contagem.quantidade_garrafas_cheias
            dados_agrupados[data_str][produto]['doses'] = float(contagem.quantidade_doses_restantes)
        else:
            dados_agrupados[data_str][produto]['cheias'] = 0
            dados_agrupados[data_str][produto]['doses'] = 0.0

        requisitado = dados_agrupados[data_str][produto]['quantidade_requisitada']
        cheias = dados_agrupados[data_str][produto]['cheias']
        dados_agrupados[data_str][produto]['diferenca'] = abs(cheias - requisitado)

    # Criar o Excel
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Consolidado Diferença"

    ws.append(["Data", "Produto", "Requisitado", "Contagem (Garrafas)", "Contagem (Doses)", "Diferença"])

    bold_font = Font(bold=True)
    for cell in ws["1:1"]:
        cell.font = bold_font

    for data in sorted(dados_agrupados.keys(), reverse=True):
        for produto, dados in dados_agrupados[data].items():
            ws.append([
                data,
                produto.nome,
                f"{dados['quantidade_requisitada']:.2f}",
                dados['cheias'],
                f"{dados['doses']:.2f}",
                f"{dados['diferenca']:.2f}"
            ])

    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    filename = f"relatorio_consolidado_{mes}_{ano}.xlsx"
    response['Content-Disposition'] = f'attachment; filename={filename}'
    wb.save(response)
    return response




@login_required
def exportar_contagem_atual_excel(request):
    bar_id = request.session.get('bar_id')
    if not bar_id:
        return HttpResponse("Nenhum bar selecionado.", status=400)

    bar_atual = Bar.objects.get(id=bar_id)
    restaurante = bar_atual.restaurante
    bares = Bar.objects.filter(restaurante=restaurante).order_by('nome')

    dados_por_bar = {}
    somatorio_total = defaultdict(lambda: {'garrafas': 0, 'doses': 0.0, 'produto': None})

    for bar in bares:
        contagens = ContagemBar.objects.filter(bar=bar).order_by('-data_contagem')
        ultima_contagem_por_produto = {}

        for contagem in contagens:
            if contagem.produto_id not in ultima_contagem_por_produto:
                ultima_contagem_por_produto[contagem.produto_id] = contagem

        contagens_finais = list(ultima_contagem_por_produto.values())
        dados_por_bar[bar.nome] = contagens_finais

        for contagem in contagens_finais:
            pid = contagem.produto_id
            somatorio_total[pid]['produto'] = contagem.produto
            somatorio_total[pid]['garrafas'] += contagem.quantidade_garrafas_cheias or 0
            somatorio_total[pid]['doses'] += float(contagem.quantidade_doses_restantes or 0)

    # Criar arquivo Excel
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet('Contagem Atual')

    # Estilos
    bold = workbook.add_format({'bold': True})
    date_format = workbook.add_format({'num_format': 'dd/mm/yyyy hh:mm'})
    number_format = workbook.add_format({'num_format': '#,##0.00'})

    row = 0

    # ✅ Título da seção de totais
    worksheet.write(row, 0, "Total por Produto no Restaurante", bold)
    row += 1

    # ✅ Cabeçalhos
    headers = ["Produto", "Total de Garrafas", "Total de Doses", "Total de Doses (ML)"]
    for col, header in enumerate(headers):
        worksheet.write(row, col, header, bold)
    row += 1

    # ✅ Dados
    for item in somatorio_total.values():
        worksheet.write(row, 0, item['produto'].nome)
        worksheet.write(row, 1, item['garrafas'])
        worksheet.write(row, 2, item['doses'], number_format)
        worksheet.write(row, 3, item['doses'] * 50, number_format)
        row += 1

    row += 2  # espaço antes da próxima seção

    # ✅ Dados por bar
    for bar_nome, contagens in dados_por_bar.items():
        worksheet.write(row, 0, f"Bar: {bar_nome}", bold)
        row += 1

        headers = ["Produto", "Garrafas", "Doses", "Doses (ML)", "Data da Contagem", "Usuário"]
        for col, header in enumerate(headers):
            worksheet.write(row, col, header, bold)
        row += 1

        for contagem in contagens:
            doses = float(contagem.quantidade_doses_restantes or 0)
            worksheet.write(row, 0, contagem.produto.nome)
            worksheet.write(row, 1, contagem.quantidade_garrafas_cheias)
            worksheet.write(row, 2, doses, number_format)
            worksheet.write(row, 3, doses * 50, number_format)

            data_contagem = contagem.data_contagem
            if is_aware(data_contagem):
                data_contagem = data_contagem.replace(tzinfo=None)

            worksheet.write(row, 4, data_contagem, date_format)
            worksheet.write(row, 5, contagem.usuario.username)
            row += 1

        row += 2  # espaço entre bares

    # Finalizar e enviar
    workbook.close()
    output.seek(0)

    filename = f"relatorio_contagem_atual_{restaurante.nome.replace(' ', '_')}.xlsx"
    response = HttpResponse(output.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = f'attachment; filename={filename}'

    return response





#                                                                                  GRAFICOS


