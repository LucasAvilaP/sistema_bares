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
from django.utils.timezone import is_aware, localtime
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.formatting.rule import ColorScaleRule
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
from django.utils.text import slugify

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
@login_required
def relatorio_eventos(request):
    # pega filtros (ou assume mês atual)
    data_inicio_param = request.GET.get('data_inicio')
    data_fim_param = request.GET.get('data_fim')
    nome_evento = (request.GET.get('nome_evento') or "").strip()

    if data_inicio_param and data_fim_param:
        try:
            data_inicio = datetime.strptime(data_inicio_param, "%Y-%m-%d").date()
            data_fim = datetime.strptime(data_fim_param, "%Y-%m-%d").date()
        except (ValueError, TypeError):
            data_inicio = date.today().replace(day=1)
            data_fim = date.today()
    else:
        data_inicio = date.today().replace(day=1)
        data_fim = date.today()

    eventos_qs = (
        Evento.objects
        .filter(data_criacao__date__range=(data_inicio, data_fim))
        .order_by('-data_criacao')
    )
    if nome_evento:
        eventos_qs = eventos_qs.filter(nome__icontains=nome_evento)

    # consolida por produto
    consolidado = defaultdict(lambda: {'garrafas': 0, 'doses': Decimal('0'), 'ml': Decimal('0')})

    # estrutura por evento (com totais)
    eventos = []  # cada item: {'obj': ev, 'data': dt, 'responsavel': ..., 'itens': [...], 'totais': {...}}

    for ev in eventos_qs:
        total_g, total_d, total_ml = 0, Decimal('0'), Decimal('0')
        itens = []
        for item in ev.produtos.all():
            g = int(item.garrafas or 0)
            d = Decimal(item.doses or 0)
            ml = d * DOSE_ML

            itens.append({
                'produto': item.produto.nome,
                'garrafas': g,
                'doses': d,
                'ml': ml,
            })

            # atualiza totais do evento
            total_g += g
            total_d += d
            total_ml += ml

            # atualiza consolidado
            nome_prod = item.produto.nome
            consolidado[nome_prod]['garrafas'] += g
            consolidado[nome_prod]['doses'] += d
            consolidado[nome_prod]['ml'] += ml

        eventos.append({
            'obj': ev,
            'data': localtime(ev.data_criacao),
            'responsavel': getattr(ev, 'responsavel', ''),  # ajuste se o campo tiver outro nome
            'itens': itens,
            'totais': {'garrafas': total_g, 'doses': total_d, 'ml': total_ml},
        })

    consolidado = OrderedDict(sorted(consolidado.items(), key=lambda kv: kv[0].lower()))

    context = {
        'eventos': eventos,                 # lista preparada acima
        'consolidado': consolidado,
        'data_inicio': data_inicio,         # datas como date (p/ usar |date no template)
        'data_fim': data_fim,
        'nome_evento': nome_evento,
    }
    return render(request, 'core/relatorios/relatorio_eventos.html', context)



@login_required
def relatorio_diferenca_contagens(request):
    bar_id = request.session.get('bar_id')
    if not bar_id:
        return render(request, 'erro.html', {'mensagem': 'Nenhum bar selecionado.'})

    bar_atual = get_object_or_404(Bar, id=bar_id)
    restaurante = bar_atual.restaurante

    # Todos os bares do restaurante (igual aos outros relatórios)
    bares = Bar.objects.filter(restaurante=restaurante).order_by('nome')

    # Estruturas de saída
    dados_por_bar = {}  # { "Nome do Bar": [ {produto, ultimo, penultimo, diffs...}, ...] }
    somatorio_total = defaultdict(lambda: {
        'produto': None,
        'diff_garrafas': Decimal('0'),
        'diff_doses': Decimal('0'),
    })

    for bar in bares:
        # Pega contagens do bar, mais novas primeiro
        contagens = (
            ContagemBar.objects
            .filter(bar=bar)
            .select_related('produto', 'usuario')
            .order_by('-data_contagem')
        )

        # Para cada produto do bar, vamos guardar as DUAS mais recentes
        # Estrutura: { produto_id: [ultima, penultima] }
        duas_ultimas_por_produto = {}

        for c in contagens:
            pid = c.produto_id
            if pid not in duas_ultimas_por_produto:
                duas_ultimas_por_produto[pid] = [c]  # primeira (última)
            elif len(duas_ultimas_por_produto[pid]) == 1:
                duas_ultimas_por_produto[pid].append(c)  # segunda (penúltima)
            # se já tem 2, ignora

        linhas_bar = []

        for pid, lista in duas_ultimas_por_produto.items():
            ultimo = lista[0]
            penultimo = lista[1] if len(lista) > 1 else None

            # Converte doses para Decimal pra evitar bug de float
            u_g = Decimal(ultimo.quantidade_garrafas_cheias or 0)
            u_d = Decimal(ultimo.quantidade_doses_restantes or 0)

            if penultimo:
                p_g = Decimal(penultimo.quantidade_garrafas_cheias or 0)
                p_d = Decimal(penultimo.quantidade_doses_restantes or 0)

                diff_g = u_g - p_g
                diff_d = u_d - p_d
            else:
                diff_g = None
                diff_d = None

            # Alimenta o somatório consolidado (só se tiver penúltima)
            if diff_g is not None and diff_d is not None:
                somatorio_total[pid]['produto'] = ultimo.produto
                somatorio_total[pid]['diff_garrafas'] += diff_g
                somatorio_total[pid]['diff_doses'] += diff_d

            linhas_bar.append({
                'produto': ultimo.produto,
                'ultimo': ultimo,
                'penultimo': penultimo,
                'u_g': u_g, 'u_d': u_d,
                'p_g': (p_g if penultimo else None),
                'p_d': (p_d if penultimo else None),
                'diff_g': diff_g,
                'diff_d': diff_d,
            })

        # Ordena por nome do produto dentro do bar (fica mais amigável)
        linhas_bar.sort(key=lambda x: x['produto'].nome.lower())

        dados_por_bar[bar.nome] = linhas_bar

    # Ordena somatório total por nome do produto
    somatorio_total_dict = dict(sorted(
        somatorio_total.items(),
        key=lambda kv: kv[1]['produto'].nome.lower() if kv[1]['produto'] else ''
    ))

    context = {
        'restaurante': restaurante,
        'dados_por_bar': dados_por_bar,
        'somatorio_total': somatorio_total_dict,
    }
    return render(request, 'core/relatorios/diferenca_contagens.html', context)



#                                                                                     SEÇÃO DE EXPORTAÇÃO DE EXPORTAÇÃO EXCEL


def _auto_fit(ws, min_w=10, max_w=45):
    for col in ws.columns:
        length = 0
        idx = col[0].column
        for cell in col:
            s = '' if cell.value is None else str(cell.value)
            length = max(length, len(s))
        ws.column_dimensions[get_column_letter(idx)].width = max(min_w, min(max_w, length + 2))

def _style_header(row, fill_color="F1F5FF"):
    fill = PatternFill("solid", fgColor=fill_color)
    border = Border(left=Side(style="thin", color="DDDDDD"),
                    right=Side(style="thin", color="DDDDDD"),
                    top=Side(style="thin", color="DDDDDD"),
                    bottom=Side(style="thin", color="DDDDDD"))
    for c in row:
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.fill = fill
        c.border = border

def _style_body_borders(ws, r1, r2, c1, c2):
    """Borda fininha no corpo da tabela (inclusive totais)."""
    border = Border(left=Side(style="thin", color="EEEEEE"),
                    right=Side(style="thin", color="EEEEEE"),
                    top=Side(style="thin", color="EEEEEE"),
                    bottom=Side(style="thin", color="EEEEEE"))
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(row=r, column=c).border = border

@login_required
def exportar_relatorio_eventos_excel(request):
    data_inicio_str = request.GET.get('data_inicio') or ""
    data_fim_str = request.GET.get('data_fim') or ""
    nome_evento = (request.GET.get('nome_evento') or "").strip()

    # Período padrão (mês atual)
    if data_inicio_str and data_fim_str:
        try:
            data_inicio = datetime.strptime(data_inicio_str, "%Y-%m-%d").date()
            data_fim = datetime.strptime(data_fim_str, "%Y-%m-%d").date()
        except (TypeError, ValueError):
            data_inicio = date.today().replace(day=1)
            data_fim = date.today()
    else:
        data_inicio = date.today().replace(day=1)
        data_fim = date.today()

    eventos_qs = Evento.objects.filter(
        data_criacao__date__range=(data_inicio, data_fim)
    ).order_by('-data_criacao')
    if nome_evento:
        eventos_qs = eventos_qs.filter(nome__icontains=nome_evento)

    # Consolidação e detalhamento
    consolidado = defaultdict(lambda: {'garrafas': 0, 'doses': Decimal('0'), 'ml': Decimal('0')})
    detalhado = []
    eventos_group = []  # para a aba "Por Evento"

    for ev in eventos_qs:
        total_g = 0
        total_d = Decimal('0')
        total_ml = Decimal('0')
        itens_ev = []

        for item in ev.produtos.all():
            prod = item.produto.nome
            gar = int(item.garrafas or 0)
            dos = Decimal(item.doses or 0)
            ml  = dos * DOSE_ML

            # consolidado
            consolidado[prod]['garrafas'] += gar
            consolidado[prod]['doses'] += dos
            consolidado[prod]['ml'] += ml

            # detalhado geral
            detalhado.append({
                'evento': ev.nome,
                'data': localtime(ev.data_criacao),
                'produto': prod,
                'garrafas': gar,
                'doses': dos,
                'ml': ml,
            })

            # por evento
            itens_ev.append({'produto': prod, 'garrafas': gar, 'doses': dos, 'ml': ml})
            total_g += gar
            total_d += dos
            total_ml += ml

        eventos_group.append({
            'nome': ev.nome,
            'data': localtime(ev.data_criacao),
            'responsavel': getattr(ev, 'responsavel', ''),
            'itens': itens_ev,
            'totais': {'garrafas': total_g, 'doses': total_d, 'ml': total_ml},
        })

    # Workbook
    wb = Workbook()

    # === Aba 1: Consolidado ===
    ws1 = wb.active
    ws1.title = "Consolidado"

    ws1.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
    ws1.cell(row=1, column=1, value="Relatório de Eventos — Consolidado por Produto").font = Font(bold=True, size=14)

    filtro_txt = f"Período: {data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}"
    if nome_evento:
        filtro_txt += f" | Evento contém: {nome_evento}"
    ws1.merge_cells(start_row=2, start_column=1, end_row=2, end_column=4)
    ws1.cell(row=2, column=1, value=filtro_txt).font = Font(italic=True, size=11)

    ws1.append([])
    ws1.append(["Produto", "Garrafas", "Doses", "Doses (ML)"])
    _style_header(ws1[4])

    r = 5
    tot_g, tot_d, tot_ml = 0, Decimal('0'), Decimal('0')
    for prod in sorted(consolidado.keys(), key=lambda s: s.lower()):
        d = consolidado[prod]
        ws1.cell(row=r, column=1, value=prod)
        ws1.cell(row=r, column=2, value=int(d['garrafas'])).number_format = "0"
        ws1.cell(row=r, column=3, value=float(d['doses'])).number_format = "0.00"
        ws1.cell(row=r, column=4, value=float(d['ml'])).number_format = "0.00"
        tot_g += int(d['garrafas']); tot_d += d['doses']; tot_ml += d['ml']
        r += 1

    ws1.cell(row=r, column=1, value="Total").font = Font(bold=True)
    ws1.cell(row=r, column=2, value=tot_g).number_format = "0"
    ws1.cell(row=r, column=3, value=float(tot_d)).number_format = "0.00"
    ws1.cell(row=r, column=4, value=float(tot_ml)).number_format = "0.00"

    ws1.freeze_panes = "A5"
    _style_body_borders(ws1, 5, r, 1, 4)
    _auto_fit(ws1)

    # === Aba 2: Detalhado ===
    ws2 = wb.create_sheet("Detalhado")
    ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    ws2.cell(row=1, column=1, value="Relatório de Eventos — Detalhado por Item").font = Font(bold=True, size=14)

    ws2.merge_cells(start_row=2, start_column=1, end_row=2, end_column=6)
    ws2.cell(row=2, column=1, value=filtro_txt).font = Font(italic=True, size=11)

    ws2.append([])
    ws2.append(["Evento", "Data", "Produto", "Garrafas", "Doses", "Doses (ML)"])
    _style_header(ws2[4])

    r2 = 5
    for linha in detalhado:
        ws2.cell(row=r2, column=1, value=linha['evento'])
        ws2.cell(row=r2, column=2, value=linha['data'].strftime("%d/%m/%Y %H:%M"))
        ws2.cell(row=r2, column=3, value=linha['produto'])
        ws2.cell(row=r2, column=4, value=int(linha['garrafas'])).number_format = "0"
        ws2.cell(row=r2, column=5, value=float(linha['doses'])).number_format = "0.00"
        ws2.cell(row=r2, column=6, value=float(linha['ml'])).number_format = "0.00"
        r2 += 1

    ws2.freeze_panes = "A5"
    _style_body_borders(ws2, 5, r2 - 1, 1, 6)
    _auto_fit(ws2)

    # === Aba 3: Por Evento ===
    ws3 = wb.create_sheet("Por Evento")

    col_count = 6  # vamos usar 6 colunas
    current_row = 1

    # título geral
    ws3.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=col_count)
    ws3.cell(row=current_row, column=1, value="Relatório de Eventos — Por Evento").font = Font(bold=True, size=14)
    current_row += 1
    ws3.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=col_count)
    ws3.cell(row=current_row, column=1, value=filtro_txt).font = Font(italic=True, size=11)
    current_row += 2  # linha em branco

    # para cada evento, um bloco
    for ev in eventos_group:
        # Cabeçalho do bloco
        ws3.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=col_count)
        ws3.cell(row=current_row, column=1,
                 value=f"Evento: {ev['nome']}  |  Data: {ev['data'].strftime('%d/%m/%Y %H:%M')}"
                       + (f"  |  Resp.: {ev['responsavel']}" if ev['responsavel'] else "")
                 ).font = Font(bold=True, size=12)
        current_row += 1

        # Cabeçalho da tabela do evento
        ws3.append([])
        current_row += 1
        ws3.append(["Produto", "Garrafas", "Doses", "Doses (ML)", "", ""])
        _style_header(ws3[current_row])
        first_data = current_row + 1

        # Linhas do evento
        for it in ev['itens']:
            current_row += 1
            ws3.cell(row=current_row, column=1, value=it['produto'])
            ws3.cell(row=current_row, column=2, value=int(it['garrafas'])).number_format = "0"
            ws3.cell(row=current_row, column=3, value=float(it['doses'])).number_format = "0.00"
            ws3.cell(row=current_row, column=4, value=float(it['ml'])).number_format = "0.00"

        # Subtotal do evento
        current_row += 1
        ws3.cell(row=current_row, column=1, value="Subtotal do evento").font = Font(bold=True)
        ws3.cell(row=current_row, column=2, value=int(ev['totais']['garrafas'])).number_format = "0"
        ws3.cell(row=current_row, column=3, value=float(ev['totais']['doses'])).number_format = "0.00"
        ws3.cell(row=current_row, column=4, value=float(ev['totais']['ml'])).number_format = "0.00"

        # borda no bloco (tabela + subtotal)
        _style_body_borders(ws3, first_data, current_row, 1, 4)

        # Espaçamento entre eventos
        current_row += 2

    _auto_fit(ws3)

    # Output
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    fname = f"relatorio_eventos_{slugify(data_inicio)}_a_{slugify(data_fim)}.xlsx"
    resp = HttpResponse(
        output.getvalue(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    resp['Content-Disposition'] = f'attachment; filename="{fname}"'
    return resp





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



DOSE_ML = Decimal('50')

def _auto_fit_columns(ws, min_width=10, max_width=45):
    """Autoajusta a largura das colunas com limites razoáveis."""
    for col_cells in ws.columns:
        length = 0
        for cell in col_cells:
            v = cell.value
            v = '' if v is None else str(v)
            length = max(length, len(v))
        col = get_column_letter(col_cells[0].column)
        # um pouquinho de folga
        ws.column_dimensions[col].width = max(min_width, min(max_width, length + 2))

def _apply_header_style(row):
    """Aplica estilo de cabeçalho (cor, bold, centralizado)."""
    header_fill = PatternFill("solid", fgColor="E0E7FF")  # indigo-100
    border = Border(left=Side(style="thin", color="CCCCCC"),
                    right=Side(style="thin", color="CCCCCC"),
                    top=Side(style="thin", color="CCCCCC"),
                    bottom=Side(style="thin", color="CCCCCC"))
    for c in row:
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.fill = header_fill
        c.border = border

def _apply_body_borders(ws, start_row, end_row, start_col, end_col):
    """Borda fininha nas células do corpo da tabela."""
    border = Border(left=Side(style="thin", color="EEEEEE"),
                    right=Side(style="thin", color="EEEEEE"),
                    top=Side(style="thin", color="EEEEEE"),
                    bottom=Side(style="thin", color="EEEEEE"))
    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            ws.cell(row=r, column=c).border = border

@login_required
def exportar_diferenca_contagens_excel(request):
    bar_id = request.session.get('bar_id')
    if not bar_id:
        return HttpResponse("Nenhum bar selecionado.", status=400)

    bar_atual = get_object_or_404(Bar, id=bar_id)
    restaurante = bar_atual.restaurante
    bares = Bar.objects.filter(restaurante=restaurante).order_by('nome')

    # ===== Reaproveita a lógica do relatório para montar os dados =====
    dados_por_bar = {}
    somatorio_total = defaultdict(lambda: {
        'produto': None,
        'diff_garrafas': Decimal('0'),
        'diff_doses': Decimal('0'),
    })

    for bar in bares:
        contagens = (
            ContagemBar.objects
            .filter(bar=bar)
            .select_related('produto', 'usuario')
            .order_by('-data_contagem')
        )

        duas_ultimas_por_produto = {}
        for c in contagens:
            pid = c.produto_id
            if pid not in duas_ultimas_por_produto:
                duas_ultimas_por_produto[pid] = [c]
            elif len(duas_ultimas_por_produto[pid]) == 1:
                duas_ultimas_por_produto[pid].append(c)

        linhas_bar = []
        for pid, lista in duas_ultimas_por_produto.items():
            ultimo = lista[0]
            penultimo = lista[1] if len(lista) > 1 else None

            u_g = Decimal(ultimo.quantidade_garrafas_cheias or 0)
            u_d = Decimal(ultimo.quantidade_doses_restantes or 0)

            if penultimo:
                p_g = Decimal(penultimo.quantidade_garrafas_cheias or 0)
                p_d = Decimal(penultimo.quantidade_doses_restantes or 0)
                diff_g = u_g - p_g
                diff_d = u_d - p_d
            else:
                p_g = None
                p_d = None
                diff_g = None
                diff_d = None

            if diff_g is not None and diff_d is not None:
                somatorio_total[pid]['produto'] = ultimo.produto
                somatorio_total[pid]['diff_garrafas'] += diff_g
                somatorio_total[pid]['diff_doses'] += diff_d

            linhas_bar.append({
                'bar': bar.nome,
                'produto': ultimo.produto,
                'u_g': u_g, 'u_d': u_d,
                'p_g': p_g, 'p_d': p_d,
                'diff_g': diff_g, 'diff_d': diff_d,
                'data_p': (penultimo.data_contagem if penultimo else None),
                'user_p': (penultimo.usuario.username if penultimo else None),
                'data_u': ultimo.data_contagem,
                'user_u': ultimo.usuario.username,
            })

        linhas_bar.sort(key=lambda x: x['produto'].nome.lower())
        dados_por_bar[bar.nome] = linhas_bar

    somatorio_total = dict(sorted(
        somatorio_total.items(),
        key=lambda kv: kv[1]['produto'].nome.lower() if kv[1]['produto'] else ''
    ))

    # ===== Monta o Excel =====
    wb = Workbook()

    # Metadados / primeira sheet: Consolidado
    ws1 = wb.active
    ws1.title = "Consolidado"

    # Título e info
    ws1.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
    titulo = ws1.cell(row=1, column=1, value=f"Diferenças (Última − Penúltima) — {restaurante.nome}")
    titulo.font = Font(bold=True, size=14)
    titulo.alignment = Alignment(horizontal="left", vertical="center")

    now = timezone.localtime(timezone.now())
    ws1.merge_cells(start_row=2, start_column=1, end_row=2, end_column=4)
    info = ws1.cell(row=2, column=1, value=f"Gerado em {now.strftime('%d/%m/%Y %H:%M')}")
    info.font = Font(italic=True, size=11)

    # Cabeçalho
    headers1 = ["Produto", "Garrafas", "Doses", "Doses (ML)"]
    ws1.append([])
    ws1.append(headers1)
    _apply_header_style(ws1[4])

    # Linhas do consolidado
    first_data_row = 5
    r = first_data_row
    for item in somatorio_total.values():
        prod = item['produto'].nome if item['produto'] else ""
        diff_g = item['diff_garrafas']
        diff_d = item['diff_doses']
        diff_ml = (diff_d * DOSE_ML) if diff_d is not None else None

        ws1.cell(row=r, column=1, value=prod)
        c2 = ws1.cell(row=r, column=2, value=float(diff_g) if diff_g is not None else None)
        c3 = ws1.cell(row=r, column=3, value=float(diff_d) if diff_d is not None else None)
        c4 = ws1.cell(row=r, column=4, value=float(diff_ml) if diff_ml is not None else None)

        c2.number_format = "0"
        c3.number_format = "0.00"
        c4.number_format = "0.00"
        r += 1

    if r > first_data_row:
        _apply_body_borders(ws1, first_data_row, r - 1, 1, 4)
        ws1.freeze_panes = "A5"

        # Escala de cores (vermelho → branco → verde) nas diferenças
        for col in [2, 3, 4]:
            col_letter = get_column_letter(col)
            ws1.conditional_formatting.add(
                f"{col_letter}{first_data_row}:{col_letter}{r-1}",
                ColorScaleRule(start_type='num', start_value=-1, start_color='FCA5A5',  # red-300
                               mid_type='num', mid_value=0, mid_color='FFFFFF',
                               end_type='num', end_value=1, end_color='86EFAC')   # green-300
            )

    _auto_fit_columns(ws1)

    # Segunda sheet: Detalhado por bar
    ws2 = wb.create_sheet(title="Detalhado")
    ws2.append([f"Diferenças por Produto (Penúltima → Última) — {restaurante.nome}"])
    ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=12)
    ws2.cell(row=1, column=1).font = Font(bold=True, size=14)

    ws2.append([f"Gerado em {now.strftime('%d/%m/%Y %H:%M')}"])
    ws2.merge_cells(start_row=2, start_column=1, end_row=2, end_column=12)
    ws2.cell(row=2, column=1).font = Font(italic=True, size=11)

    headers2 = [
        "Bar", "Produto",
        "Penúltima (Garrafas)", "Penúltima (Doses)",
        "Última (Garrafas)", "Última (Doses)",
        "Diferença (Garrafas)", "Diferença (Doses)", "Diferença (Doses ML)",
        "Data Penúltima", "Usuário Penúltima",
        "Data Última", "Usuário Última",
    ]
    ws2.append([])
    ws2.append(headers2)
    _apply_header_style(ws2[4])

    first_data_row2 = 5
    r2 = first_data_row2
    for bar_nome, linhas in dados_por_bar.items():
        for linha in linhas:
            diff_ml = (linha['diff_d'] * DOSE_ML) if linha['diff_d'] is not None else None
            ws2.cell(row=r2, column=1, value=bar_nome)
            ws2.cell(row=r2, column=2, value=linha['produto'].nome)

            c_pg = ws2.cell(row=r2, column=3, value=float(linha['p_g']) if linha['p_g'] is not None else None)
            c_pd = ws2.cell(row=r2, column=4, value=float(linha['p_d']) if linha['p_d'] is not None else None)
            c_ug = ws2.cell(row=r2, column=5, value=float(linha['u_g']))
            c_ud = ws2.cell(row=r2, column=6, value=float(linha['u_d']))
            c_dg = ws2.cell(row=r2, column=7, value=float(linha['diff_g']) if linha['diff_g'] is not None else None)
            c_dd = ws2.cell(row=r2, column=8, value=float(linha['diff_d']) if linha['diff_d'] is not None else None)
            c_ml = ws2.cell(row=r2, column=9, value=float(diff_ml) if diff_ml is not None else None)

            for c in (c_pg, c_ug, c_dg):
                if c is not None:
                    c.number_format = "0"
            for c in (c_pd, c_ud, c_dd, c_ml):
                if c is not None:
                    c.number_format = "0.00"

            ws2.cell(row=r2, column=10, value=linha['data_p'].strftime('%d/%m/%Y %H:%M') if linha['data_p'] else None)
            ws2.cell(row=r2, column=11, value=linha['user_p'] if linha['user_p'] else None)
            ws2.cell(row=r2, column=12, value=linha['data_u'].strftime('%d/%m/%Y %H:%M'))
            ws2.cell(row=r2, column=13, value=linha['user_u'])

            r2 += 1

    if r2 > first_data_row2:
        _apply_body_borders(ws2, first_data_row2, r2 - 1, 1, 13)
        ws2.freeze_panes = "A5"

        # Escala de cores para as 3 colunas de diferença (garrafas, doses, ml)
        for col in [7, 8, 9]:
            col_letter = get_column_letter(col)
            ws2.conditional_formatting.add(
                f"{col_letter}{first_data_row2}:{col_letter}{r2-1}",
                ColorScaleRule(start_type='num', start_value=-1, start_color='FCA5A5',
                               mid_type='num', mid_value=0, mid_color='FFFFFF',
                               end_type='num', end_value=1, end_color='86EFAC')
            )

    _auto_fit_columns(ws2)

    # ===== Resposta HTTP =====
    from django.utils.text import slugify
    filename = f"dif-contagens-{slugify(restaurante.nome)}-{now.strftime('%Y%m%d-%H%M')}.xlsx"
    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = f'attachment; filename="{filename}"'
    wb.save(response)
    return response




#                                                                                  GRAFICOS


