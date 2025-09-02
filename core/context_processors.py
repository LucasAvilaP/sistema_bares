def contexto_operacional(request):
    return {
        'ctx_restaurante_id': request.session.get('restaurante_id'),
        'ctx_bar_id': request.session.get('bar_id'),
    }
