from datetime import datetime, timedelta

def formatar_nome(nome_bruto: str) -> str:
    """Limpa o asterisco e inverte SOBRENOME, NOME para NOME SOBRENOME"""
    nome_limpo = nome_bruto.replace('*', '').strip()
    if ',' in nome_limpo:
        sobrenome, nome = nome_limpo.split(',', 1)
        return f"{nome.strip()} {sobrenome.strip()}"
    return nome_limpo

def calcular_refeicoes(data_arr: str, hora_arr: str, data_dep: str, hora_dep: str) -> int:
    """Calcula a quantidade de refeições cruzando os horários com a estadia"""
    try:
        chegada = datetime.strptime(f"{data_arr} {hora_arr}", "%d/%m/%Y %H:%M")
        saida = datetime.strptime(f"{data_dep} {hora_dep}", "%d/%m/%Y %H:%M")
    except ValueError:
        return 0

    refeicoes = 0
    dia_atual = chegada.replace(hour=0, minute=0, second=0)
    dia_final = saida.replace(hour=0, minute=0, second=0)

    while dia_atual <= dia_final:
        # Janela de Almoço: 12:00 às 15:00
        almoco_inicio = dia_atual.replace(hour=12, minute=0)
        almoco_fim = dia_atual.replace(hour=15, minute=0)
        if chegada <= almoco_fim and saida >= almoco_inicio:
            refeicoes += 1

        # Janela de Jantar: 19:00 até 01:30 do dia seguinte
        jantar_inicio = dia_atual.replace(hour=19, minute=0)
        jantar_fim = (dia_atual + timedelta(days=1)).replace(hour=1, minute=30)
        if chegada <= jantar_fim and saida >= jantar_inicio:
            refeicoes += 1

        dia_atual += timedelta(days=1)

    return refeicoes