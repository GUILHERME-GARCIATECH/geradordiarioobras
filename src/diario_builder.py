from __future__ import annotations

from datetime import date


def buscar_cadastro_obra(cadastros: list[dict], obra_texto: str) -> dict | None:
    obra_texto = str(obra_texto or "").strip().lower()

    for cadastro in cadastros:
        obra_id = str(cadastro.get("obra_id") or "").strip().lower()
        objeto = str(cadastro.get("objeto") or "").strip().lower()

        if obra_id and obra_id in obra_texto:
            return cadastro

        if objeto and objeto in obra_texto:
            return cadastro

    return None


def consolidar_mao_de_obra(registros: list[dict]) -> dict:
    campos = [
        "mestre_de_Obras",
        "eletricista",
        "pedreiro",
        "servente",
        "encanador",
        "pintor",
    ]

    if not registros:
        return {campo: 0 for campo in campos}

    registro = registros[-1]

    resultado = {}

    for campo in campos:
        valor = registro.get(campo, 0)
        try:
            resultado[campo] = int(float(valor or 0))
        except (ValueError, TypeError):
            resultado[campo] = 0

    return resultado


def juntar_textos_unicos(registros: list[dict], campo: str, separador: str = "\n") -> str:
    itens = []

    for r in registros:
        valor = str(r.get(campo) or "").strip()
        if valor and valor not in itens:
            itens.append(valor)

    return separador.join(itens)


def primeiro_valor_preenchido(registros: list[dict], campo: str) -> str:
    for r in registros:
        valor = str(r.get(campo) or "").strip()
        if valor:
            return valor
    return ""


def obter_etapa(registros_filtrados: list[dict]) -> str:
    for campo in ["etapa", "etapa_da_obra", "fase", "fase_obra"]:
        valor = juntar_textos_unicos(registros_filtrados, campo, " / ")
        if valor:
            return valor
    return ""


def formatar_data_br(data_ref: date) -> str:
    return data_ref.strftime("%d/%m/%Y")


def formatar_data_extenso(data_str: str) -> str:
    meses = [
        "janeiro", "fevereiro", "março", "abril", "maio", "junho",
        "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
    ]

    try:
        dia, mes, ano = data_str.split("/")
        dia = int(dia)
        mes = int(mes)
        ano = int(ano)

        return f"{dia} de {meses[mes - 1]} de {ano}"
    except Exception:
        return data_str


def formatar_data_com_dia_semana(data_ref: date) -> str:
    dias_semana = [
        "Segunda-feira",
        "Terça-feira",
        "Quarta-feira",
        "Quinta-feira",
        "Sexta-feira",
        "Sábado",
        "Domingo",
    ]

    dia_semana = dias_semana[data_ref.weekday()]
    data_formatada = data_ref.strftime("%d/%m/%Y")

    return f"{dia_semana}, {data_formatada}"


def montar_diario(
    registros_filtrados: list[dict],
    cadastro_obra: dict | None,
    periodo_texto: str,
    medicao: str,
    data_assinatura: str,
    tarefas_por_registro: list[str],
    data_referencia: date,
) -> dict:
    primeiro = registros_filtrados[0] if registros_filtrados else {}
    ultimo = registros_filtrados[-1] if registros_filtrados else {}

    return {
        "comprador": (cadastro_obra or {}).get("contratante", ""),
        "contratada": "KR Construtora LTDA",
        "periodo": periodo_texto,
        "contrato": (cadastro_obra or {}).get("contrato", ""),
        "objeto_obra": (cadastro_obra or {}).get("objeto", primeiro.get("obra", "")),
        "data": formatar_data_com_dia_semana(data_referencia),
        "medicao": medicao,
        "tempo_manha": primeiro_valor_preenchido(registros_filtrados, "tempo_manha"),
        "tempo_tarde": primeiro_valor_preenchido(registros_filtrados, "tempo_tarde"),
        "etapa_obra": obter_etapa(registros_filtrados),
        "ocorrencias": juntar_textos_unicos(registros_filtrados, "ocorrencias", "\n"),
        "fiscal": (cadastro_obra or {}).get("fiscal", ""),
        "crea_fiscal": (cadastro_obra or {}).get("Crea-MT", ""),
        "data_assinatura": formatar_data_extenso(data_assinatura),
        "mao_de_obra": consolidar_mao_de_obra(registros_filtrados),
        "tarefas": tarefas_por_registro,
        "interrupcao_manha": str(primeiro.get("interrupcao_manha") or "").strip(),
        "interrupcao_tarde": str(ultimo.get("interrupcao_tarde") or "").strip(),
    }