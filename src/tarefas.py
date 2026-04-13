from __future__ import annotations


def normalizar_texto(texto: str) -> str:
    return " ".join(str(texto or "").strip().lower().split())


def obter_valor_campo_insensivel(registro: dict, nome_campo: str):
    nome_normalizado = normalizar_texto(nome_campo)

    for chave, valor in registro.items():
        if normalizar_texto(str(chave)) == nome_normalizado:
            return valor

    return None


def descobrir_coluna_tarefas(registro: dict, mapa_etapas: dict[str, str]) -> str | None:
    etapa = (
        obter_valor_campo_insensivel(registro, "etapa")
        or obter_valor_campo_insensivel(registro, "etapa_da_obra")
        or ""
    )
    etapa = str(etapa).strip()

    if not etapa:
        return None

    etapa_normalizada = normalizar_texto(etapa)

    for chave, coluna in mapa_etapas.items():
        if normalizar_texto(chave) == etapa_normalizada:
            return coluna
        if normalizar_texto(coluna) == etapa_normalizada:
            return coluna

    return None


def quebrar_tarefas(valor, limite: int = 9) -> list[str]:
    if not valor:
        return []

    tarefas = [t.strip() for t in str(valor).split(";") if t and t.strip()]
    return tarefas[:limite]


def extrair_tarefas(registro: dict, mapa_etapas: dict[str, str], limite: int = 9) -> list[str]:
    coluna = descobrir_coluna_tarefas(registro, mapa_etapas)

    if coluna:
        valor = obter_valor_campo_insensivel(registro, coluna)
        tarefas = quebrar_tarefas(valor, limite)
        if tarefas:
            return tarefas

    for _, coluna_mapa in mapa_etapas.items():
        valor = obter_valor_campo_insensivel(registro, coluna_mapa)
        tarefas = quebrar_tarefas(valor, limite)
        if tarefas:
            return tarefas

    for nome in ["tarefas", "tarefas_realizadas", "atividade", "atividades"]:
        valor = obter_valor_campo_insensivel(registro, nome)
        tarefas = quebrar_tarefas(valor, limite)
        if tarefas:
            return tarefas

    return []