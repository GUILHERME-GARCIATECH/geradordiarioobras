from __future__ import annotations

from docxtpl import DocxTemplate, RichText


def gerar_docx_template(caminho_template, caminho_saida, diarios: list[dict]):
    doc = DocxTemplate(caminho_template)

    diarios_contexto = []
    for i, diario in enumerate(diarios):
        tarefas = diario.get("tarefas", [])
        mao = diario.get("mao_de_obra", {})

        quebra = RichText()
        quebra.add("\f")

        diarios_contexto.append({
            "medicao": diario.get("medicao", ""),
            "comprador": diario.get("comprador", ""),
            "contratada": diario.get("contratada", ""),
            "periodo": diario.get("periodo", ""),
            "contrato": diario.get("contrato", ""),
            "objeto": diario.get("objeto_obra", ""),
            "data": diario.get("data", ""),
            "data_assinatura": diario.get("data_assinatura", ""),

            "mestre": mao.get("mestre_de_Obras", 0),
            "eletricista": mao.get("eletricista", 0),
            "pedreiro": mao.get("pedreiro", 0),
            "servente": mao.get("servente", 0),
            "encanador": mao.get("encanador", 0),
            "pintor": mao.get("pintor", 0),

            "climamanha": diario.get("tempo_manha", ""),
            "climatarde": diario.get("tempo_tarde", ""),
            "intmanha": diario.get("interrupcao_manha", ""),
            "inttarde": diario.get("interrupcao_tarde", ""),

            "etapa": diario.get("etapa_obra", ""),
            "tarefa1": tarefas[0] if len(tarefas) > 0 else "",
            "tarefa2": tarefas[1] if len(tarefas) > 1 else "",
            "tarefa3": tarefas[2] if len(tarefas) > 2 else "",
            "tarefa4": tarefas[3] if len(tarefas) > 3 else "",
            "tarefa5": tarefas[4] if len(tarefas) > 4 else "",
            "tarefa6": tarefas[5] if len(tarefas) > 5 else "",
            "tarefa7": tarefas[6] if len(tarefas) > 6 else "",
            "tarefa8": tarefas[7] if len(tarefas) > 7 else "",
            "tarefa9": tarefas[8] if len(tarefas) > 8 else "",

            "ocorrencias": diario.get("ocorrencias", ""),
            "fiscal": diario.get("fiscal", ""),
            "crea": diario.get("crea_fiscal", ""),
            "quebra": quebra,
        })

    contexto = {"diarios": diarios_contexto}
    doc.render(contexto)
    doc.save(caminho_saida)