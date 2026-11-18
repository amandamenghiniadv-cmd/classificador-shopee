
"""Classificador Shopee: Necessidade, Desejo e Sazonal."""

import os
import json
from datetime import datetime, timezone
from urllib.parse import quote
from io import BytesIO

import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# REGRAS DE CLASSIFICACAO DO CAUE OLIVEIRA

RULES = """
Classifique o TIPO GENERICO do produto, ignorando a marca,
pela categoria PREDOMINANTE.

Utilize SOMENTE os conceitos abaixo:

NECESSIDADE:
- Vende o ano todo.
- Preco relativamente competitivo.
- Margem mais apertada.
- Giro e volume.
- Validade comercial longa.
- Consumidor ja sabe o que quer.
- Decisao rapida baseada em preco e prazo.
- Demanda morna a quente e constante.

DESEJO:
- Demanda hibrida.
- Pode vender o ano todo ou por determinado periodo.
- Validade comercial menor que necessidade.
- Nao compete necessariamente pelo menor preco.
- Margem potencialmente melhor.
- Consumidor pode nao estar procurando o produto.
- Ao visualizar, passa a desejar.
- Emocao, identidade, transformacao e prova social.
- Percepcao de valor e apresentacao visual.

SAZONAL:
- Procura concentrada em uma janela especifica do ano.
- Nao precisa estar relacionada a uma estacao.
- Timing e decisivo.
- Pode gerar alto volume e margem durante a janela.
- Validade comercial curta.
- Risco de estoque parado apos a janela.

REGRAS IMPORTANTES:
- Nao invente numeros de vendas ou margens.
- Nao confunda demanda constante com necessidade.
- Produtos de desejo tambem podem vender o ano todo.
- Considere o comportamento predominante do consumidor.
- Utilize as fontes encontradas na pesquisa.
- Nao invente fontes ou dados historicos.
- Se as evidencias forem fracas, indique confianca baixa.
- A classificacao e qualitativa, nao uma prova de vendas.

Retorne SOMENTE um objeto JSON com as chaves:
categoria: Necessidade, Desejo ou Sazonal
produto_generico: nome generico do produto
justificativa: explicacao objetiva da classificacao
confianca: Alta, Média ou Baixa
evidencias: lista de evidencias encontradas
lacunas: informacoes que faltam para confirmar a analise
"""


# PESQUISA DE PRODUTOS NA INTERNET

def search_web(term, key, count=5):

    if not key:
        return []

    queries = [
        f'"{term}" uso rotina maquiagem beleza Brasil',
        f'"{term}" tendência carnaval sazonalidade Brasil'
    ]

    found = []

    for query in queries:

        response = requests.get(
            "https://serpapi.com/search.json",
            params={
                "engine": "google",
                "q": query,
                "gl": "br",
                "hl": "pt",
                "num": count,
                "api_key": key
            },
            timeout=35
        )

        response.raise_for_status()

        data = response.json()

        if data.get("error"):
            raise RuntimeError(data["error"])

        for item in data.get("organic_results", []):

            if item.get("link") and item.get("snippet"):

                found.append({
                    "titulo": item.get("title", ""),
                    "resumo": item["snippet"],
                    "url": item["link"]
                })

    unique = {
        item["url"]: item
        for item in found
    }

    return list(unique.values())[:10]


# CLASSIFICACAO COM GOOGLE GEMINI

def classify(
    term,
    openai_key=None,
    search_key=None,
    model="gemini-2.5-flash"
):

    term = term.strip()

    if not term:
        raise ValueError("Informe um produto.")

    gemini_key = os.getenv("GEMINI_API_KEY")
    search_key = search_key or os.getenv("SERPAPI_KEY")

    if not gemini_key:
        raise RuntimeError(
            "Configure GEMINI_API_KEY nos Secrets do Streamlit."
        )

    if not search_key:
        raise RuntimeError(
            "Configure SERPAPI_KEY para pesquisar na internet."
        )

    # Pesquisa na internet

    sources = search_web(term, search_key)

    if not sources:
        raise RuntimeError(
            "Nenhuma fonte encontrada. "
            "Nao foi gerada classificacao."
        )

    # Envia as fontes e regras para o Gemini

    prompt = {
        "produto": term,
        "fontes_encontradas": sources,
        "instrucoes": RULES
    }

    payload = {
        "contents": [
            {
                "parts": [
                    {
                        "text": json.dumps(
                            prompt,
                            ensure_ascii=False
                        )
                    }
                ]
            }
        ],
        "generationConfig": {
            "temperature": 0,
            "responseMimeType": "application/json"
        }
    }

    url = (
        "https://generativelanguage.googleapis.com/"
        "v1beta/models/"
        + model
        + ":generateContent"
    )

    response = requests.post(
        url,
        headers={
            "x-goog-api-key": gemini_key,
            "Content-Type": "application/json"
        },
        json=payload,
        timeout=90
    )

    if not response.ok:
        raise RuntimeError(
            "Erro na API Gemini: "
            + str(response.status_code)
            + " - "
            + response.text[:500]
        )

    data = response.json()

    try:
        content = data["candidates"][0]["content"]["parts"][0]["text"]
        result = json.loads(content)
    except (KeyError, IndexError, TypeError, ValueError) as exc:
        raise RuntimeError(
            "O Gemini nao retornou uma classificacao valida."
        ) from exc

    if result.get("categoria") not in (
        "Necessidade",
        "Desejo",
        "Sazonal"
    ):
        raise ValueError(
            "Categoria inesperada na resposta da IA."
        )

    # Complementa os resultados

    result["fontes_consultadas"] = sources
    result["produto_informado"] = term

    result["google_trends_link"] = (
        "https://trends.google.com/trends/explore"
        "?date=today%2012-m&geo=BR&q="
        + quote(result.get("produto_generico") or term)
    )

    result["data_pesquisa_utc"] = (
        datetime.now(timezone.utc).isoformat()
    )

    result["nota"] = (
        "Classificacao qualitativa por IA. "
        "Nao representa validacao de vendas "
        "ou volume de buscas."
    )

    return result


# GERACAO DO EXCEL

def excel_bytes(results):

    wb = Workbook()
    wb.remove(wb.active)

    headers = [
        "Produto informado",
        "Produto genérico",
        "Categoria",
        "Confiança",
        "Justificativa",
        "Lacunas",
        "Fontes consultadas",
        "Google Trends BR 12m",
        "Data UTC",
        "Nota"
    ]

    categories = (
        "Necessidade",
        "Desejo",
        "Sazonal"
    )

    for category in categories:

        ws = wb.create_sheet(category)
        ws.append(headers)

        for result in results:

            if result.get("categoria") != category:
                continue

            sources = "\n".join(
                source.get("url", "")
                for source in result.get(
                    "fontes_consultadas", []
                )
            )

            ws.append([
                result.get("produto_informado", ""),
                result.get("produto_generico", ""),
                category,
                result.get("confianca", ""),
                result.get("justificativa", ""),
                result.get("lacunas", ""),
                sources,
                result.get("google_trends_link", ""),
                result.get("data_pesquisa_utc", ""),
                result.get("nota", "")
            ])

        # Cabecalho

        for cell in ws[1]:

            cell.fill = PatternFill(
                start_color="16324A",
                end_color="16324A",
                fill_type="solid"
            )

            cell.font = Font(
                color="FFFFFF",
                bold=True
            )

            cell.alignment = Alignment(
                wrap_text=True
            )

        # Largura das colunas

        widths = [
            31, 30, 18, 17, 75,
            65, 80, 80, 29, 65
        ]

        for index, width in enumerate(
            widths,
            start=1
        ):

            ws.column_dimensions[
                get_column_letter(index)
            ].width = width

        # Formatacao das linhas

        for row in ws.iter_rows(min_row=2):

            for cell in row:

                cell.alignment = Alignment(
                    wrap_text=True,
                    vertical="top"
                )

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    # Cria arquivo Excel em memoria

    arquivo = BytesIO()
    wb.save(arquivo)

    return arquivo.getvalue()
