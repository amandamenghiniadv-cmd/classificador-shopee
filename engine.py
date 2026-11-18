"""Classificador de produtos usando pesquisa web e critérios do Cauê Oliveira."""
import os, re, json, time
from datetime import datetime, timezone
from urllib.parse import quote
import requests

RULES = '''Classifique o TIPO GENÉRICO do produto (ignore marca) pela categoria PREDOMINANTE, segundo SOMENTE estas definições do Cauê Oliveira:
NECESSIDADE: venda o ano todo, preço competitivo, margem mais apertada, giro e volume, validade comercial longa; cliente já sabe o que quer, decisão rápida por preço/prazo, demanda morna-quente constante.
DESEJO: demanda híbrida, validade comercial menor, não necessariamente menor preço, margem potencialmente melhor; cliente pode não procurar, vê e deseja, emoção/identidade, foto, transformação, prova social e percepção de valor.
SAZONAL: procura concentrada numa janela do ano, não necessariamente estação; timing decisivo, alto volume/margem na janela, validade comercial curta, risco de sobrar estoque.
IMPORTANTE: não infira margem real sem dados; categorias são arquétipos, não provas de desempenho. Uma busca constante não implica necessidade; produtos de desejo também podem vender o ano todo. Não force confiança alta. Fontes devem apoiar fatos específicos; não invente série histórica, números ou fontes. Se a evidência for fraca, assinale confiança baixa e explique.
Retorne APENAS JSON com chaves categoria (Necessidade/Desejo/Sazonal), produto_generico, justificativa, confianca (Alta/Média/Baixa), evidencias (lista de textos breves com URLs já fornecidas), lacunas (texto).'''

def search_web(term, key, count=5):
    if not key: return []
    queries=[f'"{term}" uso rotina maquiagem beleza Brasil', f'"{term}" tendência carnaval sazonalidade Brasil']
    found=[]
    for q in queries:
        r=requests.get('https://serpapi.com/search.json',params={'engine':'google','q':q,'gl':'br','hl':'pt','num':count,'api_key':key},timeout=35)
        r.raise_for_status()
        obj=r.json()
        if obj.get('error'): raise RuntimeError(obj['error'])
        for x in obj.get('organic_results',[]):
            if x.get('link') and x.get('snippet'):
                found.append({'titulo':x.get('title',''),'resumo':x['snippet'],'url':x['link']})
    return list({x['url']:x for x in found}.values())[:10]

def classify(term, openai_key=None, search_key=None, model='gpt-4.1-mini'):
    term=term.strip()
    if not term: raise ValueError('Informe um produto.')
    openai_key=openai_key or os.getenv('OPENAI_API_KEY')
    search_key=search_key or os.getenv('SERPAPI_KEY')
    if not openai_key or not search_key:
        raise RuntimeError('Para pesquisa automática, configure OPENAI_API_KEY e SERPAPI_KEY. Nenhuma classificação foi inventada.')
    sources=search_web(term,search_key)
    if not sources:
        raise RuntimeError('Nenhuma evidência de busca disponível; não foi gerada classificação.')
    payload={'model':model,'temperature':0,'response_format':{'type':'json_object'},'messages':[
        {'role':'system','content':RULES},
        {'role':'user','content':json.dumps({'produto':term,'fontes_encontradas':sources},ensure_ascii=False)}]}
    r=requests.post('https://api.openai.com/v1/chat/completions',headers={'Authorization':'Bearer '+openai_key,'Content-Type':'application/json'},json=payload,timeout=90)
    r.raise_for_status()
    result=json.loads(r.json()['choices'][0]['message']['content'])
    if result.get('categoria') not in ('Necessidade','Desejo','Sazonal'): raise ValueError('Categoria inesperada na resposta da IA.')
    allowed={s['url'] for s in sources}
    # URLs mostradas no resultado são as efetivamente recuperadas, nunca URLs inventadas pelo modelo.
    result['fontes_consultadas']=sources
    result['produto_informado']=term
    result['google_trends_link']='https://trends.google.com/trends/explore?date=today%2012-m&geo=BR&q='+quote(result.get('produto_generico') or term)
    result['data_pesquisa_utc']=datetime.now(timezone.utc).isoformat()
    result['nota']='Classificação qualitativa por IA, não validação de volume de buscas ou vendas.'
    return result


def excel_bytes(results):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    from io import BytesIO

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

    for categoria in ("Necessidade", "Desejo", "Sazonal"):
        ws = wb.create_sheet(categoria)
        ws.append(headers)

        for resultado in results:
            if resultado.get("categoria") != categoria:
                continue

            fontes = "\n".join(
                fonte.get("url", "")
                for fonte in resultado.get("fontes_consultadas", [])
            )

            ws.append([
                resultado.get("produto_informado", ""),
                resultado.get("produto_generico", ""),
                categoria,
                resultado.get("confianca", ""),
                resultado.get("justificativa", ""),
                resultado.get("lacunas", ""),
                fontes,
                resultado.get("google_trends_link", ""),
                resultado.get("data_pesquisa_utc", ""),
                resultado.get("nota", "")
            ])

        for cell in ws[1]:
            cell.fill = PatternFill(
                start_color="16324A",
                end_color="16324A",
                fill_type="solid"
            )
            cell.font = Font(color="FFFFFF", bold=True)
            cell.alignment = Alignment(wrap_text=True)

        widths = [31, 30, 18, 17, 75, 65, 80, 80, 29, 65]

        for index, width in enumerate(widths, start=1):
            ws.column_dimensions[
                get_column_letter(index)
            ].width = width

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(
                    wrap_text=True,
                    vertical="top"
                )

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    arquivo = BytesIO()
    wb.save(arquivo)
    return arquivo.getvalue()

