import os,sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import streamlit as st
from engine import classify,excel_bytes

if "GEMINI_API_KEY" in st.secrets:
    os.environ["GEMINI_API_KEY"] = st.secrets["GEMINI_API_KEY"]

st.set_page_config(page_title='Classificador de Produtos',page_icon='🔎',layout='wide')
st.title('🔎 Classificador de Produtos')
st.caption('Necessidade · Desejo · Sazonal | Metodologia do Cauê Oliveira | Brasil')
st.write('Digite um ou mais produtos. O sistema pesquisa na web e explica a categoria predominante, sem considerar marcas.')
text=st.text_area('Produtos (um por linha)',placeholder='Esponja de maquiagem\nBlush em bastão\nDelineador neon',height=140)
items=list(dict.fromkeys(x.strip() for x in text.splitlines() if x.strip()))
if st.button('Pesquisar e classificar',type='primary',disabled=not items):
    st.session_state.results=[]
    progress=st.progress(0)
    for i,item in enumerate(items):
        with st.status(f'Pesquisando: {item}',expanded=False) as status:
            try:
                result=classify(item)
                st.session_state.results.append(result)
                status.update(label=f"{item}: {result['categoria']}",state='complete')
            except Exception as exc:
                status.update(label=f'{item}: pesquisa não concluída',state='error')
                st.error(str(exc))
        progress.progress((i+1)/len(items))
if st.session_state.get('results'):
    st.subheader('Resultados')
    for x in st.session_state.results:
        with st.expander(f"{x['produto_informado']} — {x['categoria']} ({x.get('confianca','')})",expanded=True):
            st.write(x.get('justificativa',''))
            st.caption('Limitações: '+x.get('lacunas',''))
            st.markdown('**Fontes consultadas**')
            for s in x.get('fontes_consultadas',[]): st.markdown(f"- [{s['titulo']}]({s['url']})")
            st.link_button('Conferir no Google Trends',x['google_trends_link'])
    st.download_button('Baixar Excel com três abas',data=excel_bytes(st.session_state.results),file_name='classificacao_produtos.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',type='primary')
st.divider()
st.caption('Exige credenciais OPENAI_API_KEY e SERPAPI_KEY configuradas no servidor. Google Trends é link para conferência, não uma série de dados extraída. APIs podem ter custos.')
