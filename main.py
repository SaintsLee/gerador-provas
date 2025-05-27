import pandas as pd
from datetime import date

import plotly.express as px

import streamlit as st
from streamlit import session_state

import auxiliares as aux

abas = [
        'Questões - Nível 01',
        'Questões - Nível 02',
        'Questões - Nível 03',
       ]

st.set_page_config('Gerador de Provas', layout="wide", page_icon='portfel_logo.ico')

with st.sidebar:
    st.title('Configuração da prova')
    banco_questoes = st.file_uploader('⚠️ Faça o upload do banco de dados no formato .xlsx',
    '.xlsx',
    label_visibility = 'hidden')

if banco_questoes is not None:
    st.title('📋 Research Portfel - Gerador de provas')
    with st.sidebar:
        lista_opcoes = aux.listar_opcoes(abas,banco_questoes)

        tab1,tab2,tab3 = st.tabs([
        '🟢 Nível 01',
        '🟡 Nível 02',
        '🔴 Nível 03',
       ])
        with tab1:
            selecionados_nv1 = aux.gera_opcoes(lista_opcoes[0],abas[0])
            df_1 = pd.DataFrame(list(selecionados_nv1.items()),
                                columns=['Tópicos','Quantidade'])
            qntd_nv1 = sum(selecionados_nv1.values())

        with tab2:
            selecionados_nv2 = aux.gera_opcoes(lista_opcoes[1],abas[1])
            df_2 = pd.DataFrame(list(selecionados_nv2.items()),
                                columns=['Tópicos','Quantidade'])
            qntd_nv2 = sum(selecionados_nv2.values())

        with tab3:
            selecionados_nv3 = aux.gera_opcoes(lista_opcoes[2],abas[2])
            df_3 = pd.DataFrame(list(selecionados_nv3.items()),
                                columns=['Tópicos','Quantidade'])
            qntd_nv3 = sum(selecionados_nv3.values())

    prova_dict = aux.gerar_dict_prova(abas,[selecionados_nv1,selecionados_nv2,selecionados_nv3])
    col1, col2, col3 = st.columns(3)
    with col1:
        fig = px.bar(df_1, x="Tópicos", y="Quantidade", title=f"Questões Nível 1 - {qntd_nv1}")
        st.plotly_chart(fig)

    with col2:
        fig = px.bar(df_2, x="Tópicos", y="Quantidade", title=f"Questões Nível 2 - {qntd_nv2}")
        st.plotly_chart(fig)

    with col3:
        fig = px.bar(df_3, x="Tópicos", y="Quantidade", title=f"Questões Nível 3 - {qntd_nv3}")
        st.plotly_chart(fig)


    if any(prova_dict[nivel] for nivel in prova_dict):

        df_prova = aux.monta_prova(banco_questoes,prova_dict)
        if 'prova_buffer' not in session_state:
            st.session_state.prova_buffer = df_prova
        if 'prova' not in session_state:
            st.session_state.prova = st.session_state.prova_buffer.drop(columns='Correta')
        if 'gabarito' not in session_state:
            st.session_state.gabarito = st.session_state.prova_buffer['Correta']

        if st.button("🔄 Atualizar a prova"):
            st.session_state.prova_buffer = df_prova
            st.session_state.prova = st.session_state.prova_buffer.drop(columns=['Correta',
                                                                                 'Dificuldade',
                                                                                 'Tópico',
                                                                                 'Nivel'])
            st.session_state.gabarito = st.session_state.prova_buffer['Correta']

        st.subheader('Escopo da prova:')
        st.dataframe(st.session_state.prova_buffer)
        if st.checkbox('Deseja fazer o download da prova em .docx?', value=False):
            st.write('Escreva o título que será definido na prova')
            titulo = st.text_input('Título da prova ⤵️')

            data_hoje = date.today().strftime("%d-%m-%y")
            if titulo:
                titulo_doc = f'{titulo} [{data_hoje}]'
                st.download_button(
                    label="Download da prova em Word",
                    data= aux.montar_prova_doc(st.session_state.prova, titulo),
                    file_name=f"Prova {titulo_doc}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    icon=":material/download:",
                )

        if st.checkbox('Deseja fazer o download do gabarito em .docx?', value=False):
            data_hoje = date.today().strftime("%d-%m-%y")
            if titulo:
                st.download_button(
                    label="Download do gabarito em Word",
                    data= aux.montar_gabarito_doc(st.session_state.gabarito, titulo),
                    file_name=f"Gabarito {titulo_doc}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    icon=":material/download:",
                )
            else:
                st.warning('Preencha o título da prova antes de gerar o gabarito!', icon="⚠️")

    else:
        st.subheader('Selecione pelo menos um tópico')
else:
    st.info('Faça o upload do banco de questões', icon="⚠️")