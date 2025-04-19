import streamlit as st
import pandas as pd
from datetime import datetime
import yagmail

# Inicializa o estado da sessão para armazenar os dados do fechamento
if 'fechamento' not in st.session_state:
    st.session_state.fechamento = {}

st.title("💳 Fechamento de Caixa - Villa Sonali")

# Entradas do colaborador
valor_pix = st.number_input("💵 Valor recebido em PIX:", min_value=0.0, step=0.01)
valor_dinheiro = st.number_input("💵 Valor recebido em Dinheiro:", min_value=0.0, step=0.01)
valor_cartao = st.number_input("💳 Valor recebido em Cartão:", min_value=0.0, step=0.01)
valor_pendura = st.number_input("📞 Valor em Pendura:", min_value=0.0, step=0.01)

valor_total_vendas = st.number_input("📈 Valor Total de Vendas:", min_value=0.0, step=0.01)
numero_clientes = st.number_input("👥 Número de Clientes:", min_value=0, step=1)

# Botão para gerar planilha e enviar
if st.button("📅 Gerar Planilha e Enviar por Email"):
    fechamento_data = {
        "Data": [datetime.now().strftime("%Y-%m-%d")],
        "Valor PIX (R$)": [valor_pix],
        "Valor Dinheiro (R$)": [valor_dinheiro],
        "Valor Cartão (R$)": [valor_cartao],
        "Valor Pendura (R$)": [valor_pendura],
        "Total de Entradas (R$)": [valor_pix + valor_dinheiro + valor_cartao + valor_pendura],
        "Valor Total de Vendas (R$)": [valor_total_vendas],
        "Número de Clientes": [numero_clientes],
        "Divergência?": ["Sim" if (valor_pix + valor_dinheiro + valor_cartao + valor_pendura) != valor_total_vendas else "Não"]
    }

    df_fechamento = pd.DataFrame(fechamento_data)

    nome_arquivo = f"fechamento_caixa_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
    df_fechamento.to_excel(nome_arquivo, index=False)

    try:
        yag = yagmail.SMTP(user="ale.moreira@gmail.com", password="gncuqrzzkstgeamn")
        yag.send(
            to="ale.moreira@gmail.com",
            subject="📋 Fechamento de Caixa - Villa Sonali",
            contents="Segue em anexo o fechamento de caixa realizado hoje.",
            attachments=nome_arquivo
        )
        st.success("Email enviado com sucesso!")
        st.session_state.fechamento = {}
    except Exception as e:
        st.error(f"Erro ao enviar email: {e}")
