# import streamlit as st
# import pandas as pd
# from datetime import datetime
# import yagmail

# # Inicializa o estado da sessão para armazenar os dados do fechamento
# if 'fechamento' not in st.session_state:
#     st.session_state.fechamento = {}

# st.title("💳 Fechamento de Caixa - Villa Sonali")

# # Entradas do colaborador
# valor_pix = st.number_input("💵 Valor recebido em PIX:", min_value=0.0, step=0.01)
# valor_dinheiro = st.number_input("💵 Valor recebido em Dinheiro:", min_value=0.0, step=0.01)
# valor_cartao = st.number_input("💳 Valor recebido em Cartão:", min_value=0.0, step=0.01)
# valor_pendura = st.number_input("📞 Valor em Pendura:", min_value=0.0, step=0.01)

# valor_total_vendas = st.number_input("📈 Valor Total de Vendas:", min_value=0.0, step=0.01)
# numero_clientes = st.number_input("👥 Número de Clientes:", min_value=0, step=1)

# # Botão para gerar planilha e enviar
# if st.button("📅 Gerar Planilha e Enviar por Email"):
#     fechamento_data = {
#         "Data": [datetime.now().strftime("%Y-%m-%d")],
#         "Valor PIX (R$)": [valor_pix],
#         "Valor Dinheiro (R$)": [valor_dinheiro],
#         "Valor Cartão (R$)": [valor_cartao],
#         "Valor Pendura (R$)": [valor_pendura],
#         "Total de Entradas (R$)": [valor_pix + valor_dinheiro + valor_cartao + valor_pendura],
#         "Valor Total de Vendas (R$)": [valor_total_vendas],
#         "Número de Clientes": [numero_clientes],
#         "Divergência?": ["Sim" if (valor_pix + valor_dinheiro + valor_cartao + valor_pendura) != valor_total_vendas else "Não"]
#     }

#     df_fechamento = pd.DataFrame(fechamento_data)

#     nome_arquivo = f"fechamento_caixa_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
#     df_fechamento.to_excel(nome_arquivo, index=False)

#     try:
#         yag = yagmail.SMTP(user="ale.moreira@gmail.com", password="gncuqrzzkstgeamn")
#         yag.send(
#             to="ale.moreira@gmail.com",
#             subject="📋 Fechamento de Caixa - Villa Sonali",
#             contents="Segue em anexo o fechamento de caixa realizado hoje.",
#             attachments=nome_arquivo
#         )
#         st.success("Email enviado com sucesso!")
#         st.session_state.fechamento = {}
#     except Exception as e:
#         st.error(f"Erro ao enviar email: {e}")





import streamlit as st
import pandas as pd
from datetime import datetime

st.title("📦 Fechamento de Caixa - Villa Sonali")

# Entradas do usuário
valor_pix = st.number_input("💳 Valor em PIX (R$):", min_value=0.0, step=0.01)
valor_dinheiro = st.number_input("💵 Valor em Dinheiro (R$):", min_value=0.0, step=0.01)
valor_cartao = st.number_input("💳 Valor em Cartão (R$):", min_value=0.0, step=0.01)
valor_pendura = st.number_input("🧾 Valor Pendura (R$):", min_value=0.0, step=0.01)
valor_total_vendas = st.number_input("💰 Valor Total de Vendas (com 10%) (R$):", min_value=0.0, step=0.01)
numero_clientes = st.number_input("👥 Número de Clientes:", min_value=1, step=1)

# Botão para gerar a planilha
if st.button("📥 Gerar Planilha"):
    # Calcula o total somado das entradas
    total_entradas = round(valor_pix + valor_dinheiro + valor_cartao + valor_pendura, 2)

    # Verifica divergência
    divergente = "✅ Sem divergência" if total_entradas == round(valor_total_vendas, 2) else "❌ Divergência detectada"

    # Valor bruto (sem os 10%)
    valor_bruto = round(valor_total_vendas / 1.10, 2)

    # Ticket médio
    ticket_medio = round(valor_total_vendas / numero_clientes, 2)

    # Cria DataFrame com os dados
    df = pd.DataFrame({
        "Tipo": ["Valor PIX", "Valor Dinheiro", "Valor Cartão", "Valor Pendura", "", "Valor Total de Vendas (com 10%)", "Valor de Venda Bruto (sem 10%)", "Número de Clientes", "Ticket Médio", "Verificação"],
        "Valor (R$)": [valor_pix, valor_dinheiro, valor_cartao, valor_pendura, "", valor_total_vendas, valor_bruto, numero_clientes, ticket_medio, divergente]
    })

    # Nome do arquivo
    data_hora = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    nome_arquivo = f"fechamento_caixa_{data_hora}.xlsx"

    # Salva a planilha
    df.to_excel(nome_arquivo, index=False)

    st.success(f"✅ Planilha gerada com sucesso: `{nome_arquivo}`")
    st.download_button("📂 Baixar Planilha", data=open(nome_arquivo, "rb"), file_name=nome_arquivo, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
