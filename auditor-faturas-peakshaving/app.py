"""
app.py
------
Aplicação Web Streamlit completa para Auditoria de Faturas de Energia (Grupo A)
e Simulação Técnico-Financeira de Operação de Grupos Geradores no Horário de Ponta (Peak Shaving).
Integração oficial com a base de tarifas homologadas da ANEEL:
https://englucas1.github.io/Aneel-Tarifas_v3/

Suporta:
- Mercado Cativo (ACR) vs. Mercado Livre (ACL) com verificação de conformidade da fatura.
- Tabela exaustiva de todas as tarifas utilizadas (com e sem tributos).
- Auditoria fiscal detalhada de tributos (ICMS, PIS, COFINS, CIP e Gross-up).
- Exportação para planilha Excel (.xlsx) com 4 abas estruturadas.
"""

from __future__ import annotations

import io
from pathlib import Path
from typing import Optional

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from export_report import gerar_relatorio_excel
from invoice_parser import InvoiceData, invoice_parser
from tariff_client import AneelTariffClient, tariff_client
from tariff_engine import ParametrosGerador, ResultadoSimulacao, tariff_engine

# Configuração da Página
st.set_page_config(
    page_title="Auditor Tarifário & Peak Shaving | Grupo A4 (ACR/ACL)",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Injeção de CSS Moderno e Responsivo
CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap');

html, body, [class*="css"] {
    font-family: 'Plus Jakarta Sans', sans-serif;
}

/* Header estilizado */
.app-hero-container {
    background: linear-gradient(135deg, #0b1329 0%, #111e38 50%, #0d2347 100%);
    border-radius: 16px;
    padding: 24px 32px;
    border: 1px solid #1e3a5f;
    box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.3);
    margin-bottom: 24px;
}

.hero-title {
    font-size: 1.85rem;
    font-weight: 800;
    color: #ffffff;
    margin-bottom: 4px;
    display: flex;
    align-items: center;
    gap: 12px;
}

.hero-subtitle {
    color: #94a3b8;
    font-size: 0.95rem;
    margin-bottom: 0;
}

/* Badge ANEEL */
.badge-aneel {
    background: linear-gradient(90deg, #0284c7 0%, #0369a1 100%);
    color: white;
    padding: 4px 10px;
    border-radius: 20px;
    font-size: 0.75rem;
    font-weight: 700;
    letter-spacing: 0.5px;
    text-transform: uppercase;
}

/* Cards de KPI Customizados */
.kpi-card {
    background: #0f172a;
    border: 1px solid #1e293b;
    border-radius: 12px;
    padding: 18px 20px;
    transition: transform 0.2s ease, box-shadow 0.2s ease;
}

.kpi-card:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 20px rgba(0, 0, 0, 0.4);
    border-color: #334155;
}

.kpi-label {
    font-size: 0.8rem;
    font-weight: 600;
    text-transform: uppercase;
    color: #64748b;
    letter-spacing: 0.5px;
    margin-bottom: 6px;
}

.kpi-value {
    font-size: 1.6rem;
    font-weight: 800;
    color: #f8fafc;
    margin-bottom: 4px;
}

.kpi-delta-positive {
    color: #10b981;
    font-size: 0.85rem;
    font-weight: 600;
}

.kpi-delta-neutral {
    color: #38bdf8;
    font-size: 0.85rem;
    font-weight: 600;
}

/* Cartão de Decisão Vencedor */
.recommendation-card {
    background: linear-gradient(135deg, rgba(6, 78, 59, 0.4) 0%, rgba(15, 23, 42, 0.9) 100%);
    border: 2px solid #059669;
    border-radius: 14px;
    padding: 24px 28px;
    margin-bottom: 24px;
    box-shadow: 0 12px 30px rgba(5, 150, 105, 0.15);
}

.recommendation-card-azul {
    background: linear-gradient(135deg, rgba(30, 64, 175, 0.4) 0%, rgba(15, 23, 42, 0.9) 100%);
    border: 2px solid #2563eb;
    border-radius: 14px;
    padding: 24px 28px;
    margin-bottom: 24px;
    box-shadow: 0 12px 30px rgba(37, 99, 235, 0.15);
}

.status-chip-live {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    background: rgba(16, 185, 129, 0.15);
    color: #34d399;
    padding: 6px 14px;
    border-radius: 20px;
    border: 1px solid rgba(16, 185, 129, 0.3);
    font-size: 0.85rem;
    font-weight: 600;
}

.status-dot {
    width: 8px;
    height: 8px;
    border-radius: 50%;
    background-color: #10b981;
    box-shadow: 0 0 8px #10b981;
}

/* Card de Auditoria */
.audit-box-ok {
    background: rgba(6, 78, 59, 0.25);
    border: 1px solid #059669;
    border-radius: 12px;
    padding: 16px 20px;
    margin-bottom: 20px;
}

.audit-box-warn {
    background: rgba(180, 83, 9, 0.25);
    border: 1px solid #d97706;
    border-radius: 12px;
    padding: 16px 20px;
    margin-bottom: 20px;
}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


def init_session_state():
    """Inicializa variáveis de estado da sessão."""
    if "fatura" not in st.session_state:
        samples = invoice_parser.get_sample_invoices()
        st.session_state.fatura = list(samples.values())[0]
    if "distribuidora_selecionada" not in st.session_state:
        st.session_state.distribuidora_selecionada = st.session_state.fatura.distribuidora_identificada
    if "ambiente_selecionado" not in st.session_state:
        st.session_state.ambiente_selecionado = st.session_state.fatura.ambiente_contratacao


init_session_state()

# =============================================================================
# HEADER / BANNER SUPERIOR
# =============================================================================
st.markdown(
    """
    <div class="app-hero-container">
        <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 16px;">
            <div>
                <div class="hero-title">
                    <span>⚡ Auditor Tarifário & Peak Shaving</span>
                    <span class="badge-aneel">Oficial ANEEL A4</span>
                </div>
                <p class="hero-subtitle">
                    Auditoria de faturas em PDF, análise de <strong>Mercado Livre (ACL) vs. Cativo (ACR)</strong>, simulação <strong>Verde vs. Azul</strong> e viabilidade de geração própria térmica no horário de ponta.
                </p>
            </div>
            <div class="status-chip-live">
                <span class="status-dot"></span>
                <span>Conectado: englucas1.github.io/Aneel-Tarifas_v3</span>
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# =============================================================================
# SIDEBAR: PARAMETRIZAÇÃO DO SKID GERADOR (OPEX) E IMPOSTOS
# =============================================================================
with st.sidebar:
    st.markdown("### ⚙️ Parâmetros Operacionais do Gerador")
    st.caption("Configurações do skid térmico para corte no horário de ponta (Peak Shaving).")

    with st.expander("⛽ Combustível & O&M (Custo Marginal)", expanded=True):
        preco_diesel = st.number_input(
            "Preço do Óleo Diesel (R$/L)",
            min_value=3.00,
            max_value=15.00,
            value=6.20,
            step=0.10,
            format="%.2f",
            help="Preço médio posto / entrega no reservatório da planta.",
        )
        consumo_esp = st.number_input(
            "Consumo Específico (L/kWh)",
            min_value=0.20,
            max_value=0.50,
            value=0.28,
            step=0.01,
            format="%.3f",
            help="Consumo típico de motores diesel ciclo 4 tempos (0.27 a 0.32 L/kWh).",
        )
        custo_om = st.number_input(
            "Custo de Manutenção / O&M (R$/kWh)",
            min_value=0.05,
            max_value=1.50,
            value=0.15,
            step=0.02,
            format="%.2f",
            help="Óleo lubrificante, filtros, líquido arrefecedor e revisões preventivas.",
        )

        custo_unit_gerador_prev = (preco_diesel * consumo_esp) + custo_om
        st.markdown(
            f"""
            <div style="background: rgba(30, 41, 59, 0.7); border-radius: 8px; padding: 10px 14px; border: 1px solid #334155; margin-top: 8px;">
                <span style="font-size: 0.8rem; color: #94a3b8;">Custo Marginal de Geração:</span><br>
                <strong style="font-size: 1.25rem; color: #38bdf8;">R$ {custo_unit_gerador_prev:.3f} / kWh</strong>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with st.expander("🔌 Capacidade & Regime de Operação", expanded=True):
        col_skid1, col_skid2 = st.columns(2)
        with col_skid1:
            pot_nominal_kva = st.number_input("Potência (kVA)", min_value=50.0, max_value=5000.0, value=500.0, step=50.0)
        with col_skid2:
            pot_continua_kw = st.number_input("Potência (kW)", min_value=40.0, max_value=4000.0, value=400.0, step=50.0)

        col_reg1, col_reg2 = st.columns(2)
        with col_reg1:
            dias_uteis = st.number_input("Dias Úteis/mês", min_value=15, max_value=25, value=22, step=1)
        with col_reg2:
            horas_ponta_dia = st.number_input("Horas Ponta/dia", min_value=1.0, max_value=5.0, value=3.0, step=0.5)

        capex_investimento = st.number_input(
            "CAPEX Skid / Instalação (R$)",
            min_value=0.0,
            max_value=10000000.0,
            value=0.0,
            step=10000.0,
            help="Preencha se desejar calcular o Payback Simples do projeto.",
        )

    with st.expander("🏛️ Tributos & Gross-up Fiscal", expanded=False):
        modo_impostos = st.radio(
            "Metodologia dos Tributos:",
            options=["Cálculo por Dentro (art. 13 LC 87/96)", "Aditivo Simples (Por Fora)", "Sem Tributos (Líquido ANEEL)"],
            index=0,
            help="No Brasil, a base de cálculo do ICMS embute o próprio ICMS e o PIS/COFINS (cálculo por dentro).",
        )
        aliquota_icms = st.number_input("Alíquota ICMS (%)", min_value=0.0, max_value=35.0, value=18.0, step=1.0)
        aliquota_pis = st.number_input("Alíquota PIS (%)", min_value=0.0, max_value=5.0, value=1.65, step=0.05)
        aliquota_cofins = st.number_input("Alíquota COFINS (%)", min_value=0.0, max_value=15.0, value=7.60, step=0.10)
        soma_aliquotas = aliquota_icms + aliquota_pis + aliquota_cofins

        if modo_impostos == "Sem Tributos (Líquido ANEEL)":
            aliquota_efetiva = 0.0
        elif modo_impostos == "Cálculo por Dentro (art. 13 LC 87/96)":
            if soma_aliquotas < 99.0:
                aliquota_efetiva = ((1.0 / (1.0 - (soma_aliquotas / 100.0))) - 1.0) * 100.0
            else:
                aliquota_efetiva = soma_aliquotas
        else:
            aliquota_efetiva = soma_aliquotas

        st.caption(f"Gross-up Efetivo: **+{aliquota_efetiva:.2f}%**")

    # Objeto do Gerador configurado
    gerador_cfg = ParametrosGerador(
        preco_diesel_litro=preco_diesel,
        consumo_especifico_l_kwh=consumo_esp,
        custo_manutencao_kwh=custo_om,
        potencia_nominal_kva=pot_nominal_kva,
        potencia_continua_kw=pot_continua_kw,
        dias_uteis_mes=int(dias_uteis),
        horas_ponta_dia=float(horas_ponta_dia),
        capex_instalacao_reais=capex_investimento,
    )

# =============================================================================
# BLOCO 1: AMBIENTE DE CONTRATAÇÃO & INGESTÃO DA FATURA
# =============================================================================
col_upload, col_distrib = st.columns([1.8, 1.2])

with col_upload:
    st.markdown("#### 📄 1. Ingestão da Fatura & Ambiente de Contratação")

    # Seletor Explícito de Mercado Cativo (ACR) vs. Mercado Livre (ACL)
    c_amb1, c_amb2 = st.columns([1.2, 1.8])
    with c_amb1:
        amb_escolhido = st.radio(
            "Ambiente de Contratação:",
            options=["CATIVO", "LIVRE"],
            format_func=lambda x: "🏛️ Mercado Cativo (ACR)" if x == "CATIVO" else "⚡ Mercado Livre (ACL)",
            index=0 if st.session_state.fatura.ambiente_contratacao == "CATIVO" else 1,
            horizontal=False,
            help="Cativo: compra energia e rede da distribuidora (TUSD + TE + Bandeiras). Livre: paga apenas TUSD à distribuidora e energia via contrato bilateral.",
        )
        st.session_state.ambiente_selecionado = amb_escolhido

    with c_amb2:
        if amb_escolhido == "LIVRE":
            st.info("⚡ **Mercado Livre**: A fatura da distribuidora cobre apenas a **TUSD** (demanda e transporte de rede). A Tarifa de Energia (TE) e bandeiras não são cobradas pela concessionária.")
            preco_acl_input = st.number_input(
                "Preço do Contrato Livre de Energia (R$/kWh):",
                min_value=0.05,
                max_value=1.50,
                value=float(st.session_state.fatura.preco_energia_acl_kwh or 0.22),
                step=0.01,
                format="%.3f",
                help="Ex: R$ 0,22/kWh equivale a R$ 220,00/MWh.",
            )
        else:
            st.info("🏛️ **Mercado Cativo**: O consumidor é atendido integralmente pela concessionária local, pagando **TUSD** (rede) + **TE** (energia) + **Bandeiras Tarifárias**.")
            preco_acl_input = 0.0

    tab_pdf, tab_demo = st.tabs(["📤 Upload PDF da Fatura", "⚡ Faturas de Demonstração (ACR / ACL)"])

    with tab_pdf:
        uploaded_file = st.file_uploader(
            "Selecione ou arraste a fatura em PDF de qualquer distribuidora (CEMIG, CPFL, Enel, Light, etc.)",
            type=["pdf"],
            help="O parser agnóstico extrai demandas, consumos, ambiente e tributos sem depender de posições físicas fixas.",
        )
        if uploaded_file is not None:
            try:
                with st.spinner("Processando PDF e extraindo variáveis com pdfplumber..."):
                    parsed_fatura = invoice_parser.parse_pdf(uploaded_file)
                    st.session_state.fatura = parsed_fatura
                    st.session_state.ambiente_selecionado = parsed_fatura.ambiente_contratacao
                    if parsed_fatura.concessionaria_homologada_aneel:
                        st.session_state.distribuidora_selecionada = parsed_fatura.concessionaria_homologada_aneel
                    st.success(
                        f"PDF processado com sucesso! Distribuidora: **{parsed_fatura.distribuidora_identificada}** | "
                        f"Ambiente: **{parsed_fatura.ambiente_contratacao}** "
                        f"(Confiança da leitura: {parsed_fatura.confidence_score * 100:.0f}%)"
                    )
            except Exception as e:
                st.error(f"Erro ao processar o arquivo PDF: {e}")

    with tab_demo:
        sample_dict = invoice_parser.get_sample_invoices()
        sample_selected = st.selectbox(
            "Carregar caso demonstrativo do Grupo A4:",
            options=list(sample_dict.keys()),
            index=0,
        )
        if st.button("Carregar Dados Desta Fatura", use_container_width=True):
            st.session_state.fatura = sample_dict[sample_selected]
            st.session_state.distribuidora_selecionada = st.session_state.fatura.distribuidora_identificada
            st.session_state.ambiente_selecionado = st.session_state.fatura.ambiente_contratacao
            st.success(f"Carregado: {sample_selected}")
            st.rerun()

with col_distrib:
    st.markdown("#### 🏛️ 2. Base Tarifária ANEEL Homologada")
    all_concessionarias = tariff_client.list_concessionarias("A4")

    curr_dist = st.session_state.fatura.concessionaria_homologada_aneel or st.session_state.fatura.distribuidora_identificada
    matched_idx = 0
    if curr_dist in all_concessionarias:
        matched_idx = all_concessionarias.index(curr_dist)
    else:
        for i, a in enumerate(all_concessionarias):
            if curr_dist.upper() in a.upper() or a.upper() in curr_dist.upper():
                matched_idx = i
                break

    dist_escolhida = st.selectbox(
        "Distribuidora / Concessionária A4:",
        options=all_concessionarias,
        index=matched_idx,
        help="Tarifas oficiais consultadas de https://englucas1.github.io/Aneel-Tarifas_v3/",
    )
    st.session_state.distribuidora_selecionada = dist_escolhida

    try:
        dados_tarifarios = tariff_client.get_tarifas(dist_escolhida, subgrupo=st.session_state.fatura.subgrupo)
        st.markdown(
            f"""
            <div style="background: #0f172a; border-radius: 8px; padding: 12px 16px; border: 1px solid #1e293b; font-size: 0.85rem;">
                <div style="color: #38bdf8; font-weight: 700; margin-bottom: 4px;">{dados_tarifarios.agente} | Subgrupo {dados_tarifarios.subgrupo}</div>
                <div style="color: #94a3b8; font-size: 0.78rem;">{dados_tarifarios.resolucao_homologatoria}</div>
                <div style="margin-top: 6px; display: flex; gap: 12px; font-size: 0.8rem;">
                    <span>Demanda Verde: <strong>R$ {dados_tarifarios.verde.demanda_unica:.2f}/kW</strong></span>
                    <span>TUSD Ponta: <strong>R$ {dados_tarifarios.verde.tusd_ponta:.4f}</strong></span>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    except Exception as ex:
        st.error(f"Erro ao obter tarifas para {dist_escolhida}: {ex}")
        st.stop()

# =============================================================================
# BLOCO 2: CONFERÊNCIA DOS DADOS EXTRAÍDOS DA FATURA (EDITÁVEL)
# =============================================================================
with st.expander("🔍 Conferência & Ajuste Fino dos Dados da Fatura Atual", expanded=False):
    f_curr = st.session_state.fatura
    c_f1, c_f2, c_f3, c_f4 = st.columns(4)

    with c_f1:
        mod_edit = st.selectbox("Modalidade Atual:", options=["VERDE", "AZUL"], index=0 if f_curr.modalidade == "VERDE" else 1)
        subg_edit = st.selectbox("Subgrupo:", options=["A4", "A3a"], index=0 if f_curr.subgrupo == "A4" else 1)
        mes_edit = st.text_input("Mês de Referência:", value=f_curr.mes_referencia)

    with c_f2:
        dem_contr_edit = st.number_input("Demanda Contratada (kW):", min_value=30.0, max_value=20000.0, value=float(f_curr.demanda_contratada_kw), step=10.0)
        dem_med_edit = st.number_input("Demanda Medida (kW):", min_value=30.0, max_value=20000.0, value=float(f_curr.demanda_medida_kw), step=10.0)

    with c_f3:
        cons_p_edit = st.number_input("Consumo Ponta (kWh):", min_value=0.0, max_value=2000000.0, value=float(f_curr.consumo_ponta_kwh), step=500.0)
        cons_fp_edit = st.number_input("Consumo Fora Ponta (kWh):", min_value=0.0, max_value=5000000.0, value=float(f_curr.consumo_fora_ponta_kwh), step=1000.0)

    with c_f4:
        val_tot_edit = st.number_input("Valor Total Faturado pela Concessionária (R$):", min_value=0.0, max_value=10000000.0, value=float(f_curr.valor_total_faturado), step=1000.0)
        cip_edit = st.number_input("Iluminação Pública CIP (R$):", min_value=0.0, max_value=20000.0, value=float(f_curr.tributos.cip_cosip), step=20.0)

    if st.button("Atualizar Dados da Fatura"):
        st.session_state.fatura.modalidade = mod_edit
        st.session_state.fatura.subgrupo = subg_edit
        st.session_state.fatura.mes_referencia = mes_edit
        st.session_state.fatura.demanda_contratada_kw = dem_contr_edit
        st.session_state.fatura.demanda_medida_kw = dem_med_edit
        st.session_state.fatura.consumo_ponta_kwh = cons_p_edit
        st.session_state.fatura.consumo_fora_ponta_kwh = cons_fp_edit
        st.session_state.fatura.valor_total_faturado = val_tot_edit
        st.session_state.fatura.tributos.cip_cosip = cip_edit
        st.success("Dados da fatura atualizados com sucesso!")
        st.rerun()

# =============================================================================
# BLOCO 3: EXECUÇÃO DO MOTOR DE CÁLCULO E ANÁLISE DE CENÁRIOS
# =============================================================================
resultado_sim: ResultadoSimulacao = tariff_engine.simular(
    fatura=st.session_state.fatura,
    gerador=gerador_cfg,
    concessionaria_nome=st.session_state.distribuidora_selecionada,
    aliquota_impostos_override=aliquota_efetiva,
    ambiente_contratacao_override=st.session_state.ambiente_selecionado,
    preco_energia_acl_override=preco_acl_input if st.session_state.ambiente_selecionado == "LIVRE" else 0.0,
)

# =============================================================================
# BLOCO 4: AUDITORIA DE CONFORMIDADE DA FATURA ATUAL (CONFRONTO ANEEL)
# =============================================================================
audit_box_style = "audit-box-ok" if resultado_sim.fatura_auditada_conforme else "audit-box-warn"
audit_badge = "✅ FATURA CONFORME COM REGULAÇÃO ANEEL" if resultado_sim.fatura_auditada_conforme else "⚠️ DIVERGÊNCIA IDENTIFICADA NA FATURA"
audit_badge_color = "#34d399" if resultado_sim.fatura_auditada_conforme else "#f59e0b"

st.markdown(
    f"""
    <div class="{audit_box_style}">
        <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px; margin-bottom: 8px;">
            <span style="color: {audit_badge_color}; font-weight: 800; font-size: 0.95rem; text-transform: uppercase;">
                {audit_badge} &bull; Ambiente {resultado_sim.ambiente_contratacao}
            </span>
            <span style="font-size: 0.85rem; color: #94a3b8;">
                Homologação: {resultado_sim.resolucao_homologatoria[:55]}...
            </span>
        </div>
        <p style="color: #f1f5f9; font-size: 0.95rem; margin-bottom: 12px; line-height: 1.5;">
            {resultado_sim.diagnostico_auditoria}
        </p>
        <div style="display: flex; gap: 24px; flex-wrap: wrap; font-size: 0.9rem;">
            <div>Valor Faturado na Conta: <strong style="color: #f8fafc;">R$ {resultado_sim.valor_fatura_cobrado_real:,.2f}</strong></div>
            <div>Valor Regulamentar ANEEL: <strong style="color: #38bdf8;">R$ {resultado_sim.valor_fatura_reconstituida_reais:,.2f}</strong></div>
            <div>Diferença Apurada: <strong style="color: {audit_badge_color};">R$ {resultado_sim.desvio_auditoria_reais:,.2f} ({resultado_sim.desvio_auditoria_percentual:+.2f}%)</strong></div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# =============================================================================
# BLOCO 5: CARD DE RECOMENDAÇÃO AUTOMÁTICA (DESTAQUE MÁXIMO)
# =============================================================================
card_css_class = "recommendation-card" if resultado_sim.modalidade_recomendada == "VERDE" else "recommendation-card-azul"
icone_rec = "🟢" if resultado_sim.modalidade_recomendada == "VERDE" else "🔵"

st.markdown(
    f"""
    <div class="{card_css_class}">
        <div style="display: flex; justify-content: space-between; align-items: flex-start; flex-wrap: wrap; gap: 16px;">
            <div>
                <span style="background: rgba(255, 255, 255, 0.15); color: #ffffff; padding: 4px 12px; border-radius: 20px; font-size: 0.8rem; font-weight: 700; text-transform: uppercase;">
                    Decisão Otimizada de Enquadramento ({resultado_sim.ambiente_contratacao})
                </span>
                <h2 style="color: #ffffff; font-size: 1.85rem; font-weight: 800; margin: 8px 0 4px 0;">
                    {icone_rec} Enquadramento Recomendado: TARIFA HORÁRIA {resultado_sim.modalidade_recomendada}
                </h2>
                <p style="color: #e2e8f0; font-size: 1.0rem; max-width: 900px; line-height: 1.5; margin-bottom: 0;">
                    {resultado_sim.justificativa_tecnica}
                </p>
            </div>
            <div style="text-align: right; background: rgba(0, 0, 0, 0.3); border-radius: 12px; padding: 14px 22px; border: 1px solid rgba(255, 255, 255, 0.1);">
                <div style="color: #94a3b8; font-size: 0.8rem; font-weight: 600; text-transform: uppercase;">Economia Mensal Líquida</div>
                <div style="color: #34d399; font-size: 1.9rem; font-weight: 900;">R$ {resultado_sim.economia_mensal_liquida_reais:,.2f}</div>
                <div style="color: #38bdf8; font-size: 0.95rem; font-weight: 700;">{resultado_sim.economia_mensal_percentual:+.1f}% na despesa operacional</div>
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# =============================================================================
# BLOCO 6: LINHA DE KPIS EXECUTIVOS
# =============================================================================
kpi_c1, kpi_c2, kpi_c3, kpi_c4 = st.columns(4)

with kpi_c1:
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">Despesa Total Atual (Cen. 0)</div>
            <div class="kpi-value">R$ {resultado_sim.cenario_0_atual.custo_total_operacional_reais:,.2f}</div>
            <div class="kpi-delta-neutral">Fatura Concessionária + Contrato ACL</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with kpi_c2:
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">Projetado Verde + Gerador</div>
            <div class="kpi-value">R$ {resultado_sim.cenario_1_verde.custo_total_operacional_reais:,.2f}</div>
            <div class="kpi-delta-positive">Rede: R$ {resultado_sim.cenario_1_verde.custo_fatura_concessionaria_reais:,.0f} | Gerador: R$ {resultado_sim.cenario_1_verde.gasto_gerador_total_reais:,.0f}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with kpi_c3:
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">Projetado Azul + Gerador</div>
            <div class="kpi-value">R$ {resultado_sim.cenario_2_azul.custo_total_operacional_reais:,.2f}</div>
            <div class="kpi-delta-positive">Rede: R$ {resultado_sim.cenario_2_azul.custo_fatura_concessionaria_reais:,.0f} | Gerador: R$ {resultado_sim.cenario_2_azul.gasto_gerador_total_reais:,.0f}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with kpi_c4:
    payback_txt = f"{resultado_sim.payback_meses:.1f} meses" if resultado_sim.payback_meses else "Imediato / OPEX"
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">Economia Anualizada</div>
            <div class="kpi-value" style="color: #10b981;">R$ {resultado_sim.economia_anual_projetada_reais:,.2f}</div>
            <div class="kpi-delta-neutral">Payback Projetado: {payback_txt}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown("<div style='height: 16px;'></div>", unsafe_allow_html=True)

# =============================================================================
# BLOCO 7: ABAS ANALÍTICAS (GRÁFICOS, TARIFAS UTILIZADAS E TRIBUTOS)
# =============================================================================
tab_graf1, tab_graf2, tab_memoria, tab_tarifas_todas, tab_tributos_det, tab_base_aneel = st.tabs([
    "📊 Comparativo dos Cenários",
    "⚡ Margem Unitária & Sensibilidade",
    "📑 Memória de Cálculo",
    "🏷️ Todas as Tarifas Utilizadas",
    "⚖️ Tributos Detalhados",
    "🏛️ Base Homologada ANEEL",
])

with tab_graf1:
    col_chart1, col_chart2 = st.columns([1.6, 1.2])

    with col_chart1:
        cenarios_nomes = ["Atual (Cenário 0)", "Verde + Gerador (Cen. 1)", "Azul + Gerador (Cen. 2)"]
        custos_demanda = [
            resultado_sim.cenario_0_atual.custo_demanda_total_reais,
            resultado_sim.cenario_1_verde.custo_demanda_total_reais,
            resultado_sim.cenario_2_azul.custo_demanda_total_reais,
        ]
        custos_fp = [
            resultado_sim.cenario_0_atual.custo_consumo_fp_total_reais,
            resultado_sim.cenario_1_verde.custo_consumo_fp_total_reais,
            resultado_sim.cenario_2_azul.custo_consumo_fp_total_reais,
        ]
        custos_ponta_rede = [
            resultado_sim.cenario_0_atual.custo_consumo_ponta_rede_reais,
            resultado_sim.cenario_1_verde.custo_consumo_ponta_rede_reais,
            resultado_sim.cenario_2_azul.custo_consumo_ponta_rede_reais,
        ]
        custos_acl = [
            resultado_sim.cenario_0_atual.custo_energia_livre_reais,
            resultado_sim.cenario_1_verde.custo_energia_livre_reais,
            resultado_sim.cenario_2_azul.custo_energia_livre_reais,
        ]
        custos_gerador = [
            0.0,
            resultado_sim.cenario_1_verde.gasto_gerador_total_reais,
            resultado_sim.cenario_2_azul.gasto_gerador_total_reais,
        ]

        fig_comp = go.Figure()
        fig_comp.add_trace(go.Bar(name="Demanda (Distribuidora)", x=cenarios_nomes, y=custos_demanda, marker_color="#0284c7"))
        fig_comp.add_trace(go.Bar(name="Consumo FP (Distribuidora)", x=cenarios_nomes, y=custos_fp, marker_color="#0d9488"))
        fig_comp.add_trace(go.Bar(name="Consumo Ponta (Rede Residual)", x=cenarios_nomes, y=custos_ponta_rede, marker_color="#e11d48"))
        if resultado_sim.ambiente_contratacao == "LIVRE":
            fig_comp.add_trace(go.Bar(name="Contrato Bilateral Livre (ACL)", x=cenarios_nomes, y=custos_acl, marker_color="#8b5cf6"))
        fig_comp.add_trace(go.Bar(name="OPEX Gerador (Diesel + O&M)", x=cenarios_nomes, y=custos_gerador, marker_color="#f59e0b"))

        fig_comp.update_layout(
            barmode="stack",
            title=f"<b>Composição da Despesa Mensal por Cenário ({resultado_sim.ambiente_contratacao})</b>",
            template="plotly_dark",
            paper_bgcolor="#0b1329",
            plot_bgcolor="#0f172a",
            margin=dict(l=40, r=20, t=50, b=40),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            yaxis=dict(title="Despesa Mensal (R$)", gridcolor="#1e293b"),
            xaxis=dict(gridcolor="#1e293b"),
        )
        st.plotly_chart(fig_comp, use_container_width=True)

    with col_chart2:
        diff_rede = (
            resultado_sim.cenario_0_atual.custo_consumo_ponta_rede_reais
            - resultado_sim.cenario_1_verde.custo_consumo_ponta_rede_reais
        )
        fig_eco = go.Figure(go.Waterfall(
            name="Economia",
            orientation="v",
            measure=["absolute", "relative", "relative", "total"],
            x=["Fatura Atual", "Corte Energia Ponta", "Custo Gerador", "Custo Otimizado"],
            textposition="outside",
            y=[
                resultado_sim.custo_total_atual_reais,
                -diff_rede,
                resultado_sim.cenario_1_verde.gasto_gerador_total_reais,
                resultado_sim.custo_total_recomendado_reais,
            ],
            connector={"line": {"color": "#64748b"}},
            decreasing={"marker": {"color": "#10b981"}},
            increasing={"marker": {"color": "#ef4444"}},
            totals={"marker": {"color": "#38bdf8"}},
        ))
        fig_eco.update_layout(
            title="<b>Impacto do Despacho do Gerador (R$)</b>",
            template="plotly_dark",
            paper_bgcolor="#0b1329",
            plot_bgcolor="#0f172a",
            margin=dict(l=30, r=20, t=50, b=40),
            yaxis=dict(gridcolor="#1e293b"),
        )
        st.plotly_chart(fig_eco, use_container_width=True)

with tab_graf2:
    col_marg1, col_marg2 = st.columns([1.4, 1.4])

    with col_marg1:
        st.markdown("#### ⚖️ Comparação Direta de Custos Unitários de Ponta (R$/kWh)")
        st.caption("Confronto entre o custo marginal do gerador térmico e o custo unitário evitado na rede no horário de ponta.")

        c_ger_unit = gerador_cfg.custo_gerador_unitario
        t_rede_verde = resultado_sim.comparativo_custo_ponta["tarifa_rede_verde_ponta_com_impostos"]
        t_rede_azul = resultado_sim.comparativo_custo_ponta["tarifa_rede_azul_ponta_com_impostos"]

        df_unit = pd.DataFrame({
            "Origem da Energia": ["Gerador Térmico (Diesel + O&M)", f"Rede Ponta Verde ({resultado_sim.ambiente_contratacao})", f"Rede Ponta Azul ({resultado_sim.ambiente_contratacao})"],
            "Custo Unitário (R$/kWh)": [c_ger_unit, t_rede_verde, t_rede_azul],
        })

        fig_unit = px.bar(
            df_unit,
            x="Origem da Energia",
            y="Custo Unitário (R$/kWh)",
            color="Origem da Energia",
            color_discrete_map={
                "Gerador Térmico (Diesel + O&M)": "#f59e0b",
                f"Rede Ponta Verde ({resultado_sim.ambiente_contratacao})": "#10b981",
                f"Rede Ponta Azul ({resultado_sim.ambiente_contratacao})": "#38bdf8",
            },
            text_auto=".3f",
        )
        fig_unit.update_layout(
            template="plotly_dark",
            paper_bgcolor="#0b1329",
            plot_bgcolor="#0f172a",
            showlegend=False,
            yaxis=dict(gridcolor="#1e293b", title="R$/kWh"),
            margin=dict(l=20, r=20, t=30, b=30),
        )
        st.plotly_chart(fig_unit, use_container_width=True)

    with col_marg2:
        st.markdown("#### 💡 Análise de Sensibilidade do Preço do Diesel")
        st.caption("Ponto de equilíbrio (Break-even) onde o diesel atinge o custo da rede na ponta.")

        diesel_range = [x * 0.25 for x in range(18, 37)]  # 4.50 a 9.00
        custos_g = [(d * gerador_cfg.consumo_especifico_l_kwh) + gerador_cfg.custo_manutencao_kwh for d in diesel_range]

        fig_sens = go.Figure()
        fig_sens.add_trace(go.Scatter(
            x=diesel_range,
            y=custos_g,
            mode="lines+markers",
            name="Custo Unitário Gerador",
            line=dict(color="#f59e0b", width=3),
        ))
        fig_sens.add_trace(go.Scatter(
            x=diesel_range,
            y=[t_rede_verde] * len(diesel_range),
            mode="lines",
            name="Custo Rede Verde (Ponta)",
            line=dict(color="#10b981", width=2, dash="dash"),
        ))

        fig_sens.update_layout(
            title="<b>Preço do Diesel vs. Ponto de Equilíbrio (Break-even)</b>",
            template="plotly_dark",
            paper_bgcolor="#0b1329",
            plot_bgcolor="#0f172a",
            xaxis=dict(title="Preço do Óleo Diesel (R$/L)", gridcolor="#1e293b"),
            yaxis=dict(title="Custo Unitário (R$/kWh)", gridcolor="#1e293b"),
            margin=dict(l=20, r=20, t=40, b=30),
        )
        st.plotly_chart(fig_sens, use_container_width=True)

with tab_memoria:
    st.markdown("#### 📋 Memória de Cálculo Passo a Passo (Normas ANEEL)")
    st.caption(f"Detalhamento analítico das grandezas físicas e financeiras para o {resultado_sim.ambiente_contratacao}.")

    tabela_memoria = [
        {"Componente": "Demanda Fora Ponta / Única (kW)", "Cenário 0 (Atual)": f"{resultado_sim.cenario_0_atual.demanda_contratada_fora_ponta_kw:.1f}", "Cenário 1 (Verde)": f"{resultado_sim.cenario_1_verde.demanda_contratada_fora_ponta_kw:.1f}", "Cenário 2 (Azul)": f"{resultado_sim.cenario_2_azul.demanda_contratada_fora_ponta_kw:.1f}"},
        {"Componente": "Demanda Ponta Contratada (kW)", "Cenário 0 (Atual)": f"{resultado_sim.cenario_0_atual.demanda_contratada_ponta_kw or 0.0:.1f}", "Cenário 1 (Verde)": "-", "Cenário 2 (Azul)": f"{resultado_sim.cenario_2_azul.demanda_contratada_ponta_kw or 0.0:.1f}"},
        {"Componente": "Custo Demanda Contratada (R$)", "Cenário 0 (Atual)": f"R$ {resultado_sim.cenario_0_atual.custo_demanda_total_reais:,.2f}", "Cenário 1 (Verde)": f"R$ {resultado_sim.cenario_1_verde.custo_demanda_total_reais:,.2f}", "Cenário 2 (Azul)": f"R$ {resultado_sim.cenario_2_azul.custo_demanda_total_reais:,.2f}"},
        {"Componente": "Consumo Fora Ponta Rede (kWh)", "Cenário 0 (Atual)": f"{resultado_sim.cenario_0_atual.consumo_fora_ponta_kwh:,.0f}", "Cenário 1 (Verde)": f"{resultado_sim.cenario_1_verde.consumo_fora_ponta_kwh:,.0f}", "Cenário 2 (Azul)": f"{resultado_sim.cenario_2_azul.consumo_fora_ponta_kwh:,.0f}"},
        {"Componente": "Custo Consumo Fora Ponta (R$)", "Cenário 0 (Atual)": f"R$ {resultado_sim.cenario_0_atual.custo_consumo_fp_total_reais:,.2f}", "Cenário 1 (Verde)": f"R$ {resultado_sim.cenario_1_verde.custo_consumo_fp_total_reais:,.2f}", "Cenário 2 (Azul)": f"R$ {resultado_sim.cenario_2_azul.custo_consumo_fp_total_reais:,.2f}"},
        {"Componente": "Consumo Ponta Suprido pelo Gerador (kWh)", "Cenário 0 (Atual)": "0", "Cenário 1 (Verde)": f"{resultado_sim.cenario_1_verde.consumo_ponta_gerador_kwh:,.0f}", "Cenário 2 (Azul)": f"{resultado_sim.cenario_2_azul.consumo_ponta_gerador_kwh:,.0f}"},
        {"Componente": "Consumo Ponta Residual da Rede (kWh)", "Cenário 0 (Atual)": f"{resultado_sim.cenario_0_atual.consumo_ponta_rede_kwh:,.0f}", "Cenário 1 (Verde)": f"{resultado_sim.cenario_1_verde.consumo_ponta_rede_kwh:,.0f}", "Cenário 2 (Azul)": f"{resultado_sim.cenario_2_azul.consumo_ponta_rede_kwh:,.0f}"},
        {"Componente": "Custo Consumo Ponta da Rede (R$)", "Cenário 0 (Atual)": f"R$ {resultado_sim.cenario_0_atual.custo_consumo_ponta_rede_reais:,.2f}", "Cenário 1 (Verde)": f"R$ {resultado_sim.cenario_1_verde.custo_consumo_ponta_rede_reais:,.2f}", "Cenário 2 (Azul)": f"R$ {resultado_sim.cenario_2_azul.custo_consumo_ponta_rede_reais:,.2f}"},
        {"Componente": "Contrato Bilateral Livre ACL (R$)", "Cenário 0 (Atual)": f"R$ {resultado_sim.cenario_0_atual.custo_energia_livre_reais:,.2f}", "Cenário 1 (Verde)": f"R$ {resultado_sim.cenario_1_verde.custo_energia_livre_reais:,.2f}", "Cenário 2 (Azul)": f"R$ {resultado_sim.cenario_2_azul.custo_energia_livre_reais:,.2f}"},
        {"Componente": "Gasto Operacional Skid Gerador (R$)", "Cenário 0 (Atual)": "R$ 0,00", "Cenário 1 (Verde)": f"R$ {resultado_sim.cenario_1_verde.gasto_gerador_total_reais:,.2f}", "Cenário 2 (Azul)": f"R$ {resultado_sim.cenario_2_azul.gasto_gerador_total_reais:,.2f}"},
        {"Componente": "Fatura Distribuidora / Concessionária (R$)", "Cenário 0 (Atual)": f"R$ {resultado_sim.cenario_0_atual.custo_fatura_concessionaria_reais:,.2f}", "Cenário 1 (Verde)": f"R$ {resultado_sim.cenario_1_verde.custo_fatura_concessionaria_reais:,.2f}", "Cenário 2 (Azul)": f"R$ {resultado_sim.cenario_2_azul.custo_fatura_concessionaria_reais:,.2f}"},
        {"Componente": "DESPESA OPERACIONAL LÍQUIDA TOTAL (R$)", "Cenário 0 (Atual)": f"R$ {resultado_sim.cenario_0_atual.custo_total_operacional_reais:,.2f}", "Cenário 1 (Verde)": f"R$ {resultado_sim.cenario_1_verde.custo_total_operacional_reais:,.2f}", "Cenário 2 (Azul)": f"R$ {resultado_sim.cenario_2_azul.custo_total_operacional_reais:,.2f}"},
    ]
    st.dataframe(pd.DataFrame(tabela_memoria), use_container_width=True, hide_index=True)

with tab_tarifas_todas:
    st.markdown("#### 🏷️ Tabela Exaustiva de Todas as Tarifas Utilizadas")
    st.caption("Visão detalhada de cada componente tarifário, base de cálculo sem impostos e valor efetivo praticado com impostos.")

    df_tarifas_utilizadas = pd.DataFrame(resultado_sim.tabela_tarifas_utilizadas)
    df_tarifas_utilizadas.rename(columns={
        "Componente": "Componente Tarifário",
        "Unidade": "Unidade",
        "Modalidade": "Modalidade",
        "Ambiente": "Ambiente",
        "Tarifa_Homologada_Sem_Tributos": "Tarifa Homologada ANEEL (Sem Tributos)",
        "Tarifa_Efetiva_Com_Tributos": "Tarifa Efetiva Praticada (Com Tributos)",
        "Onde_Faturado": "Local de Faturamento / Incidência",
    }, inplace=True)
    st.dataframe(df_tarifas_utilizadas, use_container_width=True, hide_index=True)

with tab_tributos_det:
    st.markdown("#### ⚖️ Detalhamento Fiscal de Tributos & Gross-up Legal")
    st.caption("Apuração individualizada de ICMS, PIS, COFINS, CIP e aplicação da sistemática de cálculo 'por dentro' (art. 13 da LC 87/96).")

    df_tributos_utilizados = pd.DataFrame(resultado_sim.tabela_tributos_utilizados)
    df_tributos_utilizados.rename(columns={
        "Tributo": "Tributo / Encargo",
        "Base_Calculo_Reais": "Base de Cálculo (R$)",
        "Aliquota_Nominal_Percent": "Alíquota Nominal (%)",
        "Metodologia": "Metodologia de Incidência Legal",
        "Valor_Apurado_Reais": "Valor do Tributo Apurado (R$)",
    }, inplace=True)
    st.dataframe(df_tributos_utilizados, use_container_width=True, hide_index=True)

with tab_base_aneel:
    st.markdown(f"#### 🏛️ Base Tarifária Homologada ANEEL: {dados_tarifarios.agente} ({dados_tarifarios.subgrupo})")
    st.caption(f"Resolução Homologatória Vigente: **{dados_tarifarios.resolucao_homologatoria}** (Data de Vigência: {dados_tarifarios.data_vigencia or '2026'})")

    tarifa_comp_df = pd.DataFrame([
        {"Componente": "Demanda Ponta", "Unidade": "R$/kW", "Verde": "-", "Azul": f"{dados_tarifarios.azul.demanda_ponta:.2f}"},
        {"Componente": "Demanda Fora Ponta / Única", "Unidade": "R$/kW", "Verde": f"{dados_tarifarios.verde.demanda_unica:.2f}", "Azul": f"{dados_tarifarios.azul.demanda_fora_ponta:.2f}"},
        {"Componente": "TUSD Ponta", "Unidade": "R$/kWh", "Verde": f"{dados_tarifarios.verde.tusd_ponta:.5f}", "Azul": f"{dados_tarifarios.azul.tusd_ponta:.5f}"},
        {"Componente": "TUSD Fora Ponta", "Unidade": "R$/kWh", "Verde": f"{dados_tarifarios.verde.tusd_fora_ponta:.5f}", "Azul": f"{dados_tarifarios.azul.tusd_fora_ponta:.5f}"},
        {"Componente": "TE Ponta", "Unidade": "R$/kWh", "Verde": f"{dados_tarifarios.verde.te_ponta:.5f}", "Azul": f"{dados_tarifarios.azul.te_ponta:.5f}"},
        {"Componente": "TE Fora Ponta", "Unidade": "R$/kWh", "Verde": f"{dados_tarifarios.verde.te_fora_ponta:.5f}", "Azul": f"{dados_tarifarios.azul.te_fora_ponta:.5f}"},
        {"Componente": "TOTAL ENERGIA PONTA (TUSD+TE)", "Unidade": "R$/kWh", "Verde": f"{dados_tarifarios.verde.total_ponta:.5f}", "Azul": f"{dados_tarifarios.azul.total_ponta:.5f}"},
        {"Componente": "TOTAL ENERGIA FORA PONTA (TUSD+TE)", "Unidade": "R$/kWh", "Verde": f"{dados_tarifarios.verde.total_fora_ponta:.5f}", "Azul": f"{dados_tarifarios.azul.total_fora_ponta:.5f}"},
    ])
    st.dataframe(tarifa_comp_df, use_container_width=True, hide_index=True)

# =============================================================================
# BLOCO 8: BOTÃO DE EXPORTAÇÃO EXCEL (.XLSX) COM TODAS AS TARIFAS E TRIBUTOS
# =============================================================================
st.markdown("<hr style='border-color: #1e293b; margin: 32px 0 20px 0;'>", unsafe_allow_html=True)
c_exp1, c_exp2 = st.columns([2, 1])

with c_exp1:
    st.markdown("### 📥 Exportar Memória Completa com Tarifas e Tributos")
    st.write(
        "Baixe a planilha detalhada em formato **Microsoft Excel (.xlsx)** contendo todas as 4 abas estruturadas: "
        "**1. Resumo & Auditoria** (confronto Cativo/Livre e ANEEL), **2. Memória de Cálculo**, **3. Tarifas Utilizadas** (com e sem tributos) e **4. Tributos Detalhados** (ICMS, PIS, COFINS e Gross-up)."
    )

with c_exp2:
    excel_bytes = gerar_relatorio_excel(
        resultado=resultado_sim,
        fatura=st.session_state.fatura,
        gerador=gerador_cfg,
        tarifas_homologadas=dados_tarifarios,
    )
    nome_arquivo = f"Relatorio_Auditoria_PeakShaving_{st.session_state.ambiente_selecionado}_{st.session_state.distribuidora_selecionada.replace(' ', '_')}.xlsx"

    st.download_button(
        label="📥 Baixar Planilha Excel Completa (.xlsx)",
        data=excel_bytes,
        file_name=nome_arquivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
