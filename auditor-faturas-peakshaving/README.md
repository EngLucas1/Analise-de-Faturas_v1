# ⚡ Auditor Tarifário & Simulador de Peak Shaving (Grupo A4)

Aplicação web analítica completa desenvolvida para Engenharia de Energia e Gestão do Setor Elétrico Brasileiro. Permite auditar faturas de energia de consumidores industriais e comerciais do **Grupo A4**, consultar tarifas homologadas da ANEEL em tempo real via **[Aneel-Tarifas_v3](https://englucas1.github.io/Aneel-Tarifas_v3/)** e simular com precisão a viabilidade técnica e financeira da operação de grupos geradores térmicos no horário de ponta (**Peak Shaving**), comparando os enquadramentos **Tarifa Verde** e **Tarifa Azul** nos ambientes **Cativo (ACR)** e **Livre (ACL)**.

---

## 🎯 Principais Funcionalidades

1. **Auditoria de Fatura: Mercado Cativo (ACR) vs. Mercado Livre (ACL)**:
   - **Mercado Cativo (ACR)**: Reconstitui a cobrança total da concessionária considerando o uso da rede (**TUSD**), a molécula de energia (**TE**), adicionais de **Bandeiras Tarifárias** e tributos com alíquotas oficiais.
   - **Mercado Livre (ACL)**: Entende que a fatura da distribuidora refere-se estritamente ao Contrato de Uso do Sistema de Distribuição (**CUSD** / apenas **TUSD**). A TE e as bandeiras da concessionária são isentas nesta fatura. O custo da energia é parametrizado via contrato bilateral livre (PPA).
   - **Diagnóstico de Conformidade**: Aponta automaticamente desvios entre o valor cobrado e o valor homologado pela ANEEL, alertando sobre possíveis cobranças a maior ou créditos.

2. **Parser de Faturas em PDF Agnóstico de Layout**:
   - Extrai texto e dados estruturados sem amarras a coordenadas fixas via `pdfplumber` e regex heurístico.
   - Captura: Distribuidora/Concessionária, Ambiente (ACR/ACL), Subgrupo (A4/A3a), Modalidade (Verde/Azul), Mês de Referência, Demanda Contratada (kW), Demanda Medida (kW), Consumos Ponta e Fora de Ponta (kWh), Bandeiras Tarifárias e Tributos (ICMS, PIS, COFINS, CIP).
   - Acompanha faturas demonstrativas integradas (CEMIG Cativo, Enel SP Livre, CPFL Paulista Livre, Equatorial Pará Cativo).

3. **Conexão Oficial com a Base de Tarifas Homologadas da ANEEL**:
   - Conecta-se à aplicação web `https://englucas1.github.io/Aneel-Tarifas_v3/`.
   - Consulta e sincroniza as tarifas homologadas vigentes para mais de 100 concessionárias brasileiras:
     - **Modalidade Verde**: Tarifa de Demanda Única (R$/kW), TUSD Ponta, TUSD Fora Ponta, TE Ponta e TE Fora Ponta (R$/kWh).
     - **Modalidade Azul**: Tarifa de Demanda Ponta (R$/kW), Demanda Fora Ponta (R$/kW), TUSD Ponta/FP e TE Ponta/FP (R$/kWh).

4. **Tabela Exaustiva de Tarifas e Auditoria Fiscal de Tributos**:
   - Apresenta na tela e em planilha Excel o valor de **todas as tarifas utilizadas** com e sem tributos (Demanda, TUSD Ponta/FP, TE Ponta/FP, Contrato Livre ACL e Custo Marginal do Gerador).
   - Detalha passo a passo os tributos: Bases de cálculo, Alíquotas nominais, Metodologia legal de incidência (**Cálculo por Dentro** conforme art. 13 da LC 87/96 para ICMS e Gross-up) e valores apurados em R$.

5. **Parametrização Operacional do Gerador (OPEX Térmico)**:
   - Preço do Óleo Diesel ($R\$/L$), Consumo Específico ($L/kWh$), Custo de Manutenção / O&M ($R\$/kWh$), Potência Contínua ($kW$ e $kVA$) e Regime de despacho (22 dias úteis, 3h/dia).
   - Fórmula do Custo Marginal Unitário:
     $$Custo_{gerador} (R\$/kWh) = (Preço_{Diesel} \times Consumo_{esp}) + Custo_{manutenção}$$

6. **Motor Matemático e Análise Comparativa de Cenários**:
   - **Cenário 0 (Fatura Atual)**: Auditoria e reconstituição analítica da cobrança.
   - **Cenário 1 (Tarifa Verde + Peak Shaving)**: Demanda única dimensionada para carga fora ponta e absorção do consumo de ponta da rede pela geração própria, cortando a elevada TUSD Ponta.
   - **Cenário 2 (Tarifa Azul + Peak Shaving Otimizado)**: Demanda Fora Ponta mantida e Demanda Ponta rebaixada ao mínimo regulatório (30 kW ou carga residual), eliminando a pesada Tarifa de Demanda Ponta ($R\$/kW$).

7. **Exportação de Relatório Executivo (.xlsx com 4 Abas)**:
   - **Aba 1 (Resumo & Auditoria)**: Confronto Cativo/Livre, Diagnóstico de Conformidade e KPIs.
   - **Aba 2 (Memória de Cálculo)**: Passo a passo de demandas, consumos, custos térmicos e tributos.
   - **Aba 3 (Tarifas Utilizadas)**: Tabela exaustiva de todas as tarifas com e sem impostos.
   - **Aba 4 (Tributos Detalhados)**: Apuração fiscal completa de ICMS, PIS, COFINS, CIP e Gross-up.

---

## 📂 Estrutura Modular do Projeto

```
auditor-faturas-peakshaving/
├── app.py                  # Dashboard Streamlit e interface analítica completa
├── tariff_client.py        # Conector HTTP/Scraper e resolvedor de tarifas ANEEL
├── invoice_parser.py       # Parser agnóstico de faturas PDF com detecção ACR/ACL
├── tariff_engine.py        # Motor matemático das regras ANEEL, auditoria e simulação
├── export_report.py        # Exportador corporativo para planilhas Excel (.xlsx com 4 abas)
├── Iniciar_Aplicacao.bat   # Script Windows para inicialização em 1 clique
├── requirements.txt        # Dependências do projeto versionadas
└── data/
    └── tarifas_consolidated.json  # Cache offline da base de tarifas homologadas
```

---

## 🚀 Como Executar

### Opção 1: Atalho em 1 Clique (Windows)
Basta dar um duplo clique no arquivo [`Iniciar_Aplicacao.bat`](file:///C:/Users/user/.gemini/antigravity-ide/scratch/auditor-faturas-peakshaving/Iniciar_Aplicacao.bat). A aplicação inicializará no navegador em:
`http://localhost:8501`

### Opção 2: Pelo Terminal / Prompt de Comando
```bash
cd "C:\Users\user\.gemini\antigravity-ide\scratch\auditor-faturas-peakshaving"
.\.venv\Scripts\streamlit.exe run app.py
```
