"""
tariff_engine.py
----------------
Motor de cálculo, auditoria tarifária ANEEL e simulação técnico-financeira de Peak Shaving.
Suporta distinção rigorosa entre:
- Mercado Cativo (ACR): compra energia (TE) e uso da rede (TUSD) da distribuidora com bandeiras.
- Mercado Livre (ACL): fatura da distribuidora contém apenas TUSD (sem TE e sem bandeiras); energia faturada via contrato bilateral livre.

Executa a comparação rigorosa entre:
- Cenário 0: Fatura Atual (Auditoria de Conformidade e Reconstituição)
- Cenário 1: Tarifa Horária Verde + Operação de Gerador a Diesel
- Cenário 2: Tarifa Horária Azul + Operação de Gerador a Diesel (com otimização de Demanda Ponta)
"""

from __future__ import annotations

import logging
from typing import Any, Dict, List, Optional, Tuple

from pydantic import BaseModel, Field

from invoice_parser import InvoiceData
from tariff_client import (
    ConcessionariaTarifas,
    TarifaModalidadeAzul,
    TarifaModalidadeVerde,
    tariff_client,
)

logger = logging.getLogger(__name__)


class ParametrosGerador(BaseModel):
    preco_diesel_litro: float = Field(default=6.20, description="Preço do Óleo Diesel em R$/litro")
    consumo_especifico_l_kwh: float = Field(default=0.28, description="Consumo específico em L/kWh (típico: 0.27 a 0.32)")
    custo_manutencao_kwh: float = Field(default=0.15, description="Custo de O&M / manutenção em R$/kWh")
    potencia_nominal_kva: float = Field(default=500.0, description="Potência nominal do skid em kVA")
    potencia_continua_kw: float = Field(default=400.0, description="Potência ativa contínua disponível em kW")
    dias_uteis_mes: int = Field(default=22, description="Dias de despacho no mês (dias úteis)")
    horas_ponta_dia: float = Field(default=3.0, description="Horas de ponta por dia de despacho (padrão ANEEL: 3h)")
    fator_potencia_gerador: float = Field(default=0.80, description="Fator de potência nominal do gerador")
    capex_instalacao_reais: float = Field(default=0.0, description="Investimento inicial opcional para cálculo de Payback (R$)")

    @property
    def custo_gerador_unitario(self) -> float:
        """Fórmula do Custo Marginal de Geração (OPEX Unitário): (Preço_Diesel * Consumo_esp) + O&M"""
        return (self.preco_diesel_litro * self.consumo_especifico_l_kwh) + self.custo_manutencao_kwh

    @property
    def horas_operacao_mes(self) -> float:
        """Horas totais de despacho no mês."""
        return self.dias_uteis_mes * self.horas_ponta_dia

    @property
    def capacidade_maxima_geracao_mes_kwh(self) -> float:
        """Energia máxima que o gerador pode fornecer no horário de ponta do mês."""
        return self.potencia_continua_kw * self.horas_operacao_mes


class DetalheCenario(BaseModel):
    cenario_nome: str
    modalidade: str
    ambiente_contratacao: str = "CATIVO"  # "CATIVO" ou "LIVRE"
    operacao_gerador: bool

    # Componentes de Demanda
    demanda_contratada_fora_ponta_kw: float
    demanda_contratada_ponta_kw: Optional[float] = None
    tarifa_demanda_fp_homologada: float
    tarifa_demanda_p_homologada: Optional[float] = None
    custo_demanda_fp_reais: float
    custo_demanda_p_reais: float = 0.0
    custo_demanda_total_reais: float

    # Componentes de Consumo da Rede (TUSD e TE)
    consumo_fora_ponta_kwh: float
    tarifa_tusd_fp_homologada: float
    tarifa_te_fp_homologada: float
    custo_tusd_fp_reais: float
    custo_te_fp_reais: float = 0.0
    custo_consumo_fp_total_reais: float

    consumo_ponta_total_kwh: float
    consumo_ponta_gerador_kwh: float = 0.0
    consumo_ponta_rede_kwh: float = 0.0
    tarifa_tusd_p_homologada: float
    tarifa_te_p_homologada: float
    custo_tusd_ponta_reais: float = 0.0
    custo_te_ponta_reais: float = 0.0
    custo_consumo_ponta_rede_reais: float = 0.0

    # Custo de Energia no Mercado Livre (ACL)
    preco_energia_livre_kwh: float = 0.0
    custo_energia_livre_reais: float = 0.0

    # Custo do Gerador
    custo_unitario_gerador_reais_kwh: float = 0.0
    gasto_gerador_diesel_reais: float = 0.0
    gasto_gerador_manutencao_reais: float = 0.0
    gasto_gerador_total_reais: float = 0.0

    # Bandeiras e Tributos Detalhados
    custo_bandeiras_reais: float = 0.0
    base_calculo_tributos_reais: float = 0.0
    aliquota_icms_percent: float = 18.0
    valor_icms_reais: float = 0.0
    aliquota_pis_percent: float = 1.65
    valor_pis_reais: float = 0.0
    aliquota_cofins_percent: float = 7.60
    valor_cofins_reais: float = 0.0
    aliquota_tributos_total_percent: float = 27.25
    fator_grossup_aplicado: float = 1.0
    valor_cip_reais: float = 0.0

    # Totais
    custo_fatura_concessionaria_reais: float
    custo_total_operacional_reais: float  # Fatura Distribuidora + Contrato ACL + Gerador

    # Métricas de apoio
    custo_medio_rede_ponta_kwh: float  # Tarifa unitária ponta efetiva evitada


class ResultadoSimulacao(BaseModel):
    distribuidora: str
    resolucao_homologatoria: str
    data_vigencia: Optional[str]
    subgrupo: str = "A4"
    ambiente_contratacao: str = "CATIVO"

    cenario_0_atual: DetalheCenario
    cenario_1_verde: DetalheCenario
    cenario_2_azul: DetalheCenario

    # Auditoria de Conformidade da Fatura
    fatura_auditada_conforme: bool
    valor_fatura_reconstituida_reais: float
    valor_fatura_cobrado_real: float
    desvio_auditoria_reais: float
    desvio_auditoria_percentual: float
    diagnostico_auditoria: str

    # Análise de Viabilidade
    viavel_tecnicamente: bool
    viavel_economicamente: bool
    modalidade_recomendada: str  # "VERDE" ou "AZUL"
    cenario_recomendado_nome: str

    custo_total_atual_reais: float
    custo_total_recomendado_reais: float
    economia_mensal_liquida_reais: float
    economia_mensal_percentual: float
    economia_anual_projetada_reais: float
    payback_meses: Optional[float] = None

    # Recomendações Técnicas
    recomendacao_demanda_ponta_kw: Optional[float] = None
    recomendacao_demanda_fp_kw: float
    justificativa_tecnica: str
    comparativo_custo_ponta: Dict[str, float]

    # Tabelas Transparentes para Exibição e Excel
    tabela_tarifas_utilizadas: List[Dict[str, Any]] = Field(default_factory=list)
    tabela_tributos_utilizados: List[Dict[str, Any]] = Field(default_factory=list)


class TariffEngine:
    """Motor central de cálculo para tarifas do Grupo A e Peak Shaving (Cativo e Livre)."""

    def __init__(self):
        pass

    def simular(
        self,
        fatura: InvoiceData,
        gerador: ParametrosGerador,
        concessionaria_nome: Optional[str] = None,
        aliquota_impostos_override: Optional[float] = None,
        ambiente_contratacao_override: Optional[str] = None,
        preco_energia_acl_override: Optional[float] = None,
    ) -> ResultadoSimulacao:
        """
        Executa a reconstituição do Cenário 0, auditoria da cobrança e simulação
        dos Cenários 1 (Verde) e 2 (Azul) considerando Mercado Cativo (ACR) ou Livre (ACL).
        """
        # 1. Obtenção das Tarifas Homologadas da Concessionária
        target_dist = (
            concessionaria_nome
            or fatura.concessionaria_homologada_aneel
            or fatura.distribuidora_identificada
        )
        tarifas_homologadas = tariff_client.get_tarifas(target_dist, subgrupo=fatura.subgrupo)

        # 2. Ambiente de Contratação e Preço Livre
        ambiente = (
            ambiente_contratacao_override
            or fatura.ambiente_contratacao
            or "CATIVO"
        ).upper()
        is_livre = (ambiente == "LIVRE")
        preco_acl = (
            preco_energia_acl_override
            if preco_energia_acl_override is not None
            else fatura.preco_energia_acl_kwh
        )

        # 3. Definição da Alíquota de Tributos (ICMS + PIS + COFINS)
        aliquota_icms = fatura.tributos.icms_aliquota if fatura.tributos.icms_aliquota > 0 else 18.0
        aliquota_pis = fatura.tributos.pis_aliquota if fatura.tributos.pis_aliquota > 0 else 1.65
        aliquota_cofins = fatura.tributos.cofins_aliquota if fatura.tributos.cofins_aliquota > 0 else 7.60
        soma_aliquotas_nominais = aliquota_icms + aliquota_pis + aliquota_cofins

        if aliquota_impostos_override is not None:
            fator_imposto = 1.0 + (aliquota_impostos_override / 100.0)
            aliquota_total_efetiva = aliquota_impostos_override
        else:
            # Cálculo oficial por dentro (LC 87/96)
            if soma_aliquotas_nominais < 99.0:
                fator_imposto = 1.0 / (1.0 - (soma_aliquotas_nominais / 100.0))
                aliquota_total_efetiva = (fator_imposto - 1.0) * 100.0
            else:
                fator_imposto = 1.0 + (soma_aliquotas_nominais / 100.0)
                aliquota_total_efetiva = soma_aliquotas_nominais

        # 4. Parametrização do Skid de Geração no Horário de Ponta
        demanda_ponta_planta = fatura.get_effective_demanda_medida_ponta()
        consumo_ponta_total = fatura.consumo_ponta_kwh
        consumo_fp_total = fatura.consumo_fora_ponta_kwh

        capacidade_max_gerador = gerador.capacidade_maxima_geracao_mes_kwh
        consumo_gerador_ponta = min(consumo_ponta_total, capacidade_max_gerador)
        consumo_rede_ponta_residual = max(0.0, consumo_ponta_total - consumo_gerador_ponta)

        custo_unit_gerador = gerador.custo_gerador_unitario
        gasto_diesel = (
            consumo_gerador_ponta
            * gerador.consumo_especifico_l_kwh
            * gerador.preco_diesel_litro
        )
        gasto_manut = consumo_gerador_ponta * gerador.custo_manutencao_kwh
        gasto_gerador_mensal = gasto_diesel + gasto_manut

        # ---------------------------------------------------------------------
        # Helper: Cálculo dos Tributos e Fatura para um Cenario
        # ---------------------------------------------------------------------
        def _calcular_tributos_cenario(subtotal_sem_trib: float, cip: float) -> Tuple[float, float, float, float]:
            """Retorna (base_calculo, valor_icms, valor_pis, valor_cofins)"""
            base = subtotal_sem_trib * fator_imposto
            v_icms = base * (aliquota_icms / 100.0)
            v_pis = base * (aliquota_pis / 100.0)
            v_cofins = base * (aliquota_cofins / 100.0)
            return round(base, 2), round(v_icms, 2), round(v_pis, 2), round(v_cofins, 2)

        # ---------------------------------------------------------------------
        # CENÁRIO 0: Fatura Atual Reconstituída (Auditoria)
        # ---------------------------------------------------------------------
        mod_atual = fatura.modalidade.upper()
        if mod_atual == "AZUL":
            t_azul = tarifas_homologadas.azul
            dem_fp_0 = fatura.demanda_contratada_kw
            dem_p_0 = fatura.get_effective_demanda_contratada_ponta()

            dem_fp_cost_raw = dem_fp_0 * t_azul.demanda_fora_ponta
            dem_p_cost_raw = dem_p_0 * t_azul.demanda_ponta
            dem_tot_cost_raw = dem_fp_cost_raw + dem_p_cost_raw

            tusd_fp_cost_raw = consumo_fp_total * t_azul.tusd_fora_ponta
            tusd_p_cost_raw = consumo_ponta_total * t_azul.tusd_ponta

            # Se Livre, TE da distribuidora é ZERO e Bandeiras na distribuidora é ZERO
            if is_livre:
                te_fp_cost_raw = 0.0
                te_p_cost_raw = 0.0
                bandeira_fatura_0 = 0.0
                custo_energia_livre_0 = (consumo_fp_total + consumo_ponta_total) * preco_acl
            else:
                te_fp_cost_raw = consumo_fp_total * t_azul.te_fora_ponta
                te_p_cost_raw = consumo_ponta_total * t_azul.te_ponta
                bandeira_fatura_0 = fatura.bandeira_adicional_reais
                custo_energia_livre_0 = 0.0

            subtotal_rede_sem_trib_0 = (
                dem_tot_cost_raw + tusd_fp_cost_raw + tusd_p_cost_raw + te_fp_cost_raw + te_p_cost_raw + bandeira_fatura_0
            )
            base_trib_0, v_icms_0, v_pis_0, v_cofins_0 = _calcular_tributos_cenario(
                subtotal_rede_sem_trib_0, fatura.tributos.cip_cosip
            )
            fatura_distribuidora_reconst_0 = (subtotal_rede_sem_trib_0 * fator_imposto) + fatura.tributos.cip_cosip

            # Custo Total = Fatura da Distribuidora + Contrato ACL
            custo_total_operacional_0 = fatura_distribuidora_reconst_0 + custo_energia_livre_0

            cenario_0 = DetalheCenario(
                cenario_nome="Cenário 0 (Fatura Atual - Tarifa Azul)",
                modalidade="AZUL",
                ambiente_contratacao=ambiente,
                operacao_gerador=False,
                demanda_contratada_fora_ponta_kw=dem_fp_0,
                demanda_contratada_ponta_kw=dem_p_0,
                tarifa_demanda_fp_homologada=t_azul.demanda_fora_ponta,
                tarifa_demanda_p_homologada=t_azul.demanda_ponta,
                custo_demanda_fp_reais=round(dem_fp_cost_raw * fator_imposto, 2),
                custo_demanda_p_reais=round(dem_p_cost_raw * fator_imposto, 2),
                custo_demanda_total_reais=round(dem_tot_cost_raw * fator_imposto, 2),
                consumo_fora_ponta_kwh=consumo_fp_total,
                tarifa_tusd_fp_homologada=t_azul.tusd_fora_ponta,
                tarifa_te_fp_homologada=0.0 if is_livre else t_azul.te_fora_ponta,
                custo_tusd_fp_reais=round(tusd_fp_cost_raw * fator_imposto, 2),
                custo_te_fp_reais=round(te_fp_cost_raw * fator_imposto, 2),
                custo_consumo_fp_total_reais=round((tusd_fp_cost_raw + te_fp_cost_raw) * fator_imposto, 2),
                consumo_ponta_total_kwh=consumo_ponta_total,
                consumo_ponta_gerador_kwh=0.0,
                consumo_ponta_rede_kwh=consumo_ponta_total,
                tarifa_tusd_p_homologada=t_azul.tusd_ponta,
                tarifa_te_p_homologada=0.0 if is_livre else t_azul.te_ponta,
                custo_tusd_ponta_reais=round(tusd_p_cost_raw * fator_imposto, 2),
                custo_te_ponta_reais=round(te_p_cost_raw * fator_imposto, 2),
                custo_consumo_ponta_rede_reais=round((tusd_p_cost_raw + te_p_cost_raw) * fator_imposto, 2),
                preco_energia_livre_kwh=preco_acl if is_livre else 0.0,
                custo_energia_livre_reais=round(custo_energia_livre_0, 2),
                custo_unitario_gerador_reais_kwh=custo_unit_gerador,
                custo_bandeiras_reais=round(bandeira_fatura_0, 2),
                base_calculo_tributos_reais=base_trib_0,
                aliquota_icms_percent=aliquota_icms,
                valor_icms_reais=v_icms_0,
                aliquota_pis_percent=aliquota_pis,
                valor_pis_reais=v_pis_0,
                aliquota_cofins_percent=aliquota_cofins,
                valor_cofins_reais=v_cofins_0,
                aliquota_tributos_total_percent=round(aliquota_total_efetiva, 2),
                fator_grossup_aplicado=round(fator_imposto, 4),
                valor_cip_reais=fatura.tributos.cip_cosip,
                custo_fatura_concessionaria_reais=round(fatura_distribuidora_reconst_0, 2),
                custo_total_operacional_reais=round(custo_total_operacional_0, 2),
                custo_medio_rede_ponta_kwh=round(
                    (t_azul.tusd_ponta if is_livre else t_azul.total_ponta) * fator_imposto, 4
                ),
            )
        else:
            # Verde Atual
            t_verde = tarifas_homologadas.verde
            dem_unica_0 = fatura.demanda_contratada_kw

            dem_tot_cost_raw = dem_unica_0 * t_verde.demanda_unica
            tusd_fp_cost_raw = consumo_fp_total * t_verde.tusd_fora_ponta
            tusd_p_cost_raw = consumo_ponta_total * t_verde.tusd_ponta

            if is_livre:
                te_fp_cost_raw = 0.0
                te_p_cost_raw = 0.0
                bandeira_fatura_0 = 0.0
                custo_energia_livre_0 = (consumo_fp_total + consumo_ponta_total) * preco_acl
            else:
                te_fp_cost_raw = consumo_fp_total * t_verde.te_fora_ponta
                te_p_cost_raw = consumo_ponta_total * t_verde.te_ponta
                bandeira_fatura_0 = fatura.bandeira_adicional_reais
                custo_energia_livre_0 = 0.0

            subtotal_rede_sem_trib_0 = (
                dem_tot_cost_raw + tusd_fp_cost_raw + tusd_p_cost_raw + te_fp_cost_raw + te_p_cost_raw + bandeira_fatura_0
            )
            base_trib_0, v_icms_0, v_pis_0, v_cofins_0 = _calcular_tributos_cenario(
                subtotal_rede_sem_trib_0, fatura.tributos.cip_cosip
            )
            fatura_distribuidora_reconst_0 = (subtotal_rede_sem_trib_0 * fator_imposto) + fatura.tributos.cip_cosip
            custo_total_operacional_0 = fatura_distribuidora_reconst_0 + custo_energia_livre_0

            cenario_0 = DetalheCenario(
                cenario_nome="Cenário 0 (Fatura Atual - Tarifa Verde)",
                modalidade="VERDE",
                ambiente_contratacao=ambiente,
                operacao_gerador=False,
                demanda_contratada_fora_ponta_kw=dem_unica_0,
                demanda_contratada_ponta_kw=None,
                tarifa_demanda_fp_homologada=t_verde.demanda_unica,
                tarifa_demanda_p_homologada=None,
                custo_demanda_fp_reais=round(dem_tot_cost_raw * fator_imposto, 2),
                custo_demanda_p_reais=0.0,
                custo_demanda_total_reais=round(dem_tot_cost_raw * fator_imposto, 2),
                consumo_fora_ponta_kwh=consumo_fp_total,
                tarifa_tusd_fp_homologada=t_verde.tusd_fora_ponta,
                tarifa_te_fp_homologada=0.0 if is_livre else t_verde.te_fora_ponta,
                custo_tusd_fp_reais=round(tusd_fp_cost_raw * fator_imposto, 2),
                custo_te_fp_reais=round(te_fp_cost_raw * fator_imposto, 2),
                custo_consumo_fp_total_reais=round((tusd_fp_cost_raw + te_fp_cost_raw) * fator_imposto, 2),
                consumo_ponta_total_kwh=consumo_ponta_total,
                consumo_ponta_gerador_kwh=0.0,
                consumo_ponta_rede_kwh=consumo_ponta_total,
                tarifa_tusd_p_homologada=t_verde.tusd_ponta,
                tarifa_te_p_homologada=0.0 if is_livre else t_verde.te_ponta,
                custo_tusd_ponta_reais=round(tusd_p_cost_raw * fator_imposto, 2),
                custo_te_ponta_reais=round(te_p_cost_raw * fator_imposto, 2),
                custo_consumo_ponta_rede_reais=round((tusd_p_cost_raw + te_p_cost_raw) * fator_imposto, 2),
                preco_energia_livre_kwh=preco_acl if is_livre else 0.0,
                custo_energia_livre_reais=round(custo_energia_livre_0, 2),
                custo_unitario_gerador_reais_kwh=custo_unit_gerador,
                custo_bandeiras_reais=round(bandeira_fatura_0, 2),
                base_calculo_tributos_reais=base_trib_0,
                aliquota_icms_percent=aliquota_icms,
                valor_icms_reais=v_icms_0,
                aliquota_pis_percent=aliquota_pis,
                valor_pis_reais=v_pis_0,
                aliquota_cofins_percent=aliquota_cofins,
                valor_cofins_reais=v_cofins_0,
                aliquota_tributos_total_percent=round(aliquota_total_efetiva, 2),
                fator_grossup_aplicado=round(fator_imposto, 4),
                valor_cip_reais=fatura.tributos.cip_cosip,
                custo_fatura_concessionaria_reais=round(fatura_distribuidora_reconst_0, 2),
                custo_total_operacional_reais=round(custo_total_operacional_0, 2),
                custo_medio_rede_ponta_kwh=round(
                    (t_verde.tusd_ponta if is_livre else t_verde.total_ponta) * fator_imposto, 4
                ),
            )

        # ---------------------------------------------------------------------
        # AUDITORIA DA FATURA COBRADA vs. RECONSTITUÍDA
        # ---------------------------------------------------------------------
        valor_fatura_cobrado = fatura.valor_total_faturado
        valor_reconstituido = cenario_0.custo_fatura_concessionaria_reais
        desvio_reais = valor_fatura_cobrado - valor_reconstituido
        desvio_pct = (desvio_reais / valor_reconstituido * 100.0) if valor_reconstituido > 0 else 0.0

        if abs(desvio_pct) <= 2.5:
            conforme = True
            diag = (
                f"✅ Fatura Conforme: O valor faturado pela distribuidora (R$ {valor_fatura_cobrado:,.2f}) "
                f"está rigorosamente alinhado à tarifa homologada ANEEL ({resultado_reh := tarifas_homologadas.resolucao_homologatoria[:45]}...) "
                f"com variação mínima de {desvio_pct:+.2f}% (R$ {desvio_reais:,.2f}), decorrente de arredondamentos legais."
            )
        elif desvio_reais > 0:
            conforme = False
            diag = (
                f"⚠️ Divergência Detectada (Possível Cobrança a Maior): O valor cobrado na fatura (R$ {valor_fatura_cobrado:,.2f}) "
                f"está R$ {desvio_reais:,.2f} ({desvio_pct:+.1f}%) ACIMA da tarifa homologada pela ANEEL e alíquotas de tributos. "
                f"Recomenda-se auditar: cobrança de ultrapassagem de demanda, encargos de reativo excedente (UFER/DMCR) ou aplicação indevida de alíquotas."
            )
        else:
            conforme = False
            diag = (
                f"ℹ️ Divergência Detectada (Cobrança a Menor / Ciclo Parcial): O valor cobrado (R$ {valor_fatura_cobrado:,.2f}) "
                f"está R$ {abs(desvio_reais):,.2f} ({abs(desvio_pct):.1f}%) abaixo da tarifa cheia homologada. "
                f"Possíveis causas: ciclo de medição inferior a 30 dias, créditos tributários/compensação de GD ou diferimento tarifário."
            )

        # ---------------------------------------------------------------------
        # CENÁRIO 1: Tarifa Verde + Operação de Gerador
        # ---------------------------------------------------------------------
        t_verde = tarifas_homologadas.verde
        dem_verde_1 = fatura.demanda_contratada_kw
        dem_cost_verde_1_raw = dem_verde_1 * t_verde.demanda_unica

        tusd_fp_cost_1_raw = consumo_fp_total * t_verde.tusd_fora_ponta
        tusd_p_cost_1_raw = consumo_rede_ponta_residual * t_verde.tusd_ponta

        if is_livre:
            te_fp_cost_1_raw = 0.0
            te_p_cost_1_raw = 0.0
            bandeira_cost_1 = 0.0
            # No ACL, a energia consumida da rede reduz pelo consumo do gerador
            custo_energia_livre_1 = (consumo_fp_total + consumo_rede_ponta_residual) * preco_acl
        else:
            te_fp_cost_1_raw = consumo_fp_total * t_verde.te_fora_ponta
            te_p_cost_1_raw = consumo_rede_ponta_residual * t_verde.te_ponta
            # Bandeira proporcional ao consumo residual da rede
            if fatura.bandeira_adicional_reais > 0 and (consumo_ponta_total + consumo_fp_total) > 0:
                t_band_unit = fatura.bandeira_adicional_reais / (consumo_ponta_total + consumo_fp_total)
                bandeira_cost_1 = (consumo_fp_total + consumo_rede_ponta_residual) * t_band_unit
            else:
                bandeira_cost_1 = 0.0
            custo_energia_livre_1 = 0.0

        subtotal_sem_trib_1 = (
            dem_cost_verde_1_raw + tusd_fp_cost_1_raw + tusd_p_cost_1_raw + te_fp_cost_1_raw + te_p_cost_1_raw + bandeira_cost_1
        )
        base_trib_1, v_icms_1, v_pis_1, v_cofins_1 = _calcular_tributos_cenario(
            subtotal_sem_trib_1, fatura.tributos.cip_cosip
        )
        fatura_distribuidora_1 = (subtotal_sem_trib_1 * fator_imposto) + fatura.tributos.cip_cosip
        custo_total_operacional_1 = fatatura_total_1 = fatura_distribuidora_1 + custo_energia_livre_1 + gasto_gerador_mensal

        cenario_1 = DetalheCenario(
            cenario_nome="Cenário 1 (Tarifa Verde + Peak Shaving)",
            modalidade="VERDE",
            ambiente_contratacao=ambiente,
            operacao_gerador=True,
            demanda_contratada_fora_ponta_kw=dem_verde_1,
            demanda_contratada_ponta_kw=None,
            tarifa_demanda_fp_homologada=t_verde.demanda_unica,
            tarifa_demanda_p_homologada=None,
            custo_demanda_fp_reais=round(dem_cost_verde_1_raw * fator_imposto, 2),
            custo_demanda_p_reais=0.0,
            custo_demanda_total_reais=round(dem_cost_verde_1_raw * fator_imposto, 2),
            consumo_fora_ponta_kwh=consumo_fp_total,
            tarifa_tusd_fp_homologada=t_verde.tusd_fora_ponta,
            tarifa_te_fp_homologada=0.0 if is_livre else t_verde.te_fora_ponta,
            custo_tusd_fp_reais=round(tusd_fp_cost_1_raw * fator_imposto, 2),
            custo_te_fp_reais=round(te_fp_cost_1_raw * fator_imposto, 2),
            custo_consumo_fp_total_reais=round((tusd_fp_cost_1_raw + te_fp_cost_1_raw) * fator_imposto, 2),
            consumo_ponta_total_kwh=consumo_ponta_total,
            consumo_ponta_gerador_kwh=consumo_gerador_ponta,
            consumo_ponta_rede_kwh=consumo_rede_ponta_residual,
            tarifa_tusd_p_homologada=t_verde.tusd_ponta,
            tarifa_te_p_homologada=0.0 if is_livre else t_verde.te_ponta,
            custo_tusd_ponta_reais=round(tusd_p_cost_1_raw * fator_imposto, 2),
            custo_te_ponta_reais=round(te_p_cost_1_raw * fator_imposto, 2),
            custo_consumo_ponta_rede_reais=round((tusd_p_cost_1_raw + te_p_cost_1_raw) * fator_imposto, 2),
            preco_energia_livre_kwh=preco_acl if is_livre else 0.0,
            custo_energia_livre_reais=round(custo_energia_livre_1, 2),
            custo_unitario_gerador_reais_kwh=custo_unit_gerador,
            gasto_gerador_diesel_reais=round(gasto_diesel, 2),
            gasto_gerador_manutencao_reais=round(gasto_manut, 2),
            gasto_gerador_total_reais=round(gasto_gerador_mensal, 2),
            custo_bandeiras_reais=round(bandeira_cost_1, 2),
            base_calculo_tributos_reais=base_trib_1,
            aliquota_icms_percent=aliquota_icms,
            valor_icms_reais=v_icms_1,
            aliquota_pis_percent=aliquota_pis,
            valor_pis_reais=v_pis_1,
            aliquota_cofins_percent=aliquota_cofins,
            valor_cofins_reais=v_cofins_1,
            aliquota_tributos_total_percent=round(aliquota_total_efetiva, 2),
            fator_grossup_aplicado=round(fator_imposto, 4),
            valor_cip_reais=fatura.tributos.cip_cosip,
            custo_fatura_concessionaria_reais=round(fatura_distribuidora_1, 2),
            custo_total_operacional_reais=round(custo_total_operacional_1, 2),
            custo_medio_rede_ponta_kwh=round(
                (t_verde.tusd_ponta if is_livre else t_verde.total_ponta) * fator_imposto, 4
            ),
        )

        # ---------------------------------------------------------------------
        # CENÁRIO 2: Tarifa Azul + Operação de Gerador (Otimizado)
        # ---------------------------------------------------------------------
        t_azul = tarifas_homologadas.azul
        dem_azul_fp = fatura.demanda_contratada_kw

        # No Azul, com o gerador absorvendo a ponta, rebaixa-se a Demanda Ponta ao mínimo regulatório
        demanda_ponta_residual = max(0.0, demanda_ponta_planta - gerador.potencia_continua_kw)
        dem_azul_p_otimizada = max(30.0, round(demanda_ponta_residual, 1))

        dem_fp_cost_2_raw = dem_azul_fp * t_azul.demanda_fora_ponta
        dem_p_cost_2_raw = dem_azul_p_otimizada * t_azul.demanda_ponta
        dem_tot_cost_2_raw = dem_fp_cost_2_raw + dem_p_cost_2_raw

        tusd_fp_cost_2_raw = consumo_fp_total * t_azul.tusd_fora_ponta
        tusd_p_cost_2_raw = consumo_rede_ponta_residual * t_azul.tusd_ponta

        if is_livre:
            te_fp_cost_2_raw = 0.0
            te_p_cost_2_raw = 0.0
            bandeira_cost_2 = 0.0
            custo_energia_livre_2 = (consumo_fp_total + consumo_rede_ponta_residual) * preco_acl
        else:
            te_fp_cost_2_raw = consumo_fp_total * t_azul.te_fora_ponta
            te_p_cost_2_raw = consumo_rede_ponta_residual * t_azul.te_ponta
            if fatura.bandeira_adicional_reais > 0 and (consumo_ponta_total + consumo_fp_total) > 0:
                t_band_unit = fatura.bandeira_adicional_reais / (consumo_ponta_total + consumo_fp_total)
                bandeira_cost_2 = (consumo_fp_total + consumo_rede_ponta_residual) * t_band_unit
            else:
                bandeira_cost_2 = 0.0
            custo_energia_livre_2 = 0.0

        subtotal_sem_trib_2 = (
            dem_tot_cost_2_raw + tusd_fp_cost_2_raw + tusd_p_cost_2_raw + te_fp_cost_2_raw + te_p_cost_2_raw + bandeira_cost_2
        )
        base_trib_2, v_icms_2, v_pis_2, v_cofins_2 = _calcular_tributos_cenario(
            subtotal_sem_trib_2, fatura.tributos.cip_cosip
        )
        fatura_distribuidora_2 = (subtotal_sem_trib_2 * fator_imposto) + fatura.tributos.cip_cosip
        custo_total_operacional_2 = fatura_distribuidora_2 + custo_energia_livre_2 + gasto_gerador_mensal

        cenario_2 = DetalheCenario(
            cenario_nome="Cenário 2 (Tarifa Azul + Peak Shaving Otimizado)",
            modalidade="AZUL",
            ambiente_contratacao=ambiente,
            operacao_gerador=True,
            demanda_contratada_fora_ponta_kw=dem_azul_fp,
            demanda_contratada_ponta_kw=dem_azul_p_otimizada,
            tarifa_demanda_fp_homologada=t_azul.demanda_fora_ponta,
            tarifa_demanda_p_homologada=t_azul.demanda_ponta,
            custo_demanda_fp_reais=round(dem_fp_cost_2_raw * fator_imposto, 2),
            custo_demanda_p_reais=round(dem_p_cost_2_raw * fator_imposto, 2),
            custo_demanda_total_reais=round(dem_tot_cost_2_raw * fator_imposto, 2),
            consumo_fora_ponta_kwh=consumo_fp_total,
            tarifa_tusd_fp_homologada=t_azul.tusd_fora_ponta,
            tarifa_te_fp_homologada=0.0 if is_livre else t_azul.te_fora_ponta,
            custo_tusd_fp_reais=round(tusd_fp_cost_2_raw * fator_imposto, 2),
            custo_te_fp_reais=round(te_fp_cost_2_raw * fator_imposto, 2),
            custo_consumo_fp_total_reais=round((tusd_fp_cost_2_raw + te_fp_cost_2_raw) * fator_imposto, 2),
            consumo_ponta_total_kwh=consumo_ponta_total,
            consumo_ponta_gerador_kwh=consumo_gerador_ponta,
            consumo_ponta_rede_kwh=consumo_rede_ponta_residual,
            tarifa_tusd_p_homologada=t_azul.tusd_ponta,
            tarifa_te_p_homologada=0.0 if is_livre else t_azul.te_ponta,
            custo_tusd_ponta_reais=round(tusd_p_cost_2_raw * fator_imposto, 2),
            custo_te_ponta_reais=round(te_p_cost_2_raw * fator_imposto, 2),
            custo_consumo_ponta_rede_reais=round((tusd_p_cost_2_raw + te_p_cost_2_raw) * fator_imposto, 2),
            preco_energia_livre_kwh=preco_acl if is_livre else 0.0,
            custo_energia_livre_reais=round(custo_energia_livre_2, 2),
            custo_unitario_gerador_reais_kwh=custo_unit_gerador,
            gasto_gerador_diesel_reais=round(gasto_diesel, 2),
            gasto_gerador_manutencao_reais=round(gasto_manut, 2),
            gasto_gerador_total_reais=round(gasto_gerador_mensal, 2),
            custo_bandeiras_reais=round(bandeira_cost_2, 2),
            base_calculo_tributos_reais=base_trib_2,
            aliquota_icms_percent=aliquota_icms,
            valor_icms_reais=v_icms_2,
            aliquota_pis_percent=aliquota_pis,
            valor_pis_reais=v_pis_2,
            aliquota_cofins_percent=aliquota_cofins,
            valor_cofins_reais=v_cofins_2,
            aliquota_tributos_total_percent=round(aliquota_total_efetiva, 2),
            fator_grossup_aplicado=round(fator_imposto, 4),
            valor_cip_reais=fatura.tributos.cip_cosip,
            custo_fatura_concessionaria_reais=round(fatura_distribuidora_2, 2),
            custo_total_operacional_reais=round(custo_total_operacional_2, 2),
            custo_medio_rede_ponta_kwh=round(
                (t_azul.tusd_ponta if is_livre else t_azul.total_ponta) * fator_imposto, 4
            ),
        )

        # ---------------------------------------------------------------------
        # DECISÃO E VIABILIDADE TÉCNICA E ECONÔMICA
        # ---------------------------------------------------------------------
        tarifa_rede_ponta_verde_efetiva = (
            (t_verde.tusd_ponta if is_livre else t_verde.total_ponta) * fator_imposto
        )
        viavel_tecnico = gerador.potencia_continua_kw > 0 and consumo_gerador_ponta > 0

        # No Cativo: compara com TUSD_P + TE_P. No Livre: compara com TUSD_P
        viavel_economico = custo_unit_gerador < tarifa_rede_ponta_verde_efetiva

        # Comparação financeira entre Cenário 1 (Verde) e Cenário 2 (Azul)
        if custo_total_operacional_1 <= custo_total_operacional_2:
            melhor_cenario = cenario_1
            modalidade_rec = "VERDE"
            rec_nome = "Cenário 1 (Tarifa Verde + Peak Shaving)"
            rec_demanda_ponta = None
            rec_demanda_fp = dem_verde_1
        else:
            melhor_cenario = cenario_2
            modalidade_rec = "AZUL"
            rec_nome = "Cenário 2 (Tarifa Azul + Peak Shaving Otimizado)"
            rec_demanda_ponta = dem_azul_p_otimizada
            rec_demanda_fp = dem_azul_fp

        custo_atual_base = cenario_0.custo_total_operacional_reais
        custo_otimo = melhor_cenario.custo_total_operacional_reais
        economia_mensal = custo_atual_base - custo_otimo
        economia_pct = (economia_mensal / custo_atual_base * 100.0) if custo_atual_base > 0 else 0.0
        economia_anual = economia_mensal * 12.0

        payback = None
        if gerador.capex_instalacao_reais > 0 and economia_mensal > 0:
            payback = round(gerador.capex_instalacao_reais / economia_mensal, 1)

        diff_unit = tarifa_rede_ponta_verde_efetiva - custo_unit_gerador
        if modalidade_rec == "VERDE":
            justificativa = (
                f"No ambiente {ambiente}, a Modalidade VERDE com Peak Shaving apresentou a maior economia líquida "
                f"(R$ {economia_mensal:,.2f}/mês, {economia_pct:.1f}% de redução). "
                f"O custo do gerador (R$ {custo_unit_gerador:.3f}/kWh) é inferior ao custo de ponta evitado na rede "
                f"(R$ {tarifa_rede_ponta_verde_efetiva:.3f}/kWh c/ tributos). "
                f"A contratação de Demanda Única dimensionada para a carga fora de ponta ({rec_demanda_fp:.1f} kW) "
                f"evita os pesados encargos fixos da Demanda Ponta da modalidade Azul."
            )
        else:
            justificativa = (
                f"No ambiente {ambiente}, a Modalidade AZUL com Demanda Ponta Otimizada apresentou a maior economia líquida "
                f"(R$ {economia_mensal:,.2f}/mês, {economia_pct:.1f}% de redução). "
                f"Como o skid térmico absorve a ponta, foi possível rebaixar a Demanda Contratada de Ponta para "
                f"{rec_demanda_ponta:.1f} kW (mínimo técnico regulatório ANEEL), eliminando a cobrança fixa da tarifa "
                f"de Demanda Ponta da distribuidora (R$ {t_azul.demanda_ponta:.2f}/kW)."
            )

        comparativo_ponta = {
            "custo_gerador_unitario": round(custo_unit_gerador, 4),
            "tarifa_rede_verde_ponta_com_impostos": round(tarifa_rede_ponta_verde_efetiva, 4),
            "tarifa_rede_verde_ponta_sem_impostos": round(t_verde.tusd_ponta if is_livre else t_verde.total_ponta, 4),
            "tarifa_rede_azul_ponta_com_impostos": round((t_azul.tusd_ponta if is_livre else t_azul.total_ponta) * fator_imposto, 4),
            "margem_unit_diesel_vs_rede_verde": round(diff_unit, 4),
        }

        # ---------------------------------------------------------------------
        # TABELAS DE DETALHAMENTO DE TODAS AS TARIFAS E TRIBUTOS UTILIZADOS
        # ---------------------------------------------------------------------
        tabela_tarifas = [
            {
                "Componente": "Demanda Fora Ponta / Única",
                "Unidade": "R$/kW",
                "Modalidade": "Verde",
                "Ambiente": ambiente,
                "Tarifa_Homologada_Sem_Tributos": t_verde.demanda_unica,
                "Tarifa_Efetiva_Com_Tributos": round(t_verde.demanda_unica * fator_imposto, 4),
                "Onde_Faturado": "Distribuidora (Rede)",
            },
            {
                "Componente": "Demanda Fora Ponta",
                "Unidade": "R$/kW",
                "Modalidade": "Azul",
                "Ambiente": ambiente,
                "Tarifa_Homologada_Sem_Tributos": t_azul.demanda_fora_ponta,
                "Tarifa_Efetiva_Com_Tributos": round(t_azul.demanda_fora_ponta * fator_imposto, 4),
                "Onde_Faturado": "Distribuidora (Rede)",
            },
            {
                "Componente": "Demanda Ponta",
                "Unidade": "R$/kW",
                "Modalidade": "Azul",
                "Ambiente": ambiente,
                "Tarifa_Homologada_Sem_Tributos": t_azul.demanda_ponta,
                "Tarifa_Efetiva_Com_Tributos": round(t_azul.demanda_ponta * fator_imposto, 4),
                "Onde_Faturado": "Distribuidora (Rede)",
            },
            {
                "Componente": "TUSD Fora Ponta",
                "Unidade": "R$/kWh",
                "Modalidade": "Verde / Azul",
                "Ambiente": ambiente,
                "Tarifa_Homologada_Sem_Tributos": t_verde.tusd_fora_ponta,
                "Tarifa_Efetiva_Com_Tributos": round(t_verde.tusd_fora_ponta * fator_imposto, 5),
                "Onde_Faturado": "Distribuidora (Rede)",
            },
            {
                "Componente": "TUSD Ponta (Verde)",
                "Unidade": "R$/kWh",
                "Modalidade": "Verde",
                "Ambiente": ambiente,
                "Tarifa_Homologada_Sem_Tributos": t_verde.tusd_ponta,
                "Tarifa_Efetiva_Com_Tributos": round(t_verde.tusd_ponta * fator_imposto, 5),
                "Onde_Faturado": "Distribuidora (Rede)",
            },
            {
                "Componente": "TUSD Ponta (Azul)",
                "Unidade": "R$/kWh",
                "Modalidade": "Azul",
                "Ambiente": ambiente,
                "Tarifa_Homologada_Sem_Tributos": t_azul.tusd_ponta,
                "Tarifa_Efetiva_Com_Tributos": round(t_azul.tusd_ponta * fator_imposto, 5),
                "Onde_Faturado": "Distribuidora (Rede)",
            },
            {
                "Componente": "TE Fora Ponta (Concessionária)",
                "Unidade": "R$/kWh",
                "Modalidade": "Verde / Azul",
                "Ambiente": ambiente,
                "Tarifa_Homologada_Sem_Tributos": 0.0 if is_livre else t_verde.te_fora_ponta,
                "Tarifa_Efetiva_Com_Tributos": 0.0 if is_livre else round(t_verde.te_fora_ponta * fator_imposto, 5),
                "Onde_Faturado": "Isento na Distribuidora (ACL)" if is_livre else "Distribuidora (ACR)",
            },
            {
                "Componente": "TE Ponta (Concessionária)",
                "Unidade": "R$/kWh",
                "Modalidade": "Verde / Azul",
                "Ambiente": ambiente,
                "Tarifa_Homologada_Sem_Tributos": 0.0 if is_livre else t_verde.te_ponta,
                "Tarifa_Efetiva_Com_Tributos": 0.0 if is_livre else round(t_verde.te_ponta * fator_imposto, 5),
                "Onde_Faturado": "Isento na Distribuidora (ACL)" if is_livre else "Distribuidora (ACR)",
            },
            {
                "Componente": "Contrato Bilateral Livre (ACL)",
                "Unidade": "R$/kWh",
                "Modalidade": "Livre",
                "Ambiente": ambiente,
                "Tarifa_Homologada_Sem_Tributos": preco_acl if is_livre else 0.0,
                "Tarifa_Efetiva_Com_Tributos": preco_acl if is_livre else 0.0,
                "Onde_Faturado": "Comercializadora / Gerador Livre" if is_livre else "Não Aplicável (ACR)",
            },
            {
                "Componente": "Custo Marginal Skid Gerador Térmico",
                "Unidade": "R$/kWh",
                "Modalidade": "Geração Própria",
                "Ambiente": "Local",
                "Tarifa_Homologada_Sem_Tributos": round(custo_unit_gerador, 4),
                "Tarifa_Efetiva_Com_Tributos": round(custo_unit_gerador, 4),
                "Onde_Faturado": "Combustível (Diesel) + O&M",
            },
        ]

        tabela_tributos = [
            {
                "Tributo": "ICMS",
                "Base_Calculo_Reais": base_trib_0,
                "Aliquota_Nominal_Percent": aliquota_icms,
                "Metodologia": "Cálculo por Dentro (art. 13 LC 87/96)",
                "Valor_Apurado_Reais": v_icms_0,
            },
            {
                "Tributo": "PIS",
                "Base_Calculo_Reais": base_trib_0,
                "Aliquota_Nominal_Percent": aliquota_pis,
                "Metodologia": "Base Faturamento Concessionária",
                "Valor_Apurado_Reais": v_pis_0,
            },
            {
                "Tributo": "COFINS",
                "Base_Calculo_Reais": base_trib_0,
                "Aliquota_Nominal_Percent": aliquota_cofins,
                "Metodologia": "Base Faturamento Concessionária",
                "Valor_Apurado_Reais": v_cofins_0,
            },
            {
                "Tributo": "CIP / COSIP (Iluminação Pública)",
                "Base_Calculo_Reais": fatura.tributos.cip_cosip,
                "Aliquota_Nominal_Percent": 0.0,
                "Metodologia": "Tarifa Municipal Fixa",
                "Valor_Apurado_Reais": fatura.tributos.cip_cosip,
            },
            {
                "Tributo": "TOTAL TRIBUTOS COMBINADOS",
                "Base_Calculo_Reais": base_trib_0,
                "Aliquota_Nominal_Percent": round(soma_aliquotas_nominais, 2),
                "Metodologia": f"Gross-up Efetivo: +{aliquota_total_efetiva:.2f}% (Fator {fator_imposto:.4f})",
                "Valor_Apurado_Reais": round(v_icms_0 + v_pis_0 + v_cofins_0 + fatura.tributos.cip_cosip, 2),
            },
        ]

        return ResultadoSimulacao(
            distribuidora=tarifas_homologadas.agente,
            resolucao_homologatoria=tarifas_homologadas.resolucao_homologatoria,
            data_vigencia=tarifas_homologadas.data_vigencia,
            subgrupo=fatura.subgrupo,
            ambiente_contratacao=ambiente,
            cenario_0_atual=cenario_0,
            cenario_1_verde=cenario_1,
            cenario_2_azul=cenario_2,
            fatura_auditada_conforme=conforme,
            valor_fatura_reconstituida_reais=round(valor_reconstituido, 2),
            valor_fatura_cobrado_real=round(valor_fatura_cobrado, 2),
            desvio_auditoria_reais=round(desvio_reais, 2),
            desvio_auditoria_percentual=round(desvio_pct, 2),
            diagnostico_auditoria=diag,
            viavel_tecnicamente=viavel_tecnico,
            viavel_economicamente=viavel_economico,
            modalidade_recomendada=modalidade_rec,
            cenario_recomendado_nome=rec_nome,
            custo_total_atual_reais=round(custo_atual_base, 2),
            custo_total_recomendado_reais=round(custo_otimo, 2),
            economia_mensal_liquida_reais=round(economia_mensal, 2),
            economia_mensal_percentual=round(economia_pct, 2),
            economia_anual_projetada_reais=round(economia_anual, 2),
            payback_meses=payback,
            recomendacao_demanda_ponta_kw=rec_demanda_ponta,
            recomendacao_demanda_fp_kw=rec_demanda_fp,
            justificativa_tecnica=justificativa,
            comparativo_custo_ponta=comparativo_ponta,
            tabela_tarifas_utilizadas=tabela_tarifas,
            tabela_tributos_utilizados=tabela_tributos,
        )


tariff_engine = TariffEngine()
