"""
invoice_parser.py
-----------------
Módulo agnóstico de extração e validação de dados de faturas de energia elétrica
de consumidores do Grupo A (A4 / A3a) a partir de arquivos PDF.
Utiliza pdfplumber, expressões regulares heurísticas e modelos Pydantic
para estruturação dos dados.
"""

from __future__ import annotations

import io
import logging
import re
import unicodedata
from typing import Any, Dict, List, Optional, Tuple, Union

import pdfplumber
from pydantic import BaseModel, Field

from tariff_client import tariff_client

logger = logging.getLogger(__name__)


class FaturaTributos(BaseModel):
    icms_base: float = 0.0
    icms_aliquota: float = 0.0  # Ex: 18.0 para 18%
    icms_valor: float = 0.0
    pis_aliquota: float = 0.0   # Ex: 1.65 para 1.65%
    pis_valor: float = 0.0
    cofins_aliquota: float = 0.0 # Ex: 7.60 para 7.60%
    cofins_valor: float = 0.0
    cip_cosip: float = 0.0      # Iluminação pública R$

    @property
    def aliquota_efetiva_total(self) -> float:
        """Alíquota combinada aproximada (ICMS + PIS + COFINS)."""
        tot = self.icms_aliquota + self.pis_aliquota + self.cofins_aliquota
        if tot <= 0.0:
            return 25.0  # Padrão médio nacional razoável
        return round(tot, 2)


class FaturaTarifasPraticadas(BaseModel):
    tarifa_demanda_unica: Optional[float] = None
    tarifa_demanda_ponta: Optional[float] = None
    tarifa_demanda_fora_ponta: Optional[float] = None
    tusd_ponta: Optional[float] = None
    tusd_fora_ponta: Optional[float] = None
    te_ponta: Optional[float] = None
    te_fora_ponta: Optional[float] = None


class InvoiceData(BaseModel):
    distribuidora_identificada: str = Field(default="Distribuidora Não Identificada")
    concessionaria_homologada_aneel: Optional[str] = None
    subgrupo: str = Field(default="A4", description="Subgrupo tarifário (A4, A3a, etc.)")
    modalidade: str = Field(default="VERDE", description="Modalidade tarifária: VERDE ou AZUL")
    mes_referencia: str = Field(default="05/2026")
    numero_instalacao: Optional[str] = None

    # Ambiente de Contratação: CATIVO (ACR) vs. LIVRE (ACL)
    ambiente_contratacao: str = Field(default="CATIVO", description="Ambiente de contratação: 'CATIVO' (ACR) ou 'LIVRE' (ACL)")
    preco_energia_acl_kwh: float = Field(default=0.22, description="Preço do contrato bilateral livre no ACL em R$/kWh (se aplicável)")

    # Demandas (kW)
    demanda_contratada_kw: float = Field(default=300.0, description="Demanda única contratada (Verde) ou Fora Ponta (Azul)")
    demanda_contratada_ponta_kw: Optional[float] = Field(default=None, description="Demanda contratada na ponta (Azul)")
    demanda_medida_kw: float = Field(default=280.0, description="Demanda máxima medida fora ponta")
    demanda_medida_ponta_kw: Optional[float] = Field(default=None, description="Demanda medida na ponta")

    # Consumos (kWh)
    consumo_ponta_kwh: float = Field(default=15000.0, description="Consumo ativo faturado no horário de ponta")
    consumo_fora_ponta_kwh: float = Field(default=85000.0, description="Consumo ativo faturado fora de ponta")
    consumo_reativo_excedente_kwh: float = Field(default=0.0, description="UFER/DMCR reativo excedente")

    # Tarifas e Valores
    tarifas_praticadas: FaturaTarifasPraticadas = Field(default_factory=FaturaTarifasPraticadas)
    bandeira_tipo: str = Field(default="Verde", description="Tipo de bandeira tarifária vigente no ciclo")
    bandeira_adicional_reais: float = Field(default=0.0, description="Valor adicional em R$ cobrado por bandeiras")
    tributos: FaturaTributos = Field(default_factory=FaturaTributos)
    valor_total_faturado: float = Field(default=0.0, description="Valor total da fatura (R$)")

    # Metadados de extração
    confidence_score: float = Field(default=0.0, description="Grau de certeza na extração dos dados (0 a 1)")
    campos_detectados: List[str] = Field(default_factory=list)
    raw_text_preview: str = Field(default="")

    def get_effective_demanda_contratada_ponta(self) -> float:
        """Retorna demanda contratada ponta ou a única se Verde."""
        if self.modalidade.upper() == "AZUL" and self.demanda_contratada_ponta_kw is not None:
            return self.demanda_contratada_ponta_kw
        return self.demanda_contratada_kw

    def get_effective_demanda_medida_ponta(self) -> float:
        """Retorna demanda medida ponta ou a única se Verde."""
        if self.demanda_medida_ponta_kw is not None and self.demanda_medida_ponta_kw > 0:
            return self.demanda_medida_ponta_kw
        return self.demanda_medida_kw


def _clean_str(text: str) -> str:
    if not text:
        return ""
    return unicodedata.normalize("NFKD", text).encode("ASCII", "ignore").decode("utf-8").upper()


def _parse_float_ptbr(val_str: str) -> Optional[float]:
    """Converte formatos numéricos brasileiros (ex: '1.250,50' ou '0,45892') para float."""
    if not val_str:
        return None
    s = val_str.strip().replace(" ", "").replace("R$", "").replace("$", "")
    # Remove separadores de milhar e converte vírgula decimal
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


class InvoiceParser:
    """Parser agnóstico e resiliente para faturas de energia do Grupo A."""

    def __init__(self):
        pass

    def extract_text_from_pdf(self, pdf_file: Union[str, bytes, io.BytesIO]) -> str:
        """Extrai todo o texto contido nas páginas do PDF com tratamento de quebras."""
        full_text = []
        try:
            with pdfplumber.open(pdf_file) as pdf:
                for idx, page in enumerate(pdf.pages):
                    text = page.extract_text(layout=True) or page.extract_text() or ""
                    full_text.append(text)
        except Exception as e:
            logger.error("Erro ao ler PDF com pdfplumber: %s", e)
            raise ValueError(f"Não foi possível processar o arquivo PDF fornecido: {e}")

        return "\n--- QUEBRA DE PAGINA ---\n".join(full_text)

    def parse_pdf(self, pdf_file: Union[str, bytes, io.BytesIO]) -> InvoiceData:
        """Processa a fatura em PDF e retorna um objeto InvoiceData estruturado e validado."""
        raw_text = self.extract_text_from_pdf(pdf_file)
        return self.parse_text(raw_text)

    def parse_text(self, text: str) -> InvoiceData:
        """Executa heurísticas e expressões regulares sobre o texto da fatura."""
        clean_text = _clean_str(text)
        lines = [line.strip() for line in clean_text.splitlines() if line.strip()]
        campos: List[str] = []

        # 1. Identificação da Distribuidora / Concessionária
        distribuidora = "Distribuidora Desconhecida"
        concessionaria_aneel = None
        for key in [
            "CEMIG", "ELETROPAULO", "ENEL", "CPFL", "LIGHT", "COPEL",
            "CELESC", "ELEKTRO", "COELBA", "COSERN", "CELPE", "EQUATORIAL",
            "EDP", "ENERGISA", "RORAIMA", "AMAZONAS"
        ]:
            if key in clean_text:
                distribuidora = key
                break

        # Tenta mapear com a base ANEEL
        matched = tariff_client.match_concessionaria(distribuidora)
        if matched:
            concessionaria_aneel = matched
            distribuidora = matched
            campos.append("distribuidora")

        # 2. Subgrupo Tarifário (A4, A3a, etc.)
        subgrupo = "A4"
        if re.search(r"\bA3A\b", clean_text):
            subgrupo = "A3a"
            campos.append("subgrupo")
        elif re.search(r"\bA4\b", clean_text) or "GRUPO A" in clean_text:
            subgrupo = "A4"
            campos.append("subgrupo")

        # 3. Ambiente de Contratação: Cativo (ACR) vs. Livre (ACL)
        ambiente = "CATIVO"
        termos_livre = [
            "MERCADO LIVRE", "CONSUMIDOR LIVRE", "CONSUMIDOR ESPECIAL",
            "ACL", "CUSD", "USO DO SISTEMA", "CCEE", "CONTRATO DE USO",
            "FATURA LIVRE", "USO DA REDE"
        ]
        if any(term in clean_text for term in termos_livre):
            ambiente = "LIVRE"
            campos.append("ambiente_contratacao")
        # Se na tabela de faturamento não constar tarifa de energia (TE) ou constar apenas TUSD
        elif "TUSD" in clean_text and not any(x in clean_text for x in [" TE ", "ENERGIA TE", "TARIFA DE ENERGIA"]):
            ambiente = "LIVRE"
            campos.append("ambiente_contratacao")

        # 4. Modalidade Tarifária (Verde ou Azul)
        modalidade = "VERDE"
        if re.search(r"\bAZUL\b", clean_text) or "HORARIA AZUL" in clean_text:
            modalidade = "AZUL"
            campos.append("modalidade")
        elif re.search(r"\bVERDE\b", clean_text) or "HORARIA VERDE" in clean_text:
            modalidade = "VERDE"
            campos.append("modalidade")

        # 5. Mês de Referência (ex: 05/2026 ou MAI/2026)
        mes_ref = "05/2026"
        m_ref = re.search(r"(?:MES\s*(?:DE\s*)?REF(?:ERENCIA)?|REFERENCIA|EMISSAO)[\s:]*([0-9]{2}/[0-9]{4}|[A-Z]{3}/[0-9]{4})", clean_text)
        if m_ref:
            mes_ref = m_ref.group(1)
            campos.append("mes_referencia")

        # 5. Demanda Contratada (kW)
        demanda_contratada = 300.0
        demanda_contratada_ponta = None
        m_dem_contr = re.findall(r"(?:DEMANDA|DEM)\s*(?:CONTRATADA|CONTR)?[\s:]+([0-9\.,]+)", clean_text)
        if m_dem_contr:
            vals = [_parse_float_ptbr(v) for v in m_dem_contr if _parse_float_ptbr(v) is not None]
            valid_vals = [v for v in vals if 30.0 <= v <= 20000.0]
            if valid_vals:
                demanda_contratada = valid_vals[0]
                if modalidade == "AZUL" and len(valid_vals) >= 2:
                    demanda_contratada_ponta = valid_vals[1]
                campos.append("demanda_contratada")

        # 6. Demanda Medida (kW)
        demanda_medida = demanda_contratada * 0.95
        demanda_medida_ponta = None
        m_dem_med = re.findall(r"(?:DEMANDA|DEM)\s*(?:MEDIDA|MED)?[\s:]+([0-9\.,]+)", clean_text)
        if m_dem_med:
            vals = [_parse_float_ptbr(v) for v in m_dem_med if _parse_float_ptbr(v) is not None]
            valid_vals = [v for v in vals if 30.0 <= v <= 20000.0]
            if valid_vals:
                demanda_medida = valid_vals[0]
                if len(valid_vals) >= 2:
                    demanda_medida_ponta = valid_vals[1]
                campos.append("demanda_medida")

        # 7. Consumo Ativo Ponta (kWh)
        consumo_ponta = 15000.0
        m_cons_ponta = re.findall(r"(?:CONSUMO|ENERGIA|ATIVA)?\s*(?:PONTA|P)[\s:]+([0-9\.,]+)\s*KWH", clean_text)
        if not m_cons_ponta:
            m_cons_ponta = re.findall(r"(?:PONTA|P)\s*KWH[\s:]+([0-9\.,]+)", clean_text)
        if m_cons_ponta:
            vals = [_parse_float_ptbr(v) for v in m_cons_ponta if _parse_float_ptbr(v) is not None]
            valid_vals = [v for v in vals if v > 100.0]
            if valid_vals:
                consumo_ponta = valid_vals[0]
                campos.append("consumo_ponta")

        # 8. Consumo Ativo Fora de Ponta (kWh)
        consumo_fora_ponta = 85000.0
        m_cons_fp = re.findall(r"(?:CONSUMO|ENERGIA|ATIVA)?\s*(?:FORA\s*PONTA|F\.?\s*PONTA|FP)[\s:]+([0-9\.,]+)\s*KWH", clean_text)
        if not m_cons_fp:
            m_cons_fp = re.findall(r"(?:FORA\s*PONTA|FP)\s*KWH[\s:]+([0-9\.,]+)", clean_text)
        if m_cons_fp:
            vals = [_parse_float_ptbr(v) for v in m_cons_fp if _parse_float_ptbr(v) is not None]
            valid_vals = [v for v in vals if v > 100.0]
            if valid_vals:
                consumo_fora_ponta = valid_vals[0]
                campos.append("consumo_fora_ponta")

        # 9. Valor Total da Fatura (R$)
        valor_total = 0.0
        m_total = re.findall(r"(?:TOTAL\s*A\s*PAGAR|VALOR\s*TOTAL|TOTAL\s*FATURA|VALOR\s*A\s*PAGAR)[\s:]*(?:R\$)?\s*([0-9\.,]+)", clean_text)
        if m_total:
            vals = [_parse_float_ptbr(v) for v in m_total if _parse_float_ptbr(v) is not None]
            valid_vals = [v for v in vals if v > 500.0]
            if valid_vals:
                valor_total = valid_vals[0]
                campos.append("valor_total")

        # 10. Tributos (ICMS, PIS, COFINS, CIP)
        icms_aliquota = 18.0
        m_icms = re.search(r"ICMS[^\n]*?([0-9]{1,2}(?:,[0-9]{1,2})?)\s*%", clean_text)
        if m_icms:
            val = _parse_float_ptbr(m_icms.group(1))
            if val and 5.0 <= val <= 35.0:
                icms_aliquota = val
                campos.append("icms")

        pis_aliquota = 1.65
        cofins_aliquota = 7.60
        cip_valor = 0.0
        m_cip = re.search(r"(?:CIP|COSIP|ILUM\.?\s*PUB)[\s:]*(?:R\$)?\s*([0-9\.,]+)", clean_text)
        if m_cip:
            val = _parse_float_ptbr(m_cip.group(1))
            if val and 0.0 <= val <= 5000.0:
                cip_valor = val
                campos.append("cip_cosip")

        tributos = FaturaTributos(
            icms_aliquota=icms_aliquota,
            pis_aliquota=pis_aliquota,
            cofins_aliquota=cofins_aliquota,
            cip_cosip=cip_valor,
        )

        # 11. Bandeira Tarifária
        bandeira_tipo = "Verde"
        bandeira_reais = 0.0
        if "AMARELA" in clean_text:
            bandeira_tipo = "Amarela"
            bandeira_reais = (consumo_ponta + consumo_fora_ponta) * 0.01885
            campos.append("bandeira")
        elif "VERMELHA" in clean_text:
            bandeira_tipo = "Vermelha"
            bandeira_reais = (consumo_ponta + consumo_fora_ponta) * 0.04463
            campos.append("bandeira")

        # Cálculo do Score de Confiança
        weights = {
            "distribuidora": 0.25,
            "modalidade": 0.15,
            "consumo_ponta": 0.20,
            "consumo_fora_ponta": 0.20,
            "demanda_contratada": 0.10,
            "valor_total": 0.10,
        }
        score = sum(weights.get(c, 0.05) for c in set(campos))
        score = min(1.0, max(0.2, score))

        # Se valor total não foi lido, estima valor razoável para não ficar zerado
        if valor_total <= 0:
            # Estimativa básica caso o PDF esteja cortado no rodapé
            valor_total = (demanda_contratada * 25.0) + (consumo_ponta * 1.5) + (consumo_fora_ponta * 0.45)
            valor_total *= (1.0 + (tributos.aliquota_efetiva_total / 100.0))

        return InvoiceData(
            distribuidora_identificada=distribuidora,
            concessionaria_homologada_aneel=concessionaria_aneel,
            subgrupo=subgrupo,
            ambiente_contratacao=ambiente,
            modalidade=modalidade,
            mes_referencia=mes_ref,
            demanda_contratada_kw=demanda_contratada,
            demanda_contratada_ponta_kw=demanda_contratada_ponta,
            demanda_medida_kw=demanda_medida,
            demanda_medida_ponta_kw=demanda_medida_ponta,
            consumo_ponta_kwh=consumo_ponta,
            consumo_fora_ponta_kwh=consumo_fora_ponta,
            valor_total_faturado=round(valor_total, 2),
            bandeira_tipo=bandeira_tipo,
            bandeira_adicional_reais=round(bandeira_reais, 2),
            tributos=tributos,
            confidence_score=round(score, 2),
            campos_detectados=campos,
            raw_text_preview=clean_text[:800],
        )

    @classmethod
    def get_sample_invoices(cls) -> Dict[str, InvoiceData]:
        """Retorna faturas de exemplo realistas do Grupo A4 nos ambientes Cativo (ACR) e Livre (ACL)."""
        return {
            "CEMIG A4 Verde - Indústria Têxtil (Mercado Cativo ACR - 350 kW)": InvoiceData(
                distribuidora_identificada="CEMIG-D",
                concessionaria_homologada_aneel="CEMIG-D",
                subgrupo="A4",
                ambiente_contratacao="CATIVO",
                modalidade="VERDE",
                mes_referencia="05/2026",
                numero_instalacao="3009845120",
                demanda_contratada_kw=350.0,
                demanda_medida_kw=335.0,
                consumo_ponta_kwh=18500.0,
                consumo_fora_ponta_kwh=96000.0,
                valor_total_faturado=124650.80,
                bandeira_tipo="Verde",
                bandeira_adicional_reais=0.0,
                tributos=FaturaTributos(icms_aliquota=18.0, pis_aliquota=1.65, cofins_aliquota=7.60, cip_cosip=180.0),
                confidence_score=1.0,
                campos_detectados=["distribuidora", "modalidade", "ambiente_contratacao", "demanda_contratada", "consumo_ponta", "consumo_fora_ponta", "valor_total"],
            ),
            "ELETROPAULO (Enel SP) A4 Azul - Shopping Center (Mercado Livre ACL - 800 kW)": InvoiceData(
                distribuidora_identificada="ELETROPAULO",
                concessionaria_homologada_aneel="ELETROPAULO",
                subgrupo="A4",
                ambiente_contratacao="LIVRE",
                preco_energia_acl_kwh=0.23,
                modalidade="AZUL",
                mes_referencia="05/2026",
                numero_instalacao="7014529011",
                demanda_contratada_kw=800.0,
                demanda_contratada_ponta_kw=450.0,
                demanda_medida_kw=780.0,
                demanda_medida_ponta_kw=430.0,
                consumo_ponta_kwh=32000.0,
                consumo_fora_ponta_kwh=210000.0,
                valor_total_faturado=86540.20,  # Fatura da concessionária no ACL contém apenas TUSD + impostos
                bandeira_tipo="Verde",
                bandeira_adicional_reais=0.0,  # Consumidor livre não paga bandeira na fatura de distribuição
                tributos=FaturaTributos(icms_aliquota=18.0, pis_aliquota=1.65, cofins_aliquota=7.60, cip_cosip=350.0),
                confidence_score=1.0,
                campos_detectados=["distribuidora", "modalidade", "ambiente_contratacao", "demanda_contratada", "consumo_ponta", "consumo_fora_ponta", "valor_total"],
            ),
            "CPFL Paulista A4 Verde - Metalúrgica (Mercado Livre ACL - 500 kW)": InvoiceData(
                distribuidora_identificada="CPFL-PAULISTA",
                concessionaria_homologada_aneel="CPFL-PAULISTA",
                subgrupo="A4",
                ambiente_contratacao="LIVRE",
                preco_energia_acl_kwh=0.21,
                modalidade="VERDE",
                mes_referencia="05/2026",
                numero_instalacao="5028419302",
                demanda_contratada_kw=500.0,
                demanda_medida_kw=485.0,
                consumo_ponta_kwh=24000.0,
                consumo_fora_ponta_kwh=145000.0,
                valor_total_faturado=67890.50,  # Fatura CUSD (apenas TUSD)
                bandeira_tipo="Verde",
                bandeira_adicional_reais=0.0,
                tributos=FaturaTributos(icms_aliquota=18.0, pis_aliquota=1.65, cofins_aliquota=7.60, cip_cosip=220.0),
                confidence_score=1.0,
                campos_detectados=["distribuidora", "modalidade", "ambiente_contratacao", "demanda_contratada", "consumo_ponta", "consumo_fora_ponta", "valor_total"],
            ),
            "EQUATORIAL PA A4 Verde - Supermercado / CD (Mercado Cativo ACR - 400 kW)": InvoiceData(
                distribuidora_identificada="EQUATORIAL PA",
                concessionaria_homologada_aneel="EQUATORIAL PA",
                subgrupo="A4",
                ambiente_contratacao="CATIVO",
                modalidade="VERDE",
                mes_referencia="05/2026",
                numero_instalacao="4019283741",
                demanda_contratada_kw=400.0,
                demanda_medida_kw=390.0,
                consumo_ponta_kwh=22000.0,
                consumo_fora_ponta_kwh=115000.0,
                valor_total_faturado=142350.80,
                bandeira_tipo="Verde",
                bandeira_adicional_reais=0.0,
                tributos=FaturaTributos(icms_aliquota=20.0, pis_aliquota=1.65, cofins_aliquota=7.60, cip_cosip=250.0),
                confidence_score=1.0,
                campos_detectados=["distribuidora", "modalidade", "ambiente_contratacao", "demanda_contratada", "consumo_ponta", "consumo_fora_ponta", "valor_total"],
            ),
        }


invoice_parser = InvoiceParser()
