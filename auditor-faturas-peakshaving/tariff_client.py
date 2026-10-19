"""
tariff_client.py
----------------
Módulo responsável por consultar, raspar e filtrar as tarifas homologadas da ANEEL
a partir da aplicação web oficial do projeto:
https://englucas1.github.io/Aneel-Tarifas_v3/

Fornece mapeamento de aliases de distribuidoras brasileiras, cache inteligente
e extração estruturada dos componentes tarifários (Demanda, TUSD e TE) para Grupo A4.
"""

from __future__ import annotations

import json
import logging
import re
import unicodedata
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import requests
from bs4 import BeautifulSoup
from pydantic import BaseModel, Field

logger = logging.getLogger(__name__)

PRIMARY_URL = "https://englucas1.github.io/Aneel-Tarifas_v3/data/tarifas_consolidated.json"
BASE_SITE_URL = "https://englucas1.github.io/Aneel-Tarifas_v3/"
FALLBACK_LOCAL_PATH = Path(__file__).parent / "data" / "tarifas_consolidated.json"


class TarifaModalidadeVerde(BaseModel):
    demanda_unica: float = Field(..., description="Tarifa de Demanda Única (R$/kW)")
    tusd_ponta: float = Field(..., description="TUSD Ponta (R$/kWh)")
    tusd_fora_ponta: float = Field(..., description="TUSD Fora Ponta (R$/kWh)")
    te_ponta: float = Field(..., description="TE Ponta (R$/kWh)")
    te_fora_ponta: float = Field(..., description="TE Fora Ponta (R$/kWh)")

    @property
    def total_ponta(self) -> float:
        """Custo unitário de energia na ponta (TUSD + TE) em R$/kWh."""
        return self.tusd_ponta + self.te_ponta

    @property
    def total_fora_ponta(self) -> float:
        """Custo unitário de energia fora de ponta (TUSD + TE) em R$/kWh."""
        return self.tusd_fora_ponta + self.te_fora_ponta


class TarifaModalidadeAzul(BaseModel):
    demanda_ponta: float = Field(..., description="Tarifa de Demanda Ponta (R$/kW)")
    demanda_fora_ponta: float = Field(..., description="Tarifa de Demanda Fora Ponta (R$/kW)")
    tusd_ponta: float = Field(..., description="TUSD Ponta (R$/kWh)")
    tusd_fora_ponta: float = Field(..., description="TUSD Fora Ponta (R$/kWh)")
    te_ponta: float = Field(..., description="TE Ponta (R$/kWh)")
    te_fora_ponta: float = Field(..., description="TE Fora Ponta (R$/kWh)")

    @property
    def total_ponta(self) -> float:
        """Custo unitário de energia na ponta (TUSD + TE) em R$/kWh."""
        return self.tusd_ponta + self.te_ponta

    @property
    def total_fora_ponta(self) -> float:
        """Custo unitário de energia fora de ponta (TUSD + TE) em R$/kWh."""
        return self.tusd_fora_ponta + self.te_fora_ponta


class ConcessionariaTarifas(BaseModel):
    agente: str
    subgrupo: str = "A4"
    cnpj: Optional[str] = None
    resolucao_homologatoria: str
    data_vigencia: Optional[str] = None
    verde: TarifaModalidadeVerde
    azul: TarifaModalidadeAzul
    resolucao_anterior: Optional[str] = None
    vigencia_anterior: Optional[str] = None
    fonte: str = "https://englucas1.github.io/Aneel-Tarifas_v3/"


def _normalize_name(name: str) -> str:
    """Normaliza texto para comparação fonética/lexical de nomes de distribuidoras."""
    if not name:
        return ""
    text = unicodedata.normalize("NFKD", name)
    text = "".join(c for c in text if not unicodedata.combining(c)).upper()
    text = re.sub(r"[^A-Z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


# Dicionário de sinônimos/aliases que mapeiam nomes comerciais às chaves oficiais da ANEEL
CONCESSIONARIA_ALIASES: Dict[str, List[str]] = {
    "CEMIG-D": ["CEMIG", "COMPANHIA ENERGETICA DE MINAS GERAIS", "CEMIG DISTRIBUICAO", "CEMIG D", "CEMIG-D"],
    "CPFL-PAULISTA": ["CPFL", "CPFL PAULISTA", "COMPANHIA PAULISTA DE FORCA E LUZ", "PAULISTA"],
    "CPFL-PIRATINING": ["PIRATININGA", "CPFL PIRATININGA", "CPFL-PIRATININGA"],
    "CPFL Santa Cruz": ["SANTA CRUZ", "CPFL SANTA CRUZ"],
    "ELETROPAULO": ["ENEL SP", "ELETROPAULO", "ENEL SAO PAULO", "ELETROPAULO METROPOLITANA", "ENEL S.A."],
    "ENEL RJ": ["ENEL RJ", "AMPLA", "ENEL RIO", "AMPLA ENERGIA"],
    "ENEL CE": ["ENEL CE", "COELCE", "COMPANHIA ENERGETICA DO CEARA", "ENEL CEARA"],
    "LIGHT SESA": ["LIGHT", "LIGHT SERVICOS DE ELETRICIDADE", "LIGHT RJ", "LIGHT SESA"],
    "COPEL-DIS": ["COPEL", "COPEL DISTRIBUICAO", "COPEL DIS", "COMPANHIA PARANAENSE DE ENERGIA"],
    "CELESC-DIS": ["CELESC", "CELESC DISTRIBUICAO", "CELESC DIS", "CENTRAIS ELETRICAS DE SANTA CATARINA"],
    "ELEKTRO": ["ELEKTRO", "NEOENERGIA ELEKTRO", "ELEKTRO REDES"],
    "COELBA": ["COELBA", "NEOENERGIA COELBA", "COMPANHIA DE ELETRICIDADE DO ESTADO DA BAHIA"],
    "Neoenergia PE": ["CELPE", "NEOENERGIA PERNAMBUCO", "NEOENERGIA PE", "COMPANHIA ENERGETICA DE PERNAMBUCO"],
    "Neoenergia Brasília": ["CEB", "NEOENERGIA BRASILIA", "CEB DISTRIBUICAO", "CEB D"],
    "EQUATORIAL PA": ["CELPA", "EQUATORIAL PARA", "EQUATORIAL PA", "CENTRAIS ELETRICAS DO PARA"],
    "EQUATORIAL MA": ["CEMAR", "EQUATORIAL MARANHAO", "EQUATORIAL MA", "COMPANHIA ENERGETICA DO MARANHAO"],
    "EQUATORIAL PI": ["CEPISA", "EQUATORIAL PIAUI", "EQUATORIAL PI"],
    "EQUATORIAL AL": ["CEAL", "EQUATORIAL ALAGOAS", "EQUATORIAL AL"],
    "EQUATORIAL GO": ["EQUATORIAL GOIAS", "EQUATORIAL GO", "ENEL GOIAS", "CELG"],
    "EDP SP": ["EDP SP", "EDP SAO PAULO", "BANDEIRANTE", "EDP BANDEIRANTE"],
    "EDP ES": ["EDP ES", "EDP ESPIRITO SANTO", "ESCELSA", "ESPIRITO SANTO CENTRAIS ELETRICAS"],
    "EMR": ["ENERGISA MG", "ENERGISA MINAS GERAIS", "ENERGISA MINAS RIO", "EMR"],
    "EMT": ["ENERGISA MT", "ENERGISA MATO GROSSO", "CEMAT", "EMT"],
    "EMS": ["ENERGISA MS", "ENERGISA MATO GROSSO DO SUL", "ENERSUL", "EMS"],
    "EPB": ["ENERGISA PB", "ENERGISA PARAIBA", "SAELPA", "EPB"],
    "ESE": ["ENERGISA SE", "ENERGISA SERGIPE", "ENERGIPE", "ESE"],
    "ETO": ["ENERGISA TO", "ENERGISA TOCANTINS", "CELTINS", "ETO"],
    "ERO": ["ENERGISA RO", "ENERGISA RONDONIA", "CERON", "ERO"],
    "EAC": ["ENERGISA AC", "ENERGISA ACRE", "ELETROACRE", "EAC"],
    "ESS": ["ENERGISA SUL SUDESTE", "ESS"],
    "EBO": ["ENERGISA BORBOREMA", "EBO"],
    "ÂMBAR ENERGIA RR": ["RORAIMA ENERGIA", "CER", "ELETRORORAIMA", "AMBAR ENERGIA RR"],
    "AME": ["AMAZONAS ENERGIA", "AME", "ELETROBRAS AMAZONAS"],
}


class AneelTariffClient:
    """Cliente para consulta e sincronização com https://englucas1.github.io/Aneel-Tarifas_v3/."""

    def __init__(self, cache_ttl_seconds: int = 3600):
        self.cache_ttl = cache_ttl_seconds
        self._raw_data: Optional[Dict[str, Any]] = None
        self._available_agents: List[str] = []

    def fetch_database(self, force_refresh: bool = False) -> Dict[str, Any]:
        """
        Obtém a base tarifária completa. Tenta primeiro via HTTP da URL oficial
        https://englucas1.github.io/Aneel-Tarifas_v3/data/tarifas_consolidated.json.
        Em caso de erro ou offline, carrega a cópia local sincronizada.
        """
        if self._raw_data and not force_refresh:
            return self._raw_data

        # Tentativa 1: Requisição direta ao JSON consolidado no GitHub Pages
        try:
            headers = {
                "User-Agent": "Aneel-Tariff-Auditor/1.0 (Peak-Shaving Analyzer; Python)",
                "Accept": "application/json",
            }
            resp = requests.get(PRIMARY_URL, headers=headers, timeout=8)
            if resp.status_code == 200:
                data = resp.json()
                if isinstance(data, dict) and len(data) > 0:
                    self._raw_data = data
                    logger.info("Tarifas carregadas com sucesso via HTTP da fonte oficial: %d chaves", len(data))
                    return self._raw_data
        except Exception as ex:
            logger.warning("Falha ao obter dados via HTTP (%s). Acionando fallback local.", ex)

        # Tentativa 2: Scraping de apoio se o site HTML estiver acessível
        try:
            resp_html = requests.get(BASE_SITE_URL, timeout=6)
            if resp_html.status_code == 200:
                soup = BeautifulSoup(resp_html.text, "html.parser")
                scripts = [s.get("src", "") for s in soup.find_all("script")]
                logger.debug("Scripts identificados no portal ANEEL: %s", scripts)
        except Exception:
            pass

        # Tentativa 3: Fallback local robusto
        if FALLBACK_LOCAL_PATH.exists():
            try:
                with open(FALLBACK_LOCAL_PATH, "r", encoding="utf-8") as f:
                    self._raw_data = json.load(f)
                    logger.info("Tarifas carregadas da cópia local (%s): %d chaves", FALLBACK_LOCAL_PATH, len(self._raw_data))
                    return self._raw_data
            except Exception as ex:
                logger.error("Falha ao ler arquivo local de tarifas: %s", ex)

        raise RuntimeError(
            "Não foi possível obter a base de dados tarifária nem via HTTP nem da cópia local."
        )

    def list_concessionarias(self, subgrupo: str = "A4") -> List[str]:
        """Retorna a lista ordenada de concessionárias disponíveis para o subgrupo informado."""
        data = self.fetch_database()
        agents = set()
        suffix = f"|{subgrupo}"
        for key, val in data.items():
            if key.endswith(suffix):
                agente_name = val.get("Agente") or key.split("|")[0]
                agents.add(agente_name)
        return sorted(list(agents))

    def match_concessionaria(self, query: str) -> Optional[str]:
        """
        Encontra o nome exato da concessionária na base a partir de texto livre ou extraído de PDF.
        Aplica busca exata, busca por alias e comparação de similaridade de tokens.
        """
        data = self.fetch_database()
        all_agents = self.list_concessionarias("A4")
        if not all_agents:
            all_agents = [k.split("|")[0] for k in data.keys()]

        norm_query = _normalize_name(query)
        if not norm_query:
            return None

        # 1. Correspondência exata normalizada com agentes da base
        for agent in all_agents:
            if _normalize_name(agent) == norm_query:
                return agent

        # 2. Busca via tabela de aliases
        for canonical, aliases in CONCESSIONARIA_ALIASES.items():
            norm_canonical = _normalize_name(canonical)
            if norm_canonical in norm_query or norm_query in norm_canonical:
                for agent in all_agents:
                    if _normalize_name(agent) == norm_canonical or canonical.upper() in agent.upper():
                        return agent
            for alias in aliases:
                norm_alias = _normalize_name(alias)
                if norm_alias in norm_query or norm_query == norm_alias:
                    for agent in all_agents:
                        if _normalize_name(agent) == norm_canonical or canonical.upper() in agent.upper():
                            return agent

        # 3. Busca por contenção de palavras-chave
        best_match = None
        best_score = 0
        query_tokens = set(norm_query.split())

        for agent in all_agents:
            agent_tokens = set(_normalize_name(agent).split())
            intersection = query_tokens.intersection(agent_tokens)
            score = len(intersection)
            if score > best_score:
                best_score = score
                best_match = agent

        if best_score >= 1 and best_match:
            return best_match

        return None

    def get_tarifas(self, concessionaria: str, subgrupo: str = "A4") -> ConcessionariaTarifas:
        """
        Retorna os componentes tarifários completos (Verde e Azul) vigentes
        homologados para o Grupo A4.
        """
        data = self.fetch_database()
        matched_agent = self.match_concessionaria(concessionaria)
        target_agent = matched_agent or concessionaria

        target_key = f"{target_agent}|{subgrupo}"
        item = data.get(target_key)

        # Se não encontrar com o subgrupo exato, busca chave correspondente
        if not item:
            for k, v in data.items():
                if k.lower().startswith(target_agent.lower()) and f"|{subgrupo}".lower() in k.lower():
                    item = v
                    target_key = k
                    break

        if not item:
            raise KeyError(
                f"Tarifas não localizadas para a concessionária '{concessionaria}' "
                f"(agente mapeado: '{target_agent}') e subgrupo '{subgrupo}'."
            )

        atual = item.get("Atual", {})
        verde_data = atual.get("Verde", {})
        azul_data = atual.get("Azul", {})
        anterior = item.get("Anterior", {})

        verde_obj = TarifaModalidadeVerde(
            demanda_unica=float(verde_data.get("Demanda", 0.0)),
            tusd_ponta=float(verde_data.get("TUSD_P", 0.0)),
            tusd_fora_ponta=float(verde_data.get("TUSD_FP", 0.0)),
            te_ponta=float(verde_data.get("TE_P", 0.0)),
            te_fora_ponta=float(verde_data.get("TE_FP", 0.0)),
        )

        azul_obj = TarifaModalidadeAzul(
            demanda_ponta=float(azul_data.get("Demanda_P", 0.0)),
            demanda_fora_ponta=float(azul_data.get("Demanda_FP", 0.0)),
            tusd_ponta=float(azul_data.get("TUSD_P", 0.0)),
            tusd_fora_ponta=float(azul_data.get("TUSD_FP", 0.0)),
            te_ponta=float(azul_data.get("TE_P", 0.0)),
            te_fora_ponta=float(azul_data.get("TE_FP", 0.0)),
        )

        return ConcessionariaTarifas(
            agente=item.get("Agente", target_agent),
            subgrupo=item.get("SubGrupo", subgrupo),
            cnpj=item.get("CNPJ"),
            resolucao_homologatoria=atual.get("REH") or atual.get("RawREH", "REH ANEEL Vigente"),
            data_vigencia=atual.get("Vigencia"),
            verde=verde_obj,
            azul=azul_obj,
            resolucao_anterior=anterior.get("REH") if anterior else None,
            vigencia_anterior=anterior.get("Vigencia") if anterior else None,
        )


# Instância padrão para uso geral
tariff_client = AneelTariffClient()
