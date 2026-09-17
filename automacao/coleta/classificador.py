"""
Classificação de documentos coletados (boleto, nota fiscal, demonstrativo,
autorização de pagamento).

A decisão sai da soma de dois conjuntos de sinais:

1. **Nome do arquivo** — padrões que a base real usa há anos
   (``Boleto_363700.pdf``, ``nfs_0104_359212.PDF``, ``Fatura_Extrato - *.pdf``…).
2. **Texto da 1ª página** — lido com ``pypdf``. É o sinal que decide os casos
   em que o nome não diz nada (``INTERNET-LINK.pdf``, ``fatura X.pdf``).

Cada sinal tem um peso. O tipo com maior pontuação vence e o `motivo` do
`Documento` diz, em português, quais sinais pesaram — é esse texto que o
painel mostra para o humano conferir.

Nada aqui escreve arquivo: é tudo leitura.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Iterable, Sequence

from automacao.nucleo.modelos import Documento, TipoDocumento, normalizar

log = logging.getLogger("automacao.coleta.classificador")

__all__ = [
    "classificar",
    "classificar_lote",
    "extrair_texto_primeira_pagina",
    "extrair_valor",
    "extrair_vencimento",
    "extrair_numero_documento",
    "extrair_cnpjs",
]

# Quanto texto da 1ª página é suficiente para decidir. Faturas de operadora
# passam de 5 mil caracteres e não precisamos de tudo.
LIMITE_TEXTO = 12_000

# Extensões que o classificador sabe ler de verdade.
EXTENSOES_PDF = {".pdf"}
EXTENSOES_XML = {".xml"}


# --------------------------------------------------------------------------- #
# Expressões reaproveitadas
# --------------------------------------------------------------------------- #

# Linha digitável de boleto bancário (representação numérica, 5 blocos).
# Aceita espaço múltiplo entre os blocos porque cada banco espaça de um jeito.
RE_LINHA_DIGITAVEL = re.compile(
    r"\d{5}[.\s]\d{5}\s+\d{5}[.\s]\d{6}\s+\d{5}[.\s]\d{6}\s+\d\s+\d{14}"
)

# Linha digitável de arrecadação/concessionária (4 blocos de 11 dígitos + DV).
# É o que aparece nas faturas das operadoras de telefonia.
RE_LINHA_ARRECADACAO = re.compile(
    r"\b\d{11}[-\s]\d\s+\d{11}[-\s]\d\s+\d{11}[-\s]\d\s+\d{11}[-\s]\d\b"
)

# Valor monetário no padrão brasileiro: 1.234,56 ou 45083,16.
RE_DINHEIRO = re.compile(r"(?:\d{1,3}(?:\.\d{3})+|\d+),\d{2}")

RE_DATA_BR = re.compile(r"\b(\d{2})/(\d{2})/(\d{4})\b")

RE_CNPJ = re.compile(r"\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b")

# Nome no padrão do projeto: "... - 07-2026", "... - 2026-07", "...- 2026 -07".
RE_NOME_COMPETENCIA = re.compile(
    r"[-–]\s*(?:\d{2}\s*[-_ ]\s*\d{4}|\d{4}\s*[-_ ]\s*\d{2})\s*$"
)

# Nome tipo "9829452_2026-05" — relatório mensal identificado por número de conta.
RE_NOME_RELATORIO_MENSAL = re.compile(r"^\d{5,}[_\-]\d{4}[-_]\d{2}$")


# --------------------------------------------------------------------------- #
# Catálogo de sinais
# --------------------------------------------------------------------------- #


@dataclass(frozen=True)
class Sinal:
    """Um indício de que o arquivo é de determinado tipo."""

    padrao: re.Pattern[str]
    tipo: TipoDocumento
    peso: float
    descricao: str


def _p(expressao: str) -> re.Pattern[str]:
    """Compila a expressão já pensando em texto normalizado (sem acento, minúsculo)."""
    return re.compile(expressao)


# --- sinais do nome do arquivo --------------------------------------------- #
# Peso menor que os do texto: nome é palpite, texto é prova.

SINAIS_NOME: tuple[Sinal, ...] = (
    Sinal(_p(r"\bboleto"), TipoDocumento.BOLETO, 3.0, "nome do arquivo contém 'boleto'"),
    Sinal(_p(r"bloqueto"), TipoDocumento.BOLETO, 3.0, "nome do arquivo contém 'bloqueto'"),
    Sinal(_p(r"boletobancario"), TipoDocumento.BOLETO, 1.5, "nome do arquivo contém 'boletobancario'"),
    Sinal(_p(r"2\s*via|segunda via"), TipoDocumento.BOLETO, 1.0, "nome do arquivo fala em 2ª via"),
    Sinal(_p(r"\bcobranca\b"), TipoDocumento.BOLETO, 1.0, "nome do arquivo contém 'cobranca'"),

    Sinal(_p(r"\bdanfe\b"), TipoDocumento.NOTA_FISCAL, 3.5, "nome do arquivo contém 'danfe'"),
    Sinal(_p(r"nota fiscal|notafiscal|documento fiscal"), TipoDocumento.NOTA_FISCAL, 3.0,
          "nome do arquivo contém 'nota fiscal'"),
    Sinal(_p(r"\bnfs?e?\b|\bnfs[_\-]|^nfs|nfse|nf-e"), TipoDocumento.NOTA_FISCAL, 3.0,
          "nome do arquivo contém 'nf'/'nfs'/'nfse'"),
    Sinal(_p(r"^nota\b|\bnota\b"), TipoDocumento.NOTA_FISCAL, 1.5, "nome do arquivo contém 'nota'"),
    Sinal(_p(r"\binvoice\b"), TipoDocumento.NOTA_FISCAL, 1.5, "nome do arquivo contém 'invoice'"),

    Sinal(_p(r"extrato"), TipoDocumento.DEMONSTRATIVO, 3.0, "nome do arquivo contém 'extrato'"),
    Sinal(_p(r"demonstrativo"), TipoDocumento.DEMONSTRATIVO, 3.0, "nome do arquivo contém 'demonstrativo'"),
    Sinal(_p(r"relatorio"), TipoDocumento.DEMONSTRATIVO, 2.5, "nome do arquivo contém 'relatorio'"),
    Sinal(_p(r"detalhamento|detalhe"), TipoDocumento.DEMONSTRATIVO, 2.0, "nome do arquivo contém 'detalhamento'"),
    Sinal(_p(r"\bseats?\b|consumo|resumo"), TipoDocumento.DEMONSTRATIVO, 1.5,
          "nome do arquivo contém 'seats'/'consumo'/'resumo'"),

    Sinal(_p(r"autorizacao"), TipoDocumento.AUTORIZACAO, 3.0, "nome do arquivo contém 'autorizacao'"),
)

# --- sinais do texto da 1ª página ------------------------------------------ #

SINAIS_TEXTO: tuple[Sinal, ...] = (
    # -- boleto --
    Sinal(RE_LINHA_DIGITAVEL, TipoDocumento.BOLETO, 5.0,
          "texto da página 1 tem linha digitável de boleto bancário"),
    Sinal(_p(r"recibo do pagador|recibo do sacado"), TipoDocumento.BOLETO, 3.0,
          "texto tem 'RECIBO DO PAGADOR'"),
    Sinal(_p(r"ficha de compensacao"), TipoDocumento.BOLETO, 2.5,
          "texto tem 'Ficha de Compensação'"),
    Sinal(RE_LINHA_ARRECADACAO, TipoDocumento.BOLETO, 2.5,
          "texto tem linha digitável de arrecadação (código de barras de concessionária)"),
    Sinal(_p(r"local de pagamento"), TipoDocumento.BOLETO, 1.5, "texto tem 'Local de Pagamento'"),
    Sinal(_p(r"boleto bancario"), TipoDocumento.BOLETO, 1.2, "texto tem 'BOLETO BANCÁRIO'"),
    Sinal(_p(r"pagavel em qualquer banco|em qualquer banco ou corresp|qualquer banco ate o vencimento"),
          TipoDocumento.BOLETO, 1.0, "texto tem 'pagável em qualquer banco'"),
    Sinal(_p(r"autenticacao mecanica"), TipoDocumento.BOLETO, 1.0, "texto tem 'Autenticação Mecânica'"),
    Sinal(_p(r"\bcedente\b|sacador|\bsacado\b"), TipoDocumento.BOLETO, 0.8,
          "texto tem 'cedente'/'sacador'"),
    Sinal(_p(r"nosso numero"), TipoDocumento.BOLETO, 1.0, "texto tem 'Nosso número'"),
    Sinal(_p(r"linha digitavel"), TipoDocumento.BOLETO, 2.0, "texto fala em 'linha digitável'"),

    # -- nota fiscal --
    Sinal(_p(r"nota fiscal de servico eletronica|nota fiscal de servicos eletronica"),
          TipoDocumento.NOTA_FISCAL, 5.0, "texto tem 'NOTA FISCAL DE SERVIÇO ELETRÔNICA'"),
    Sinal(_p(r"documento auxiliar da nota fiscal"), TipoDocumento.NOTA_FISCAL, 4.5,
          "texto tem 'DOCUMENTO AUXILIAR DA NOTA FISCAL' (DANFE)"),
    Sinal(_p(r"\bdanfe\b"), TipoDocumento.NOTA_FISCAL, 4.0, "texto tem 'DANFE'"),
    Sinal(_p(r"\bnfcom\b|nf-com"), TipoDocumento.NOTA_FISCAL, 4.0,
          "texto tem 'NFCom' (nota fiscal de serviço de comunicação)"),
    Sinal(_p(r"\bnfs-e\b|\bnfse\b|\bnf-e\b"), TipoDocumento.NOTA_FISCAL, 3.0, "texto tem 'NFS-e'"),
    Sinal(_p(r"prefeitura municipal|secretaria (?:municipal )?de financas"),
          TipoDocumento.NOTA_FISCAL, 3.0, "texto tem 'PREFEITURA MUNICIPAL'"),
    Sinal(_p(r"chave de acesso"), TipoDocumento.NOTA_FISCAL, 2.5, "texto tem 'Chave de acesso'"),
    Sinal(_p(r"codigo de verificacao"), TipoDocumento.NOTA_FISCAL, 2.0, "texto tem 'Código de Verificação'"),
    Sinal(_p(r"prestador de servico"), TipoDocumento.NOTA_FISCAL, 1.5, "texto tem 'PRESTADOR DE SERVIÇOS'"),
    Sinal(_p(r"tomador d[eo]"), TipoDocumento.NOTA_FISCAL, 1.0, "texto tem 'TOMADOR DOS SERVIÇOS'"),
    Sinal(_p(r"protocolo de autorizacao"), TipoDocumento.NOTA_FISCAL, 1.5,
          "texto tem 'Protocolo de autorização'"),
    Sinal(_p(r"\bcfop\b"), TipoDocumento.NOTA_FISCAL, 1.5, "texto tem 'CFOP'"),
    Sinal(_p(r"inscricao municipal"), TipoDocumento.NOTA_FISCAL, 1.0, "texto tem 'Inscrição Municipal'"),
    Sinal(_p(r"discriminacao do|descricao dos servicos"), TipoDocumento.NOTA_FISCAL, 1.0,
          "texto tem 'Discriminação dos Serviços'"),
    Sinal(_p(r"iss retido|base de calculo|base calculo|retencoes federais"),
          TipoDocumento.NOTA_FISCAL, 0.8, "texto tem campos fiscais (ISS/base de cálculo)"),
    Sinal(_p(r"\bnota fiscal\b"), TipoDocumento.NOTA_FISCAL, 1.0, "texto menciona 'nota fiscal'"),

    # -- demonstrativo --
    Sinal(_p(r"sem validade fiscal"), TipoDocumento.DEMONSTRATIVO, 4.0,
          "texto avisa que é 'documento sem validade fiscal'"),
    Sinal(_p(r"\bdemonstrativo\b"), TipoDocumento.DEMONSTRATIVO, 3.5, "texto tem 'Demonstrativo'"),
    Sinal(_p(r"resumo da (?:sua )?fatura"), TipoDocumento.DEMONSTRATIVO, 3.0,
          "texto tem 'RESUMO DA SUA FATURA'"),
    Sinal(_p(r"relatorio d[eo]"), TipoDocumento.DEMONSTRATIVO, 3.0, "texto tem 'Relatório de …'"),
    Sinal(_p(r"detalhamento"), TipoDocumento.DEMONSTRATIVO, 2.5, "texto tem 'Detalhamento'"),
    Sinal(_p(r"\bextrato\b"), TipoDocumento.DEMONSTRATIVO, 2.5, "texto tem 'Extrato'"),
    Sinal(_p(r"historico de cobranca"), TipoDocumento.DEMONSTRATIVO, 2.0,
          "texto tem 'Histórico de cobrança'"),
    Sinal(_p(r"periodo de uso|dias de uso|\bseats\b|quantidade de licencas"),
          TipoDocumento.DEMONSTRATIVO, 2.0, "texto lista itens/licenças consumidos"),
    Sinal(_p(r"simples conferencia|para conferencia"), TipoDocumento.DEMONSTRATIVO, 2.0,
          "texto diz que é 'para simples conferência'"),
    Sinal(_p(r"consumo unitario|valor unitario"), TipoDocumento.DEMONSTRATIVO, 1.0,
          "texto tem coluna de valor unitário"),
    Sinal(_p(r"numero da fatura|n[o°º] da fatura"), TipoDocumento.DEMONSTRATIVO, 2.0,
          "texto tem 'Número da fatura' (fatura/invoice de plataforma)"),
    Sinal(_p(r"resumo relativo ao periodo|referente ao periodo de"), TipoDocumento.DEMONSTRATIVO, 2.0,
          "texto tem 'resumo relativo ao período'"),
    Sinal(_p(r"\bfatura\b"), TipoDocumento.DEMONSTRATIVO, 0.5, "texto menciona 'fatura' (sinal fraco)"),

    # -- autorização de pagamento (o formulário interno da empresa) --
    Sinal(_p(r"autorizacao de pagamento"), TipoDocumento.AUTORIZACAO, 6.0,
          "texto tem 'AUTORIZAÇÃO DE PAGAMENTO'"),
    Sinal(_p(r"pgto previsto"), TipoDocumento.AUTORIZACAO, 2.0, "texto tem 'PGTO PREVISTO'"),
    Sinal(_p(r"dados do beneficiario"), TipoDocumento.AUTORIZACAO, 1.5,
          "texto tem 'DADOS DO BENEFICIÁRIO'"),
    Sinal(_p(r"centro de custo"), TipoDocumento.AUTORIZACAO, 1.5, "texto tem 'CENTRO DE CUSTO'"),
    Sinal(_p(r"nat\.? despesa|natureza da despesa"), TipoDocumento.AUTORIZACAO, 1.5,
          "texto tem 'NAT. DESPESA'"),
    Sinal(_p(r"forma pgto"), TipoDocumento.AUTORIZACAO, 1.5, "texto tem 'FORMA PGTO'"),
    Sinal(_p(r"solicitante.*(?:gerente financeiro|diretoria)"), TipoDocumento.AUTORIZACAO, 1.5,
          "texto tem o rodapé de assinaturas da autorização"),
)

# Combinação clássica de boleto: os três campos juntos valem mais que separados.
_TRIO_BOLETO = (_p(r"beneficiario"), _p(r"vencimento"), _p(r"nosso numero"))

# Âncoras usadas para achar o valor do documento.
ANCORAS_VALOR: tuple[str, ...] = (
    "valor do documento",
    "valor total da nota",
    "valor total da nf",
    "valor total",
    "total a pagar",
    "valor a pagar",
    "valor cobrado",
    "valor original",
    "valor liquido",
    "valor dos servicos",
    "valores dos servicos",
    "total em r$",
    "total da fatura",
    "valor da fatura",
    "total geral",
    "valor:",
)

JANELA_ANTES = 90
JANELA_DEPOIS = 160


# --------------------------------------------------------------------------- #
# Leitura do PDF
# --------------------------------------------------------------------------- #


def extrair_texto_primeira_pagina(caminho: Path) -> tuple[str, int | None, str | None]:
    """
    Lê o texto da primeira página do PDF.

    Devolve ``(texto, total_de_paginas, erro)``. Nunca levanta exceção: PDF
    protegido, corrompido ou só com imagem devolve texto vazio e o motivo em
    `erro`, para o painel exibir.
    """
    try:
        from pypdf import PdfReader
    except ImportError as exc:  # pragma: no cover - dependência declarada
        return "", None, f"pypdf indisponível ({exc})"

    try:
        leitor = PdfReader(str(caminho))
        if leitor.is_encrypted:
            # Muitos boletos vêm com senha vazia só para travar edição.
            try:
                if not leitor.decrypt(""):
                    return "", None, "PDF protegido por senha — não foi possível ler o texto"
            except Exception:
                return "", None, "PDF protegido por senha — não foi possível ler o texto"

        paginas = len(leitor.pages)
        if paginas == 0:
            return "", 0, "PDF sem páginas"
        texto = leitor.pages[0].extract_text() or ""
    except Exception as exc:
        log.warning("falha ao ler %s: %s: %s", caminho.name, type(exc).__name__, exc)
        return "", None, f"PDF ilegível ({type(exc).__name__})"

    if not texto.strip():
        return "", paginas, "PDF sem texto extraível (provavelmente digitalizado/imagem)"
    return texto[:LIMITE_TEXTO], paginas, None


# --------------------------------------------------------------------------- #
# Extração de metadados
# --------------------------------------------------------------------------- #


def _para_float(bruto: str) -> float | None:
    try:
        return float(bruto.replace(".", "").replace(",", "."))
    except ValueError:
        return None


def _janelas_de_ancora(texto: str, ancoras: Iterable[str]) -> list[str]:
    """Trechos em volta de cada âncora encontrada (antes e depois)."""
    trechos: list[str] = []
    for ancora in ancoras:
        inicio = 0
        while True:
            pos = texto.find(ancora, inicio)
            if pos < 0:
                break
            trechos.append(texto[max(0, pos - JANELA_ANTES): pos + len(ancora) + JANELA_DEPOIS])
            inicio = pos + len(ancora)
    return trechos


def extrair_valor(texto_normalizado: str) -> float | None:
    """
    Maior valor monetário que aparece perto de uma âncora de total.

    Trabalha no texto já normalizado (minúsculo, sem acento, espaços colapsados).
    """
    candidatos: list[float] = []
    for trecho in _janelas_de_ancora(texto_normalizado, ANCORAS_VALOR):
        for achado in RE_DINHEIRO.findall(trecho):
            valor = _para_float(achado)
            if valor is not None and valor > 0:
                candidatos.append(valor)
    if candidatos:
        return max(candidatos)

    # Sem âncora: aceita o maior valor da página, mas só se houver poucos
    # candidatos (documento simples). Em tabelão isso vira chute e não ajuda.
    todos = [v for v in (_para_float(a) for a in RE_DINHEIRO.findall(texto_normalizado)) if v]
    if todos and len(todos) <= 6:
        return max(todos)
    return None


def extrair_vencimento(texto_normalizado: str) -> date | None:
    """Data ``dd/mm/aaaa`` mais próxima da palavra 'vencimento'."""
    melhor: tuple[int, date] | None = None
    for encontro in re.finditer(r"vencimento", texto_normalizado):
        centro = encontro.start()
        trecho_ini = max(0, centro - 140)
        trecho = texto_normalizado[trecho_ini: centro + 160]
        for data_encontrada in RE_DATA_BR.finditer(trecho):
            dia, mes, ano = (int(g) for g in data_encontrada.groups())
            try:
                quando = date(ano, mes, dia)
            except ValueError:
                continue
            posicao_absoluta = trecho_ini + data_encontrada.start()
            distancia = abs(posicao_absoluta - centro)
            if melhor is None or distancia < melhor[0]:
                melhor = (distancia, quando)
    return melhor[1] if melhor else None


def extrair_numero_documento(texto_normalizado: str) -> str | None:
    """Número da nota / do documento / da fatura, quando o layout deixa claro."""
    tentativas = (
        r"\bnota\s*(?:fiscal)?\s*[:\-]\s*(\d{3,12})\b",
        r"\bnota fiscal\s*n?[o°º.]?\s*(\d{3,12})\b",
        r"\bn[o°º]\s*(\d{4,12})\b\s*\d{2}[-/]\d{2}[-/]\d{4}",
        r"\bn[o°º]?\s*d[ao] (?:documento|nota|fatura)\s*[:\-]?\s*([a-z0-9][a-z0-9\-/]{2,19})",
        r"\bfatura\s*[:\-]\s*([a-z0-9][a-z0-9\-/]{2,19})",
        r"\bcodigo\s*[:\-]?\s*(\d{2,4}/\d{6,12})",
    )
    for expressao in tentativas:
        achado = re.search(expressao, texto_normalizado)
        if achado:
            bruto = achado.group(1).strip(" .-/")
            if bruto and not bruto.isspace():
                return bruto.upper()
    return None


def extrair_cnpjs(texto: str) -> list[str]:
    """Todos os CNPJs formatados que aparecem no texto, sem repetir."""
    vistos: list[str] = []
    for cnpj in RE_CNPJ.findall(texto):
        if cnpj not in vistos:
            vistos.append(cnpj)
    return vistos


# --------------------------------------------------------------------------- #
# Pontuação
# --------------------------------------------------------------------------- #


def _pontuar(
    alvo: str,
    sinais: Sequence[Sinal],
    peso_relativo: float = 1.0,
) -> tuple[dict[TipoDocumento, float], list[tuple[TipoDocumento, float, str]]]:
    """Roda o catálogo de sinais sobre `alvo` e acumula a pontuação por tipo."""
    placar: dict[TipoDocumento, float] = {}
    encontrados: list[tuple[TipoDocumento, float, str]] = []
    for sinal in sinais:
        if sinal.padrao.search(alvo):
            peso = sinal.peso * peso_relativo
            placar[sinal.tipo] = placar.get(sinal.tipo, 0.0) + peso
            encontrados.append((sinal.tipo, peso, sinal.descricao))
    return placar, encontrados


def _somar(destino: dict[TipoDocumento, float], origem: dict[TipoDocumento, float]) -> None:
    for tipo, peso in origem.items():
        destino[tipo] = destino.get(tipo, 0.0) + peso


def _confianca(maior: float, segundo: float) -> float:
    """
    Converte a pontuação em confiança de 0 a 1.

    Leva em conta a força do sinal (quanto somou) e a separação para o
    segundo colocado (o quanto a decisão foi disputada).
    """
    if maior <= 0:
        return 0.0
    forca = min(1.0, maior / 12.0)
    separacao = max(0.0, (maior - segundo) / maior)
    return round(min(0.98, 0.25 + 0.45 * forca + 0.30 * separacao), 2)


def _montar_motivo(
    tipo: TipoDocumento,
    encontrados: list[tuple[TipoDocumento, float, str]],
    limite: int = 3,
) -> str:
    """Frase curta com os sinais que mais pesaram para o tipo vencedor."""
    do_tipo = sorted(
        (e for e in encontrados if e[0] is tipo), key=lambda e: e[1], reverse=True
    )
    if not do_tipo:
        return "nenhum sinal reconhecido"
    partes = [descricao for _, _, descricao in do_tipo[:limite]]
    sobra = len(do_tipo) - len(partes)
    texto = "; ".join(partes)
    return f"{texto} (+{sobra} outro(s) sinal(is))" if sobra else texto


# --------------------------------------------------------------------------- #
# API pública
# --------------------------------------------------------------------------- #


def classificar(caminho: Path, *, contexto: str = "") -> Documento:
    """
    Decide o tipo de um arquivo coletado.

    Args:
        caminho: arquivo já baixado (PDF, XML ou outro).
        contexto: texto extra que ajuda a decidir — normalmente o assunto do
            e-mail de onde o anexo veio. Pesa menos que o nome do arquivo.

    Returns:
        Um `Documento` com `tipo`, `confianca` (0 a 1) e `motivo` legível.
        Nunca levanta exceção por causa do conteúdo do arquivo.
    """
    caminho = Path(caminho)
    documento = Documento(caminho=caminho)
    extensao = caminho.suffix.lower()

    # XML de NF-e não precisa ser aberto: a extensão já diz tudo.
    if extensao in EXTENSOES_XML:
        documento.tipo = TipoDocumento.XML_NFE
        documento.confianca = 0.95
        documento.motivo = "extensão .xml — XML da NF-e (arquivo não foi aberto)"
        return documento

    nome_normalizado = normalizar(caminho.stem)
    placar: dict[TipoDocumento, float] = {}
    encontrados: list[tuple[TipoDocumento, float, str]] = []

    # --- 1) nome do arquivo -------------------------------------------------
    placar_nome, achados_nome = _pontuar(nome_normalizado, SINAIS_NOME)
    _somar(placar, placar_nome)
    encontrados.extend(achados_nome)

    if RE_NOME_COMPETENCIA.search(nome_normalizado):
        placar[TipoDocumento.AUTORIZACAO] = placar.get(TipoDocumento.AUTORIZACAO, 0.0) + 2.5
        encontrados.append((
            TipoDocumento.AUTORIZACAO, 2.5,
            "nome do arquivo segue o padrão da autorização do mês (termina em MM-AAAA)",
        ))
    if RE_NOME_RELATORIO_MENSAL.match(nome_normalizado):
        placar[TipoDocumento.DEMONSTRATIVO] = placar.get(TipoDocumento.DEMONSTRATIVO, 0.0) + 2.5
        encontrados.append((
            TipoDocumento.DEMONSTRATIVO, 2.5,
            "nome do arquivo é 'numero_AAAA-MM' (relatório mensal da conta)",
        ))
    if re.search(r"\bfatura\b", nome_normalizado) and not placar:
        # "fatura X.pdf" sozinho não decide nada — empata boleto e demonstrativo
        # de propósito, para o texto ser quem escolhe.
        for tipo in (TipoDocumento.BOLETO, TipoDocumento.DEMONSTRATIVO):
            placar[tipo] = placar.get(tipo, 0.0) + 0.5
        encontrados.append((
            TipoDocumento.BOLETO, 0.5, "nome do arquivo contém 'fatura' (sinal fraco)",
        ))
        encontrados.append((
            TipoDocumento.DEMONSTRATIVO, 0.5, "nome do arquivo contém 'fatura' (sinal fraco)",
        ))

    # --- 2) contexto (assunto do e-mail, pasta de origem) -------------------
    if contexto:
        placar_ctx, achados_ctx = _pontuar(normalizar(contexto), SINAIS_NOME, peso_relativo=0.4)
        _somar(placar, placar_ctx)
        encontrados.extend(
            (t, p, f"{d} (via contexto: assunto do e-mail)") for t, p, d in achados_ctx
        )

    # --- 3) texto da 1ª página ----------------------------------------------
    texto_bruto = ""
    erro_leitura: str | None = None
    if extensao in EXTENSOES_PDF:
        texto_bruto, documento.paginas, erro_leitura = extrair_texto_primeira_pagina(caminho)
    else:
        erro_leitura = f"extensão {extensao or '(sem extensão)'} não é lida — decidido só pelo nome"

    texto_normalizado = normalizar(texto_bruto) if texto_bruto else ""

    if texto_normalizado:
        placar_texto, achados_texto = _pontuar(texto_normalizado, SINAIS_TEXTO)
        _somar(placar, placar_texto)
        encontrados.extend(achados_texto)

        if all(exp.search(texto_normalizado) for exp in _TRIO_BOLETO):
            placar[TipoDocumento.BOLETO] = placar.get(TipoDocumento.BOLETO, 0.0) + 2.5
            encontrados.append((
                TipoDocumento.BOLETO, 2.5,
                "texto tem 'Beneficiário' + 'Vencimento' + 'Nosso número' juntos",
            ))

        documento.valor = extrair_valor(texto_normalizado)
        documento.vencimento = extrair_vencimento(texto_normalizado)
        documento.numero_documento = extrair_numero_documento(texto_normalizado)

    # --- 4) veredito ---------------------------------------------------------
    if not placar:
        documento.tipo = TipoDocumento.DESCONHECIDO
        documento.confianca = 0.0
        documento.motivo = erro_leitura or "nenhum sinal de boleto, nota, demonstrativo ou autorização"
        log.info("não classificado: %s (%s)", caminho.name, documento.motivo)
        return documento

    ordenado = sorted(placar.items(), key=lambda item: item[1], reverse=True)
    tipo_vencedor, pontos = ordenado[0]
    pontos_segundo = ordenado[1][1] if len(ordenado) > 1 else 0.0

    documento.tipo = tipo_vencedor
    documento.confianca = _confianca(pontos, pontos_segundo)
    if extensao not in EXTENSOES_PDF and extensao not in EXTENSOES_XML:
        # Sem texto para confirmar, a confiança fica limitada.
        documento.confianca = min(documento.confianca, 0.50)

    motivo = _montar_motivo(tipo_vencedor, encontrados)
    if erro_leitura and extensao in EXTENSOES_PDF:
        motivo = f"{motivo} — atenção: {erro_leitura}"
    if len(ordenado) > 1 and pontos_segundo > 0:
        motivo = f"{motivo} [2º lugar: {ordenado[1][0].rotulo}, {pontos_segundo:.1f} pts]"
    documento.motivo = motivo

    log.debug(
        "%s -> %s (%.2f) | %s", caminho.name, tipo_vencedor.value, documento.confianca, motivo
    )
    return documento


def classificar_lote(caminhos: Iterable[Path], contexto: str = "") -> list[Documento]:
    """Classifica vários arquivos de uma vez, na ordem recebida."""
    return [classificar(Path(c), contexto=contexto) for c in caminhos]


# --------------------------------------------------------------------------- #
# Teste manual
# --------------------------------------------------------------------------- #

if __name__ == "__main__":
    import sys

    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")

    if len(sys.argv) < 2:
        print("uso: python -m automacao.classificador <pasta-ou-arquivo> [...]")
        raise SystemExit(2)

    alvos: list[Path] = []
    for argumento in sys.argv[1:]:
        p = Path(argumento)
        if p.is_dir():
            alvos.extend(sorted(x for x in p.iterdir() if x.is_file()))
        elif p.is_file():
            alvos.append(p)

    for documento in classificar_lote(alvos):
        valor = f"R$ {documento.valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") \
            if documento.valor else "-"
        venc = documento.vencimento.strftime("%d/%m/%Y") if documento.vencimento else "-"
        print(f"\n{documento.nome}")
        print(f"  tipo      : {documento.tipo.value} (confiança {documento.confianca:.2f})")
        print(f"  valor     : {valor}   vencimento: {venc}   nº doc: {documento.numero_documento or '-'}")
        print(f"  motivo    : {documento.motivo}")
