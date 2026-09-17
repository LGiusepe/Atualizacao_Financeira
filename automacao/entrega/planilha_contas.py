"""
Leitura e escrita controlada da planilha `CONTAS E ACESSOS.xlsx`.

Este módulo é a única porta de entrada do projeto para aquela planilha.

Leitura
    Sempre permitida, feita com openpyxl sobre o arquivo do OneDrive.
    Nada é aberto em modo de escrita, nada é salvo.

Escrita
    Só acontece com `confirmado=True`. Sem isso, as funções de escrita
    devolvem `Situacao.PENDENTE` com um `detalhes` que descreve exatamente
    o que SERIA alterado — o painel mostra isso ao usuário, que só então
    chama de novo confirmando. Exigência direta do usuário: nada muda na
    pasta de trabalho dele sem explicação prévia.

Por que a gravação usa Excel COM e não openpyxl
    Ver o comentário longo em `_MOTIVO_COM`, logo abaixo. Resumo: openpyxl
    descarta as partes `customXml/*` do pacote, que nesta planilha carregam
    a ligação com o content type e as colunas gerenciadas da biblioteca do
    SharePoint. Gravar com openpyxl devolveria ao OneDrive um arquivo com
    metadados corporativos amputados.
"""

from __future__ import annotations

import difflib
import gc
import logging
import re
import shutil
import tempfile
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterator

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from automacao.nucleo.config import Ambiente, ambiente
from automacao.nucleo.modelos import Competencia, Conta, Etapa, ResultadoEtapa, Situacao, normalizar

log = logging.getLogger("automacao.entrega.planilha_contas")


# --------------------------------------------------------------------------- #
# Por que Excel COM na escrita — investigação feita sobre o arquivo real
# --------------------------------------------------------------------------- #

_MOTIVO_COM = """\
Investigação do impacto real de salvar esta planilha com openpyxl
(feita comparando o pacote OOXML antes e depois de um load/save em cópia local):

1. Validações de dados: NÃO EXISTEM neste arquivo. Nenhuma aba tem
   <dataValidation>, e não há extensão x14 de validação no extLst. O aviso
   "Data Validation extension is not supported" não é emitido por esta
   planilha — a ida-e-volta com openpyxl 3.1.5 roda sem nenhum UserWarning.
   Ou seja: aqui não há validação a perder.

2. O que openpyxl REALMENTE descarta neste arquivo:
     - customXml/item1.xml, item2.xml, item3.xml (+ itemProps + _rels)
       ~13,6 KB de metadados do SharePoint: o binding do content type da
       biblioteca, o esquema das colunas gerenciadas (managed metadata) e
       os FormTemplates do DocumentLibraryForm.
     - xl/calcChain.xml (irrelevante, o Excel reconstrói)
     - xl/sharedStrings.xml (irrelevante, vira string inline)
     - mc:AlternateContent / x15ac:absPath do workbook.xml (dica de caminho
       absoluto usada na coautoria do SharePoint)

3. Preservado corretamente por openpyxl: fórmulas, hyperlinks, mesclagens,
   formatação condicional, larguras de coluna, autofiltro e bordas.

Conclusão: o risco não é validação de dados, é a perda das partes customXml.
O arquivo mora em uma biblioteca do SharePoint
(uma biblioteca do SharePoint da empresa) e é coautorado
(coauthVersionLast=47). Devolver ao OneDrive um pacote sem customXml pode
zerar as colunas de metadados do documento na biblioteca.

Somado a isso, escrever bytes direto numa pasta sincronizada enquanto o
arquivo pode estar aberto por outra pessoa gera conflito de sincronização.

Por isso a ESCRITA é feita pelo Excel COM (pywin32): o próprio Excel regrava
o pacote inteiro, respeita bloqueio de arquivo e coautoria, e não perde
nenhuma parte. A LEITURA continua com openpyxl, que é mais rápido, não abre
o Excel e não corre risco nenhum por ser somente leitura.
"""


# --------------------------------------------------------------------------- #
# Layout da planilha
# --------------------------------------------------------------------------- #

ABA_CONTAS_FIXAS = "CONTAS FIXAS"
ABA_ACESSOS = "ACESSOS E CONTATOS"

LINHA_TITULO = 1
LINHA_CABECALHO = 2
PRIMEIRA_LINHA_DADOS = 3

# Colunas A..G
COL_EMPRESA = 1
COL_FORNECEDOR = 2
COL_VENCIMENTO = 3
COL_GERAR_BOLETO = 4
COL_VALOR = 5          # na prática guarda a situação da conta
COL_CONTATO = 6
COL_OBSERVACAO = 7
ULTIMA_COLUNA = COL_OBSERVACAO

# Verde de "já enviado". O valor efetivo vem de settings.yaml (checklist.cor_enviado).
COR_ENVIADO_PADRAO = "FF92D050"

# Situações da coluna E que tiram a conta do ar.
SITUACOES_INATIVAS = frozenset({"descontinuado", "cancelado"})

# Abaixo disso o painel deve pedir confirmação humana do casamento.
LIMIAR_CASAMENTO = 0.80
# Abaixo disso nem consideramos casamento.
ESCORE_MINIMO = 0.55

# Blocos verticais de ACESSOS E CONTATOS: primeira linha de dados.
PRIMEIRA_LINHA_ACESSOS = 4

_ROTULOS_LOGIN = frozenset({"login", "usuario", "user", "e-mail", "email"})
_ROTULOS_SENHA = frozenset({"senha", "password", "pass"})
_ROTULOS_SITE = frozenset({"site", "url", "endereco", "link"})
_ROTULOS_OBS = frozenset({"observacao", "obs", "nota"})
_ROTULOS_CONHECIDOS = _ROTULOS_LOGIN | _ROTULOS_SENHA | _ROTULOS_SITE | _ROTULOS_OBS

# "=C15-10" -> dia de gerar boleto derivado do vencimento.
_RX_FORMULA_DIA = re.compile(r"^=\s*C(\d+)\s*([+-])\s*(\d+)\s*$", re.IGNORECASE)


# --------------------------------------------------------------------------- #
# Estruturas
# --------------------------------------------------------------------------- #


@dataclass
class LinhaConta:
    """Uma linha A..G de CONTAS FIXAS ou de uma aba Checklist-*."""

    linha: int
    empresa: str
    fornecedor: str
    vencimento: Any
    gerar_boleto: Any
    # Coluna E (VALOR). Na prática guarda o estado da conta: vazio,
    # DESCONTINUADO, Redirecionado, cancelado, cartão, aguardando renovação.
    situacao: str
    contato: str
    observacao: str
    verde: bool
    # -- extras (não fazem parte do contrato mínimo; têm padrão) ------------ #
    #: Estado da caixa "Enviou para Financeiro?". `None` = a aba não tem essa
    #: coluna (abas antigas), e aí o verde só pode vir do preenchimento fixo.
    marcado: bool | None = None
    #: Endereço da caixa, ex. "K12". Vazio quando a coluna não existe.
    celula_marcacao: str = ""
    aba: str = ""
    # Valor cru da coluna D antes de resolver a fórmula, quando havia fórmula.
    gerar_boleto_bruto: Any = None
    # Preenchidos por casar_conta().
    escore: float = 0.0
    motivo_casamento: str = ""

    @property
    def ativa(self) -> bool:
        """False quando a conta está DESCONTINUADO ou cancelado."""
        return normalizar(self.situacao) not in SITUACOES_INATIVAS

    @property
    def casamento_duvidoso(self) -> bool:
        """True quando o casamento com uma Conta ficou abaixo do limiar."""
        return 0.0 < self.escore < LIMIAR_CASAMENTO

    @property
    def intervalo(self) -> str:
        """Endereço A..G desta linha, ex.: 'A12:G12'."""
        return f"A{self.linha}:G{self.linha}"

    @property
    def valores(self) -> list[Any]:
        return [
            self.empresa,
            self.fornecedor,
            self.vencimento,
            self.gerar_boleto,
            self.situacao,
            self.contato,
            self.observacao,
        ]

    def resumo(self) -> str:
        return f"{self.aba or '?'}!{self.linha} {self.empresa} / {self.fornecedor}"


@dataclass
class Credencial:
    """
    Um bloco de credencial da aba ACESSOS E CONTATOS (ou do cofre).

    A senha NUNCA aparece em repr()/str(). Use `.senha` de propósito quando
    precisar do valor — assim ninguém vaza senha em log por descuido.
    """

    servico: str
    login: str = ""
    senha: str = ""
    site: str = ""
    linha: int = 0
    observacao: str = ""

    @property
    def tem_senha(self) -> bool:
        return bool(self.senha)

    def mascarada(self) -> str:
        return "***" if self.senha else ""

    def como_dict(self, *, com_senha: bool = False) -> dict:
        """Dicionário serializável. Sem senha por padrão."""
        return {
            "servico": self.servico,
            "login": self.login,
            "senha": self.senha if com_senha else self.mascarada(),
            "site": self.site,
            "linha": self.linha,
            "observacao": self.observacao,
        }

    def __repr__(self) -> str:
        return (
            f"Credencial(servico={self.servico!r}, login={self.login!r}, "
            f"senha={self.mascarada()!r}, site={self.site!r}, linha={self.linha!r})"
        )

    __str__ = __repr__


class ErroPlanilha(RuntimeError):
    """Problema ao ler ou gravar a planilha de contas."""


# --------------------------------------------------------------------------- #
# Utilidades de leitura
# --------------------------------------------------------------------------- #


def caminho_planilha(ambiente_: Ambiente | None = None) -> Path:
    return (ambiente_ or ambiente()).caminhos.planilha_contas


def cor_enviado(ambiente_: Ambiente | None = None) -> str:
    bruto = (ambiente_ or ambiente()).checklist.get("cor_enviado") or COR_ENVIADO_PADRAO
    return _normalizar_cor(str(bruto))


def _normalizar_cor(rgb: str) -> str:
    """'92D050', '#92D050' e 'FF92D050' viram todos 'FF92D050'."""
    limpo = rgb.strip().lstrip("#").upper()
    return limpo if len(limpo) == 8 else f"FF{limpo[-6:]}"


def _texto(valor: Any) -> str:
    return "" if valor is None else str(valor).strip()


def _cor_de_fundo(celula) -> str | None:
    """Cor sólida RGB da célula, ou None (tema, indexada ou sem preenchimento)."""
    preenchimento = celula.fill
    if preenchimento is None or preenchimento.fill_type != "solid":
        return None
    fundo = preenchimento.fgColor
    if fundo is None or fundo.type != "rgb":
        return None
    valor = fundo.rgb
    return valor.upper() if isinstance(valor, str) else None


def _linha_esta_verde(
    ws: Worksheet, linha: int, cor: str, ultima_coluna: int = ULTIMA_COLUNA
) -> bool:
    """
    Verde por preenchimento fixo — só para as abas sem caixa de seleção.

    A largura vem do cabeçalho, não é fixa: a aba de agosto tem colunas a mais
    que as anteriores, e comparar sempre A..G daria falso negativo.

    Onde existe a coluna "Enviou para Financeiro?" quem responde é
    `_esta_marcada`: ali o verde é formatação condicional e não deixa
    preenchimento nenhum nas células.
    """
    return all(
        _cor_de_fundo(ws.cell(linha, coluna)) == cor
        for coluna in range(COL_EMPRESA, ultima_coluna + 1)
    )


def _esta_marcada(bruto: Any) -> bool:
    """
    Se a caixa "Enviou para Financeiro?" está marcada.

    O Excel guarda a caixa como booleano na própria célula, mas a mesma coluna
    já foi preenchida à mão em meses anteriores — daí aceitar também texto.
    """
    if isinstance(bruto, bool):
        return bruto
    return normalizar(_texto(bruto)) in {"true", "verdadeiro", "sim", "x", "1"}


def _resolver_dia_boleto(ws: Worksheet, bruto: Any) -> Any:
    """
    Resolve a fórmula trivial '=C15-10' usando o vencimento da própria linha.

    A planilha é lida sem data_only, então a coluna D vem como texto de
    fórmula. O painel precisa do número; a conta é simples o bastante para
    resolver aqui, sem abrir o Excel.
    """
    if not isinstance(bruto, str):
        return bruto
    casamento = _RX_FORMULA_DIA.match(bruto.strip())
    if casamento is None:
        return bruto
    linha_ref, sinal, quantidade = casamento.groups()
    vencimento = ws.cell(int(linha_ref), COL_VENCIMENTO).value
    if not isinstance(vencimento, (int, float)):
        return bruto
    delta = int(quantidade)
    return int(vencimento) - delta if sinal == "-" else int(vencimento) + delta


# Cache leve das leituras, invalidado por mtime/tamanho do arquivo.
_cache_leitura: dict[str, tuple[tuple, Any]] = {}


def _assinatura(caminho: Path) -> tuple:
    estado = caminho.stat()
    return (str(caminho), estado.st_mtime_ns, estado.st_size)


def _memorizado(chave: str, caminho: Path, produtor):
    assinatura = _assinatura(caminho)
    guardado = _cache_leitura.get(chave)
    if guardado is not None and guardado[0] == assinatura:
        return guardado[1]
    resultado = produtor()
    _cache_leitura[chave] = (assinatura, resultado)
    return resultado


def limpar_cache() -> None:
    """Descarta as leituras memorizadas — usar depois de gravar."""
    _cache_leitura.clear()


@contextmanager
def _planilha_aberta(caminho: Path) -> Iterator[openpyxl.Workbook]:
    """
    Abre a planilha para leitura, mesmo com o Excel segurando o arquivo.

    Sem data_only: a planilha tem fórmulas e não queremos trabalhar com o
    cache de valores. keep_vba=False porque o arquivo é .xlsx puro.

    Quando a planilha está aberta no Excel, o Windows recusa a leitura direta
    (`PermissionError`) — mas `shutil.copy2` continua funcionando, porque usa a
    API de cópia do sistema. Nesse caso lemos uma cópia temporária. É leitura,
    então trabalhar sobre a cópia não muda nada no resultado; só evita que a
    automação pare de funcionar sempre que você estiver com a planilha aberta.
    """
    if not caminho.is_file():
        raise ErroPlanilha(f"planilha não encontrada: {caminho}")

    try:
        wb = openpyxl.load_workbook(caminho, keep_vba=False, data_only=False)
    except PermissionError:
        log.info(
            "%s está aberta em outro programa — lendo uma cópia temporária.",
            caminho.name,
        )
        with tempfile.TemporaryDirectory(prefix="automacao_planilha_") as pasta:
            copia = Path(pasta) / caminho.name
            try:
                shutil.copy2(caminho, copia)
            except OSError as erro:
                raise ErroPlanilha(
                    f"não consegui ler {caminho.name}: o arquivo está bloqueado e "
                    f"a cópia também falhou ({erro}). Feche a planilha no Excel "
                    "e tente de novo."
                ) from erro
            wb = openpyxl.load_workbook(copia, keep_vba=False, data_only=False)
            try:
                yield wb
            finally:
                wb.close()
            return

    try:
        yield wb
    finally:
        wb.close()


# Nomes que cada coluna já teve. A planilha é editada à mão e o layout muda:
# hoje convivem três formatos — CONTAS FIXAS e Março–Junho com título na linha 1
# e cabeçalho na 2; Julho com cabeçalho na linha 1 e FORNECEDOR na coluna B;
# Agosto com cabeçalho na linha 1, FORNECEDOR na D e colunas novas de
# SERVIÇO/PRODUTO, CENTRO DE CUSTO e Contrato / conta.
# Por isso a coluna é achada pelo NOME, nunca pela posição.
APELIDOS_COLUNA: dict[str, tuple[str, ...]] = {
    "empresa": ("empresa",),
    "fornecedor": ("fornecedor",),
    "servico": ("servico produto", "servico", "produto"),
    "centro_custo": ("centro de custo", "centro custo"),
    "contrato": ("contrato conta", "contrato", "conta"),
    "vencimento": ("vencimento",),
    "gerar_boleto": ("gerar boleto", "gerar"),
    "situacao": ("valor", "valor medio ou estimado", "situacao"),
    "contato": ("contato",),
    "observacao": ("observacao", "observacoes", "obs"),
    # "Enviou para Financeiro?\nSIM ou NÃO" — a caixa de seleção que dispara a
    # formatação condicional da aba. É ela que pinta a linha; o preenchimento
    # fixo que a automação usava antes era um verde paralelo, que deixava a
    # caixa desmarcada e a coluna dizendo "não enviado".
    "enviado": ("enviou para financeiro", "enviou financeiro", "enviou"),
}


def _titulo_normalizado(bruto: Any) -> str:
    """
    Título de coluna pronto para comparar: sem acento, sem pontuação, minúsculo.

    'SERVIÇO/PRODUTO'            -> 'servico produto'
    'VALOR \\n(médio ou estimado)' -> 'valor medio ou estimado'
    """
    limpo = normalizar(_texto(bruto))
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9]+", " ", limpo)).strip()


def _casa_titulo(titulo: str, apelidos: tuple[str, ...]) -> bool:
    """
    Compara título com apelidos por PALAVRA INTEIRA.

    Substring solta não serve: 'conta' casava com 'contato' e a coluna de
    contato virava número de contrato, corrompendo a leitura de julho.
    """
    return any(
        titulo == apelido or titulo.startswith(apelido + " ")
        for apelido in apelidos
    )

LINHAS_CANDIDATAS_CABECALHO = (1, 2, 3)


def _mapear_colunas(ws: Worksheet) -> tuple[int, dict[str, int]]:
    """
    Descobre em que linha está o cabeçalho e em que coluna mora cada campo.

    Devolve `(linha_do_cabecalho, {campo: numero_da_coluna})`. Campo ausente
    simplesmente não aparece no dicionário — quem lê usa `.get()`.
    """
    for linha in LINHAS_CANDIDATAS_CABECALHO:
        if linha > ws.max_row:
            break
        titulos = {
            coluna: _titulo_normalizado(ws.cell(linha, coluna).value)
            for coluna in range(1, min(ws.max_column, 16) + 1)
        }
        if not any(t in ("empresa", "fornecedor") for t in titulos.values()):
            continue

        mapa: dict[str, int] = {}
        for campo, apelidos in APELIDOS_COLUNA.items():
            for coluna, titulo in titulos.items():
                if not titulo or campo in mapa:
                    continue
                if _casa_titulo(titulo, apelidos):
                    mapa[campo] = coluna
                    break
        if "fornecedor" in mapa:
            return linha, mapa

    # Nenhum cabeçalho reconhecido: cai no layout histórico.
    log.warning(
        "não reconheci o cabeçalho da aba %r — usando o layout antigo A..G", ws.title
    )
    return LINHA_CABECALHO, {
        "empresa": COL_EMPRESA,
        "fornecedor": COL_FORNECEDOR,
        "vencimento": COL_VENCIMENTO,
        "gerar_boleto": COL_GERAR_BOLETO,
        "situacao": COL_VALOR,
        "contato": COL_CONTATO,
        "observacao": COL_OBSERVACAO,
    }


def _ler_aba(ws: Worksheet, cor: str) -> list[LinhaConta]:
    """Lê as linhas de dados da aba, seja qual for o layout de colunas."""

    def celula(numero: int, campo: str):
        coluna = mapa.get(campo)
        return ws.cell(numero, coluna).value if coluna else None

    cabecalho, mapa = _mapear_colunas(ws)
    ultima_coluna = max(mapa.values())

    linhas: list[LinhaConta] = []
    for numero in range(cabecalho + 1, ws.max_row + 1):
        empresa = _texto(celula(numero, "empresa"))
        fornecedor = _texto(celula(numero, "fornecedor"))
        if not empresa and not fornecedor:
            continue  # linha só formatada, sem conteúdo

        bruto_marcacao = celula(numero, "enviado")
        coluna_marcacao = mapa.get("enviado")
        bruto_boleto = celula(numero, "gerar_boleto")
        resolvido = _resolver_dia_boleto(ws, bruto_boleto)

        # A coluna nova "Contrato / conta" é o que finalmente distingue as
        # contas irmãs de um mesmo fornecedor — vai junto da observação.
        contrato = _texto(celula(numero, "contrato"))
        servico = _texto(celula(numero, "servico"))
        observacao = _texto(celula(numero, "observacao"))
        extras = " | ".join(p for p in (servico, contrato, observacao) if p)

        linhas.append(
            LinhaConta(
                linha=numero,
                empresa=empresa,
                fornecedor=f"{fornecedor} - {contrato}" if contrato else fornecedor,
                vencimento=celula(numero, "vencimento"),
                gerar_boleto=resolvido,
                situacao=_texto(celula(numero, "situacao")),
                contato=_texto(celula(numero, "contato")),
                observacao=extras,
                # A caixa marcada É o verde: a formatação condicional da aba
                # (A2:K195, fórmula `$K2`) pinta a linha a partir dela. Só onde
                # a coluna não existe é que sobra olhar o preenchimento.
                verde=(
                    _esta_marcada(bruto_marcacao)
                    if coluna_marcacao
                    else _linha_esta_verde(ws, numero, cor, ultima_coluna)
                ),
                marcado=_esta_marcada(bruto_marcacao) if coluna_marcacao else None,
                celula_marcacao=(
                    f"{get_column_letter(coluna_marcacao)}{numero}"
                    if coluna_marcacao
                    else ""
                ),
                aba=ws.title,
                gerar_boleto_bruto=bruto_boleto if resolvido != bruto_boleto else None,
            )
        )
    return linhas


# --------------------------------------------------------------------------- #
# Leitura pública
# --------------------------------------------------------------------------- #


def ler_contas_fixas(*, caminho: Path | None = None) -> list[LinhaConta]:
    """Todas as linhas da aba CONTAS FIXAS (a lista mestra)."""
    alvo = caminho or caminho_planilha()
    cor = cor_enviado()

    def produzir() -> list[LinhaConta]:
        with _planilha_aberta(alvo) as wb:
            if ABA_CONTAS_FIXAS not in wb.sheetnames:
                raise ErroPlanilha(
                    f"aba {ABA_CONTAS_FIXAS!r} não existe em {alvo.name} "
                    f"(abas: {wb.sheetnames})"
                )
            return _ler_aba(wb[ABA_CONTAS_FIXAS], cor)

    linhas = _memorizado(f"fixas::{alvo}", alvo, produzir)
    log.debug("CONTAS FIXAS: %d linhas", len(linhas))
    return list(linhas)


def ler_checklist(
    competencia: Competencia, *, caminho: Path | None = None
) -> list[LinhaConta]:
    """
    Linhas da aba do mês (`competencia.aba_checklist`).

    Se a aba não existir devolve lista vazia — quem chamou decide se cria
    a aba, se pula o mês ou se avisa o usuário.
    """
    alvo = caminho or caminho_planilha()
    aba = competencia.aba_checklist
    cor = cor_enviado()

    def produzir() -> list[LinhaConta]:
        with _planilha_aberta(alvo) as wb:
            if aba not in wb.sheetnames:
                log.info("aba %r ainda não existe em %s", aba, alvo.name)
                return []
            return _ler_aba(wb[aba], cor)

    linhas = _memorizado(f"checklist::{aba}::{alvo}", alvo, produzir)
    log.debug("%s: %d linhas (%d verdes)", aba, len(linhas), sum(l.verde for l in linhas))
    return list(linhas)


def abas_checklist_existentes(*, caminho: Path | None = None) -> list[str]:
    """Nomes das abas Checklist-*, na ordem em que estão na pasta de trabalho."""
    alvo = caminho or caminho_planilha()
    prefixo = ambiente().checklist.get("prefixo_aba") or "Checklist-"

    def produzir() -> list[str]:
        with _planilha_aberta(alvo) as wb:
            return [n for n in wb.sheetnames if n.startswith(prefixo)]

    return list(_memorizado(f"abas::{alvo}", alvo, produzir))


def ler_acessos(*, caminho: Path | None = None) -> list[Credencial]:
    """
    Faz o parser dos blocos verticais de credencial da aba ACESSOS E CONTATOS.

    Formato de cada bloco (a partir da linha 4), separados por linha em branco:

        A: <nome do serviço>          (B vazio)
        A: Login | USUÁRIO   B: <valor>
        A: Senha | SENHA     B: <valor>
        A: SITE              B: <valor>

    Os rótulos variam de bloco para bloco (`Login`/`USUÁRIO`, `Senha`/` Senha`/
    `SENHA`), por isso a comparação é feita com `normalizar()`. Só as colunas
    A e B são lidas — o bloco de contatos em I..K e os CNPJs em H..I ficam
    naturalmente de fora.
    """
    alvo = caminho or caminho_planilha()

    def produzir() -> list[Credencial]:
        with _planilha_aberta(alvo) as wb:
            if ABA_ACESSOS not in wb.sheetnames:
                log.warning("aba %r não encontrada em %s", ABA_ACESSOS, alvo.name)
                return []
            ws = wb[ABA_ACESSOS]
            return _parsear_blocos_acesso(ws)

    credenciais = _memorizado(f"acessos::{alvo}", alvo, produzir)
    log.debug("ACESSOS E CONTATOS: %d credenciais", len(credenciais))
    return list(credenciais)


def _parsear_blocos_acesso(ws: Worksheet) -> list[Credencial]:
    credenciais: list[Credencial] = []
    atual: Credencial | None = None

    def fechar() -> None:
        nonlocal atual
        if atual is not None and (atual.login or atual.senha or atual.site):
            credenciais.append(atual)
        atual = None

    for numero in range(PRIMEIRA_LINHA_ACESSOS, ws.max_row + 1):
        rotulo = _texto(ws.cell(numero, 1).value)
        valor = _texto(ws.cell(numero, 2).value)

        if not rotulo:
            # Linha em branco na coluna A separa os blocos.
            if not valor:
                fechar()
            continue

        chave = normalizar(rotulo)
        if chave not in _ROTULOS_CONHECIDOS:
            # Cabeçalho de bloco novo: fecha o anterior e começa este.
            fechar()
            atual = Credencial(servico=_limpar_nome_servico(rotulo), linha=numero)
            if valor:
                atual.observacao = valor
            continue

        if atual is None:
            log.warning(
                "%s!A%d: rótulo %r sem serviço acima — ignorado", ws.title, numero, rotulo
            )
            continue

        if chave in _ROTULOS_LOGIN:
            atual.login = valor
        elif chave in _ROTULOS_SENHA:
            atual.senha = valor
        elif chave in _ROTULOS_SITE:
            atual.site = valor
        else:
            atual.observacao = f"{atual.observacao} {valor}".strip()

    fechar()
    return credenciais


def _limpar_nome_servico(bruto: str) -> str:
    """'Operadora A & B:  ' -> 'Operadora A & B'."""
    return re.sub(r"\s+", " ", bruto).strip().rstrip(":").strip()


# --------------------------------------------------------------------------- #
# Casamento Conta <-> linha da planilha
# --------------------------------------------------------------------------- #


def _similaridade(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    return difflib.SequenceMatcher(None, a, b).ratio()


def casar_conta(conta: Conta, linhas: list[LinhaConta]) -> LinhaConta | None:
    """
    Encontra a linha da planilha que corresponde a `conta`.

    Estratégia, em ordem:
      1. `conta.chaves_planilha` — texto exato da coluna FORNECEDOR
         (normalizado). Escore 1.0.
      2. `conta.chaves_planilha` contida no FORNECEDOR (ou vice-versa).
         Escore 0.95.
      3. Similaridade normalizada (difflib) entre `conta.rotulo`,
         `conta.pasta` e `conta.subunidade` e o campo FORNECEDOR.

    A linha devolvida traz `.escore` e `.motivo_casamento` preenchidos, e
    `.casamento_duvidoso` fica True abaixo de 0.80 — é isso que o painel usa
    para pedir conferência humana. Devolve None se nada passar de 0.55.
    """
    if not linhas:
        return None

    melhor: LinhaConta | None = None
    melhor_escore = 0.0
    melhor_motivo = ""

    chaves = [normalizar(c) for c in conta.chaves_planilha if _texto(c)]
    candidatos = [
        ("rotulo", normalizar(conta.rotulo)),
        ("pasta", normalizar(conta.pasta)),
        ("subunidade", normalizar(conta.subunidade or "")),
    ]

    for linha in linhas:
        fornecedor = normalizar(linha.fornecedor)
        if not fornecedor:
            continue

        escore = 0.0
        motivo = ""

        for chave in chaves:
            if chave == fornecedor:
                escore, motivo = 1.0, f"chave_planilha exata: {chave!r}"
                break
            if chave in fornecedor or fornecedor in chave:
                if 0.95 > escore:
                    escore, motivo = 0.95, f"chave_planilha contida: {chave!r}"

        if escore < 1.0:
            for nome, texto in candidatos:
                if not texto:
                    continue
                razao = _similaridade(texto, fornecedor)
                if razao > escore:
                    escore, motivo = razao, f"similaridade {nome} ({razao:.2f})"

        if escore > melhor_escore:
            melhor, melhor_escore, melhor_motivo = linha, escore, motivo

    if melhor is None or melhor_escore < ESCORE_MINIMO:
        log.info(
            "conta %r sem casamento na planilha (melhor escore %.2f)",
            conta.id,
            melhor_escore,
        )
        return None

    melhor.escore = round(melhor_escore, 4)
    melhor.motivo_casamento = melhor_motivo
    if melhor.casamento_duvidoso:
        log.warning(
            "casamento duvidoso de %r com %s (escore %.2f, %s)",
            conta.id,
            melhor.resumo(),
            melhor_escore,
            melhor_motivo,
        )
    return melhor


# --------------------------------------------------------------------------- #
# Escrita — travas comuns
# --------------------------------------------------------------------------- #


def _fazer_backup(caminho: Path, amb: Ambiente) -> Path | None:
    """Copia a planilha para dados/backups/ antes de qualquer gravação."""
    if not amb.seguranca.backup_antes_de_sobrescrever:
        log.info("backup desativado em settings.yaml — seguindo sem cópia")
        return None
    amb.caminhos.backups.mkdir(parents=True, exist_ok=True)
    carimbo = datetime.now().strftime("%Y%m%d-%H%M%S")
    destino = amb.caminhos.backups / f"{caminho.stem}_{carimbo}{caminho.suffix}"
    shutil.copy2(caminho, destino)
    log.info("backup criado: %s", destino)
    return destino


def _nome_backup_previsto(caminho: Path, amb: Ambiente) -> str | None:
    if not amb.seguranca.backup_antes_de_sobrescrever:
        return None
    carimbo = datetime.now().strftime("%Y%m%d-%H%M%S")
    return str(amb.caminhos.backups / f"{caminho.stem}_{carimbo}{caminho.suffix}")


def _diagnostico_permissao(caminho: Path, amb: Ambiente) -> dict:
    """Informa (sem levantar exceção) se a gravação seria liberada agora."""
    permitido, motivo = amb.pode_gravar(caminho)
    return {
        "permitido": permitido,
        "bloqueio": motivo or None,
        "modo_simulacao": amb.simulando,
        "confirmar_antes_de_gravar": amb.seguranca.confirmar_antes_de_gravar,
    }


# --------------------------------------------------------------------------- #
# Escrita — motor Excel COM
# --------------------------------------------------------------------------- #

_XL_SOLIDO = 1
_XL_SEM_PADRAO = -4142


def _cor_ole(rgb: str) -> int:
    """'FF92D050' -> inteiro BGR que o Excel COM entende."""
    hexa = _normalizar_cor(rgb)[-6:]
    vermelho, verde, azul = int(hexa[0:2], 16), int(hexa[2:4], 16), int(hexa[4:6], 16)
    return azul * 0x010000 + verde * 0x000100 + vermelho


@contextmanager
def _excel_com(caminho: Path) -> Iterator[Any]:
    """
    Abre a pasta de trabalho pelo Excel e salva ao final.

    Se o arquivo já estiver aberto numa instância do Excel do usuário,
    reaproveita essa instância (não fecha nada dele). Caso contrário sobe uma
    instância própria, invisível, e a encerra no final.

    ATENÇÃO ao mexer aqui: o pywin32 fala com o Excel por late binding, e nesse
    modo os ARGUMENTOS NOMEADOS SÃO SILENCIOSAMENTE IGNORADOS. `Copy(After=x)`
    não copia para o lugar pedido — o Excel entende `Copy()` sem argumento e
    joga a cópia numa PASTA DE TRABALHO NOVA, sem erro nenhum. Por isso todas
    as chamadas COM deste módulo usam argumentos POSICIONAIS.
    """
    try:
        import pythoncom
        import win32com.client as com
    except ImportError as exc:  # pragma: no cover
        raise ErroPlanilha(
            "pywin32 não disponível — a gravação exige Excel COM. "
            f"Detalhe: {exc}"
        ) from exc

    pythoncom.CoInitialize()
    app = None
    wb = None
    nossa_instancia = False
    pid_nosso: int | None = None
    alertas_antes = None
    try:
        alvo = str(caminho.resolve())

        # 1) Excel já aberto com este arquivo? Toda esta sondagem é best effort:
        #    o objeto devolvido pode ser uma instância morta que ficou no ROT,
        #    e aí qualquer atributo levanta AttributeError. Se der qualquer
        #    problema, seguimos para a instância própria.
        try:
            existente = com.GetActiveObject("Excel.Application")
            for aberto in existente.Workbooks:
                if str(aberto.FullName).lower() == alvo.lower():
                    app, wb = existente, aberto
                    log.info("reaproveitando o Excel já aberto com a planilha")
                    break
        except Exception as exc:
            log.debug("nenhum Excel reaproveitável (%s: %s)", type(exc).__name__, exc)
            app = wb = None

        # 2) Instância própria. Open(Filename, UpdateLinks, ReadOnly) posicional.
        if wb is None:
            app = com.DispatchEx("Excel.Application")
            nossa_instancia = True
            app.Visible = False
            app.AskToUpdateLinks = False
            # Antes de abrir qualquer coisa: guarda o PID. É o único jeito de
            # conferir depois se o processo realmente morreu — `DispatchEx`
            # não registra a instância no ROT, então nem dá para reencontrá-la
            # pelo COM se ela escapar.
            pid_nosso = _pid_do_excel(app)
            wb = app.Workbooks.Open(alvo, 0, False)

        alertas_antes = app.DisplayAlerts
        app.DisplayAlerts = False
        app.ScreenUpdating = False

        yield wb

        wb.Save()
        log.info("pasta de trabalho salva pelo Excel: %s", caminho.name)
    finally:
        try:
            if app is not None:
                app.ScreenUpdating = True
                if alertas_antes is not None:
                    app.DisplayAlerts = alertas_antes
            if nossa_instancia:
                if wb is not None:
                    wb.Close(False)  # posicional: SaveChanges
                if app is not None:
                    app.Quit()
        except Exception:  # pragma: no cover - limpeza best effort
            log.exception("falha ao encerrar o Excel")
        finally:
            # Soltar as referências É PARTE do encerramento, não faxina: o
            # Excel só sai quando a última some. Sem estas três linhas o
            # `Quit()` acima vira pedido ignorado e o processo fica segurando
            # a planilha — que foi exatamente o que aconteceu.
            wb = None
            app = None
            gc.collect()
            pythoncom.CoUninitialize()
            if nossa_instancia and pid_nosso:
                _garantir_morte(pid_nosso, caminho)


def _pid_do_excel(app: Any) -> int | None:
    """PID da instância do Excel, pela janela dela (existe mesmo invisível)."""
    try:
        import win32process

        return int(win32process.GetWindowThreadProcessId(int(app.Hwnd))[1])
    except Exception as exc:  # noqa: BLE001 — sem PID seguimos sem a rede
        log.debug("não consegui descobrir o PID do Excel: %s", exc)
        return None


def _garantir_morte(pid: int, caminho: Path, espera: float = 3.0) -> None:
    """
    Confere se o Excel que abrimos morreu; se sobreviveu ao `Quit()`, encerra.

    Rede de segurança, não o caminho normal: se cair aqui, alguma referência
    COM escapou e o registro no log serve para achar qual. Só encerramos o
    processo que NÓS criamos, pelo PID guardado na abertura — nunca um Excel
    do usuário, que pode ter trabalho não salvo.
    """
    import time

    fim = time.monotonic() + espera
    while time.monotonic() < fim:
        if not _processo_vivo(pid):
            return
        time.sleep(0.25)

    log.warning(
        "o Excel que abrimos (PID %s) não encerrou sozinho e estava segurando "
        "%s — encerrando à força para não travar a planilha para você.",
        pid, caminho.name,
    )
    try:
        import win32api
        import win32con

        alca = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, pid)
        win32api.TerminateProcess(alca, 0)
        win32api.CloseHandle(alca)
        log.info("PID %s encerrado; a planilha voltou a ficar livre", pid)
    except Exception as exc:  # noqa: BLE001
        log.error(
            "não consegui encerrar o PID %s (%s). Feche o EXCEL.EXE sem janela "
            "pelo Gerenciador de Tarefas para liberar %s.",
            pid, exc, caminho.name,
        )


def _processo_vivo(pid: int) -> bool:
    try:
        import win32api
        import win32con
        import win32event

        alca = win32api.OpenProcess(win32con.SYNCHRONIZE, False, pid)
        vivo = win32event.WaitForSingleObject(alca, 0) != win32event.WAIT_OBJECT_0
        win32api.CloseHandle(alca)
        return vivo
    except Exception:  # noqa: BLE001 — processo inexistente cai aqui
        return False


def _pintar_linha(ws_com: Any, linha: int, cor: str) -> None:
    """Preenchimento fixo em A..G. Só para as abas sem caixa de seleção."""
    faixa = ws_com.Range(f"A{linha}:G{linha}")
    faixa.Interior.Pattern = _XL_SOLIDO
    faixa.Interior.Color = _cor_ole(cor)
    faixa.Interior.TintAndShade = 0
    faixa.Interior.PatternTintAndShade = 0


def _marcar_caixa(ws_com: Any, celula: str, marcada: bool = True) -> None:
    """
    Marca (ou desmarca) a caixa "Enviou para Financeiro?".

    É tudo o que precisa ser feito: a aba tem formatação condicional em
    `A2:K195` com a fórmula `$K2`, então a linha inteira fica verde a partir
    daqui — inclusive H..K, que o preenchimento fixo nem alcançava.

    Escrever a caixa em vez de pintar também conserta um dado errado: antes a
    linha ficava verde com a coluna ainda dizendo "não enviado", e quem lesse
    a planilha sem olhar a cor concluía o contrário do que aconteceu.
    """
    ws_com.Range(celula).Value = bool(marcada)


def _duplicar_aba(wb_com: Any, nome_base: str, depois_de: str, nome_novo: str) -> Any:
    """
    Duplica `nome_base` logo depois de `depois_de` e renomeia para `nome_novo`.

    Usa `Copy(Before, After)` POSICIONAL — a forma nomeada é ignorada no late
    binding e a cópia acabaria numa pasta de trabalho nova (ver `_excel_com`).
    A aba criada é localizada por diferença de nomes, nunca por índice, e a
    função falha alto se a cópia não apareceu nesta pasta de trabalho.
    """
    antes = {f.Name for f in wb_com.Sheets}
    wb_com.Sheets(nome_base).Copy(None, wb_com.Sheets(depois_de))
    novas = {f.Name for f in wb_com.Sheets} - antes

    if len(novas) != 1:
        raise ErroPlanilha(
            f"a cópia de {nome_base!r} não apareceu nesta pasta de trabalho "
            f"(abas novas: {sorted(novas)}). Nada foi renomeado."
        )

    nova = wb_com.Sheets(novas.pop())
    nova.Name = nome_novo
    return nova


def _conferir_cabecalho(ws_com: Any, esperado: list[str]) -> None:
    """Confirma que a aba criada é mesmo cópia do layout A..G esperado."""
    obtido = [
        _texto(ws_com.Cells(LINHA_CABECALHO, coluna).Value)
        for coluna in range(COL_EMPRESA, ULTIMA_COLUNA + 1)
    ]
    if obtido != esperado:
        raise ErroPlanilha(
            f"a aba criada não tem o cabeçalho da base — esperado {esperado}, "
            f"obtido {obtido}"
        )


# --------------------------------------------------------------------------- #
# Escrita pública
# --------------------------------------------------------------------------- #


def marcar_enviado(
    conta: Conta,
    competencia: Competencia,
    *,
    ambiente_: Ambiente | None = None,
    confirmado: bool = False,
) -> ResultadoEtapa:
    """
    Pinta de verde a linha da conta na aba do mês.

    Com `confirmado=False` (padrão) NADA é gravado: devolve `Situacao.PENDENTE`
    e `detalhes` com a aba, o número da linha, o conteúdo atual da linha e a
    cor que seria aplicada. Só depois de o usuário ver isso no painel é que
    vale chamar de novo com `confirmado=True`.
    """
    amb = ambiente_ or ambiente()
    caminho = amb.caminhos.planilha_contas
    aba = competencia.aba_checklist
    cor = cor_enviado(amb)

    linhas = ler_checklist(competencia, caminho=caminho)
    if not linhas:
        return ResultadoEtapa.erro(
            Etapa.CHECKLIST,
            f"a aba {aba!r} não existe na planilha — crie o mês antes (criar_aba_mes)",
            detalhes={"aba": aba, "arquivo": str(caminho)},
        )

    alvo = casar_conta(conta, linhas)
    if alvo is None:
        return ResultadoEtapa.erro(
            Etapa.CHECKLIST,
            f"conta {conta.id!r} não encontrada em {aba!r}",
            detalhes={
                "aba": aba,
                "arquivo": str(caminho),
                "chaves_planilha": list(conta.chaves_planilha),
                "rotulo": conta.rotulo,
            },
        )

    # Aba com caixa de seleção: marcar a caixa é o mecanismo oficial e a
    # formatação condicional faz o resto. Sem a coluna (meses antigos), sobra
    # o preenchimento fixo em A..G.
    por_caixa = bool(alvo.celula_marcacao)
    onde = f"{aba}!{alvo.celula_marcacao}" if por_caixa else f"{aba}!{alvo.intervalo}"

    detalhes = {
        "acao": "marcar_caixa" if por_caixa else "pintar_linha",
        "arquivo": str(caminho),
        "aba": aba,
        "linha": alvo.linha,
        "intervalo": alvo.intervalo,
        "celula_marcacao": alvo.celula_marcacao,
        "conteudo": {
            "EMPRESA": alvo.empresa,
            "FORNECEDOR": alvo.fornecedor,
            "VENCIMENTO": alvo.vencimento,
            "GERAR BOLETO": alvo.gerar_boleto,
            "VALOR": alvo.situacao,
            "CONTATO": alvo.contato,
            "OBSERVAÇÃO": alvo.observacao,
        },
        "cor": cor,
        "escore_casamento": alvo.escore,
        "motivo_casamento": alvo.motivo_casamento,
        "casamento_duvidoso": alvo.casamento_duvidoso,
        "motor": "excel_com",
        "backup_previsto": _nome_backup_previsto(caminho, amb),
        **_diagnostico_permissao(caminho, amb),
    }

    if alvo.verde:
        return ResultadoEtapa.pulado(
            Etapa.CHECKLIST,
            (f"{onde} já está marcada — nada a fazer" if por_caixa
             else f"{onde} já está verde — nada a fazer"),
            detalhes=detalhes,
        )

    if not confirmado:
        return ResultadoEtapa(
            etapa=Etapa.CHECKLIST,
            situacao=Situacao.PENDENTE,
            mensagem=(
                (f"Aguardando confirmação: marcar a caixa {onde} "
                 f"({alvo.empresa} / {alvo.fornecedor}), o que deixa a linha verde. "
                 f"Nenhuma alteração foi feita."
                 if por_caixa else
                 f"Aguardando confirmação: pintar {onde} de {cor} "
                 f"({alvo.empresa} / {alvo.fornecedor}). Nenhuma alteração foi feita.")
            ),
            detalhes=detalhes,
        )

    if amb.simulando:
        # Em simulação isto não é erro: é o modo funcionando. Devolvemos
        # PULADO com o detalhe exato do que seria pintado, para a conta poder
        # ser dada como concluída no ensaio sem tocar na planilha real.
        return ResultadoEtapa.pulado(
            Etapa.CHECKLIST,
            (
                f"SIMULADO — {onde} NÃO foi "
                f"{'marcada' if por_caixa else 'pintado'} "
                f"({alvo.empresa} / {alvo.fornecedor}). Desligue "
                "seguranca.simulacao para marcar de verdade."
            ),
            detalhes={**detalhes, "simulacao": True},
        )

    amb.exigir_permissao(caminho)
    backup = _fazer_backup(caminho, amb)
    detalhes["backup"] = str(backup) if backup else None

    try:
        with _excel_com(caminho) as wb:
            if aba not in [f.Name for f in wb.Sheets]:
                raise ErroPlanilha(f"aba {aba!r} sumiu entre a leitura e a gravação")
            if por_caixa:
                _marcar_caixa(wb.Sheets(aba), alvo.celula_marcacao)
            else:
                _pintar_linha(wb.Sheets(aba), alvo.linha, cor)
    except ErroPlanilha:
        raise
    except Exception as exc:
        log.exception("falha ao marcar %s", onde)
        return ResultadoEtapa.erro(
            Etapa.CHECKLIST,
            f"erro ao gravar em {onde}: {type(exc).__name__}: {exc}",
            detalhes=detalhes,
        )

    # Confere o que ficou gravado — relendo do disco, não do cache.
    limpar_cache()
    conferida = next(
        (l for l in ler_checklist(competencia, caminho=caminho) if l.linha == alvo.linha),
        None,
    )
    if conferida is None or not conferida.verde:
        return ResultadoEtapa.erro(
            Etapa.CHECKLIST,
            f"gravou sem erro mas {onde} não ficou "
            f"{'marcada' if por_caixa else 'verde'} — confira à mão",
            detalhes=detalhes,
        )

    return ResultadoEtapa.sucesso(
        Etapa.CHECKLIST,
        (f"caixa {onde} marcada — a linha {alvo.linha} ficou verde ({alvo.fornecedor})"
         if por_caixa else
         f"{onde} pintada de {cor} ({alvo.fornecedor})"),
        detalhes=detalhes,
        artefatos=[caminho],
    )


def criar_aba_mes(
    competencia: Competencia,
    *,
    base: str | None = None,
    somente_ativas: bool = True,
    confirmado: bool = False,
    ambiente_: Ambiente | None = None,
) -> ResultadoEtapa:
    """
    Cria a aba `Checklist-<Mês>` copiando o layout da aba `base`.

    A cópia é feita pelo próprio Excel (`Worksheet.Copy`), então larguras de
    coluna, mesclagens das linhas 1-2, bordas, formatação do cabeçalho,
    hyperlinks e fórmulas vêm de graça e idênticos. Depois disso apenas:
      - as linhas de contas inativas são REMOVIDAS (quando somente_ativas);
      - qualquer verde herdado é apagado (a aba nova nasce toda por fazer);
      - o autofiltro herdado é desligado, como nas demais abas de checklist.

    Nenhuma coluna é criada, nenhuma linha é reordenada, nenhum conteúdo é
    editado. Com `confirmado=False` nada é gravado.
    """
    amb = ambiente_ or ambiente()
    caminho = amb.caminhos.planilha_contas
    nova_aba = competencia.aba_checklist
    cor = cor_enviado(amb)

    existentes = abas_checklist_existentes(caminho=caminho)
    if nova_aba in existentes:
        return ResultadoEtapa.pulado(
            Etapa.CHECKLIST,
            f"a aba {nova_aba!r} já existe — nada a criar",
            detalhes={"aba": nova_aba, "arquivo": str(caminho), "abas": existentes},
        )

    aba_base = base or ABA_CONTAS_FIXAS
    with _planilha_aberta(caminho) as wb:
        if aba_base not in wb.sheetnames:
            return ResultadoEtapa.erro(
                Etapa.CHECKLIST,
                f"aba base {aba_base!r} não existe (abas: {wb.sheetnames})",
                detalhes={"aba_base": aba_base, "arquivo": str(caminho)},
            )
        ws_base = wb[aba_base]
        linhas_base = _ler_aba(ws_base, cor)
        cabecalho_base = [
            _texto(ws_base.cell(LINHA_CABECALHO, coluna).value)
            for coluna in range(COL_EMPRESA, ULTIMA_COLUNA + 1)
        ]
        # A nova aba entra logo depois da última Checklist-* (antes de ACESSOS).
        aba_anterior = existentes[-1] if existentes else aba_base
        posicao_destino = wb.sheetnames.index(aba_anterior) + 2

    manter = [l for l in linhas_base if l.ativa or not somente_ativas]
    numeros_mantidos = {l.linha for l in manter}
    remover = [l for l in linhas_base if l.linha not in numeros_mantidos]
    verdes_herdados = [l.linha for l in manter if l.verde]

    detalhes = {
        "acao": "criar_aba",
        "arquivo": str(caminho),
        "aba_nova": nova_aba,
        "aba_base": aba_base,
        "depois_de": aba_anterior,
        "posicao": posicao_destino,
        "somente_ativas": somente_ativas,
        "linhas_base": len(linhas_base),
        "linhas_previstas": len(manter),
        "linhas_removidas": [
            {"linha": l.linha, "fornecedor": l.fornecedor, "situacao": l.situacao}
            for l in remover
        ],
        "verdes_a_limpar": verdes_herdados,
        "cor_limpa": cor,
        "motor": "excel_com",
        "backup_previsto": _nome_backup_previsto(caminho, amb),
        **_diagnostico_permissao(caminho, amb),
    }

    if not confirmado:
        return ResultadoEtapa(
            etapa=Etapa.CHECKLIST,
            situacao=Situacao.PENDENTE,
            mensagem=(
                f"Aguardando confirmação: criar {nova_aba!r} a partir de {aba_base!r} "
                f"com {len(manter)} linhas ({len(remover)} inativas ficam de fora, "
                f"{len(verdes_herdados)} verdes seriam limpos). "
                "Nenhuma alteração foi feita."
            ),
            detalhes=detalhes,
        )

    amb.exigir_permissao(caminho)
    backup = _fazer_backup(caminho, amb)
    detalhes["backup"] = str(backup) if backup else None

    try:
        with _excel_com(caminho) as wb:
            abas_antes = {str(f.Name) for f in wb.Sheets}
            nova = _duplicar_aba(wb, aba_base, aba_anterior, nova_aba)
            _conferir_cabecalho(nova, cabecalho_base)

            # A ÚNICA exclusão de conteúdo do projeto, e ela só pode tocar a
            # aba recém-criada. Se `nova` apontasse para uma aba existente,
            # aqui apagaria histórico de verdade — por isso a conferência é
            # feita imediatamente antes, e não confiamos no que veio de cima.
            if str(nova.Name) != nova_aba:
                raise ErroPlanilha(
                    f"abortando: esperava trabalhar na aba nova {nova_aba!r} mas "
                    f"o objeto aponta para {nova.Name!r}. Nada foi alterado."
                )
            if nova_aba in abas_antes:
                raise ErroPlanilha(
                    f"abortando: a aba {nova_aba!r} já existia antes desta "
                    "execução. Excluir linhas dela apagaria histórico."
                )

            # Remove de baixo para cima para não bagunçar a numeração.
            for linha in sorted((l.linha for l in remover), reverse=True):
                nova.Rows(linha).Delete()

            # A aba do mês nasce sem nada marcado: desmarca as caixas
            # herdadas (é elas que a formatação condicional lê) e, por
            # garantia, tira também qualquer verde fixo que tenha vindo de
            # uma aba antiga.
            ultima = PRIMEIRA_LINHA_DADOS + len(manter) - 1
            alvo_ole = _cor_ole(cor)
            coluna_caixa = next(
                (l.celula_marcacao for l in manter if l.celula_marcacao), ""
            )
            letra_caixa = re.sub(r"\d+", "", coluna_caixa)
            for numero in range(PRIMEIRA_LINHA_DADOS, ultima + 1):
                if letra_caixa:
                    nova.Range(f"{letra_caixa}{numero}").Value = False
                faixa = nova.Range(f"A{numero}:G{numero}")
                if int(faixa.Cells(1, 1).Interior.Color) == alvo_ole:
                    faixa.Interior.Pattern = _XL_SEM_PADRAO

            if nova.AutoFilterMode:
                nova.AutoFilterMode = False
    except Exception as exc:
        log.exception("falha ao criar a aba %r", nova_aba)
        return ResultadoEtapa.erro(
            Etapa.CHECKLIST,
            f"erro ao criar {nova_aba!r}: {type(exc).__name__}: {exc}",
            detalhes=detalhes,
        )

    # Confere o que ficou gravado — relendo do disco, não do cache.
    limpar_cache()
    criadas = ler_checklist(competencia, caminho=caminho)
    detalhes["linhas_gravadas"] = len(criadas)
    detalhes["verdes_gravados"] = sum(l.verde for l in criadas)

    if len(criadas) != len(manter) or detalhes["verdes_gravados"]:
        return ResultadoEtapa.atencao(
            Etapa.CHECKLIST,
            f"aba {nova_aba!r} criada, mas saiu com {len(criadas)} linhas "
            f"(esperadas {len(manter)}) e {detalhes['verdes_gravados']} verdes "
            "(esperados 0) — confira à mão",
            detalhes=detalhes,
            artefatos=[caminho],
        )

    return ResultadoEtapa.sucesso(
        Etapa.CHECKLIST,
        f"aba {nova_aba!r} criada com {len(manter)} linhas a partir de {aba_base!r}",
        detalhes=detalhes,
        artefatos=[caminho],
    )


# --------------------------------------------------------------------------- #
# Teste manual: python -m automacao.planilha_contas
# --------------------------------------------------------------------------- #

if __name__ == "__main__":
    import hashlib
    import sys

    sys.stdout.reconfigure(encoding="utf-8")
    logging.basicConfig(
        level=logging.INFO, format="%(levelname)-7s %(name)s: %(message)s"
    )

    def _sha256(p: Path) -> str:
        h = hashlib.sha256()
        with p.open("rb") as fh:
            for pedaco in iter(lambda: fh.read(1 << 16), b""):
                h.update(pedaco)
        return h.hexdigest()

    planilha = caminho_planilha()
    print(f"planilha: {planilha}")
    print(f"cor de enviado: {cor_enviado()}")
    antes = _sha256(planilha)
    print(f"SHA-256 antes: {antes}\n")

    fixas = ler_contas_fixas()
    ativas = [l for l in fixas if l.ativa]
    print(f"CONTAS FIXAS ........ {len(fixas)} linhas | {len(ativas)} ativas "
          f"| {sum(l.verde for l in fixas)} verdes")

    print(f"\nabas de checklist: {abas_checklist_existentes()}")
    for ano, mes in ((2026, 3), (2026, 4), (2026, 5), (2026, 6), (2026, 7), (2026, 8)):
        comp = Competencia(ano, mes)
        linhas = ler_checklist(comp)
        print(f"  {comp.aba_checklist:<20} {len(linhas):>3} linhas | "
              f"{sum(l.verde for l in linhas):>3} verdes | "
              f"{sum(l.ativa for l in linhas):>3} ativas")

    acessos = ler_acessos()
    print(f"\nACESSOS E CONTATOS .. {len(acessos)} blocos")
    for cred in acessos:
        print(f"  linha {cred.linha:>3}  {cred.servico:<42} login={cred.login[:34]:<34} "
              f"senha={cred.mascarada()}")

    # Casamento — usa o registro real quando existir, senão contas sintéticas
    # montadas a partir da própria planilha.
    print("\n--- casar_conta ---")
    from automacao.nucleo.config import contas as _contas

    try:
        registro = list(_contas())
    except Exception as exc:
        registro = []
        print(f"  (registro de contas indisponível: {exc})")

    agosto = ler_checklist(Competencia(2026, 8))
    if not registro:
        registro = [
            Conta(id="operadora-a-000000000", pasta="OPERADORA A",
                  chaves_planilha=["OPERADORA A - 000000000"]),
            Conta(id="operadora-b-unidade", pasta="OPERADORA B",
                  subunidade="UNIDADE"),
            Conta(id="software-gama", pasta="SOFTWARE GAMA LTDA"),
            Conta(id="inexistente", pasta="FORNECEDOR QUE NAO EXISTE"),
        ]
    for c in registro[:8]:
        achada = casar_conta(c, agosto)
        if achada is None:
            print(f"  {c.id:<28} -> sem casamento")
        else:
            marca = " (DUVIDOSO)" if achada.casamento_duvidoso else ""
            print(f"  {c.id:<28} -> linha {achada.linha:>3} {achada.fornecedor[:34]:<34} "
                  f"escore {achada.escore:.2f}{marca} [{achada.motivo_casamento}]")

    # Simulações — não podem tocar no arquivo.
    print("\n--- simulações (confirmado=False) ---")
    conta_teste = next((c for c in registro if casar_conta(c, agosto) is not None), None)
    if conta_teste is not None:
        r = marcar_enviado(conta_teste, Competencia(2026, 8))
        print(f"  marcar_enviado({conta_teste.id}) -> {r.situacao.value}")
        print(f"    {r.mensagem}")
        for chave in ("aba", "linha", "intervalo", "cor", "conteudo", "permitido",
                      "bloqueio", "backup_previsto", "motor"):
            print(f"    {chave}: {r.detalhes.get(chave)}")

    r = criar_aba_mes(Competencia(2026, 9))
    print(f"\n  criar_aba_mes -> {r.situacao.value}")
    print(f"    {r.mensagem}")
    for chave in ("aba_nova", "aba_base", "posicao", "linhas_base", "linhas_previstas",
                  "verdes_a_limpar", "permitido", "bloqueio", "backup_previsto"):
        print(f"    {chave}: {r.detalhes.get(chave)}")
    print(f"    linhas_removidas: {len(r.detalhes.get('linhas_removidas') or [])}")

    depois = _sha256(planilha)
    print(f"\nSHA-256 depois: {depois}")
    print("planilha INTACTA" if antes == depois else "*** PLANILHA ALTERADA ***")

    print("\n" + _MOTIVO_COM)
