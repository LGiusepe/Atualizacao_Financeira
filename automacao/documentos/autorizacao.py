"""
Preenchimento da planilha "Autorização de Pagamento" do mês.

A planilha do mês novo nasce de uma cópia do xlsx do mês anterior (o
`modelo_base` da conta). A cópia é aberta com openpyxl **sem**
`data_only=True`, de forma que todas as fórmulas do modelo continuam vivas
(`=VLOOKUP(...)`, `=TODAY()`, `=SUM(E29:E36)`, `=C22+F22`, `=H7`, `=C22`).
Só os campos de valor são reescritos; nenhuma fórmula é tocada.

Regra de segurança: o arquivo gerado vai SEMPRE para a área local de
trabalho (`ambiente().pasta_trabalho(...)`). A publicação no OneDrive é
responsabilidade de outro módulo.

Diferenças conhecidas entre o modelo e o arquivo gerado
-------------------------------------------------------
O openpyxl reescreve o pacote xlsx inteiro ao salvar e, no caminho, joga fora
algumas partes. `restaurar_recursos()` devolve as que importam (imagens,
desenhos e configuração de impressora — é onde vive a logomarca da GOL no
cabeçalho). Continuam se perdendo, de propósito ou por limitação:

* `xl/calcChain.xml` — some por escolha nossa: força o Excel a recalcular tudo;
* `xl/sharedStrings.xml` — o openpyxl grava textos em linha, é equivalente;
* `customXml/*` — metadados do SharePoint, não referenciados pelo arquivo novo;
* validações de dados da extensão x14 (as listas suspensas que apontam para
  outra aba) — o openpyxl não sabe reescrevê-las. As validações comuns
  sobrevivem. Como o arquivo gerado é um entregável, e não um modelo para
  digitação, a perda das listas não afeta o PDF nem os valores.
"""

from __future__ import annotations

import logging
import re
import shutil
import zipfile
from datetime import date, datetime
from pathlib import Path
from string import Formatter
from typing import Any
from xml.etree import ElementTree as ET

import openpyxl
from openpyxl.utils.cell import coordinate_to_tuple, get_column_letter
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from automacao.nucleo.config import Ambiente, ambiente, beneficiario_por_nome, pagante_por_nome
from automacao.nucleo.modelos import Competencia, Conta, Etapa, ResultadoEtapa, normalizar

registrador = logging.getLogger("automacao.documentos.autorizacao")


# --------------------------------------------------------------------------- #
# Mapa de células — confirmado na base real, idêntico nas duas variantes de
# template ("AUTORIZAÇÃO" e "MODELO").
# --------------------------------------------------------------------------- #

CELULAS: dict[str, str] = {
    "departamento": "B6",
    "pagante": "B7",
    "vencimento": "H7",
    "beneficiario": "C13",
    "forma_pgto": "F18",
    "meio_pgto": "J18",
    "valor": "C22",
    "setor": "B26",
    "motivo": "F26",
    "regional": "J26",
    "rotulo_documento": "A29",
    "numero_documento": "C29",
    "centro_custo": "F29",
    "natureza": "H29",
    "cooperativa": "J29",
    "descricao": "A32",
    "pgto_previsto": "C42",
}

#: Campos que são explicitamente um valor — os únicos autorizados a
#: sobrescrever uma célula que porventura contenha fórmula.
CAMPOS_VALOR: frozenset[str] = frozenset({"valor", "vencimento", "numero_documento"})

#: Células calculadas do template. Nunca são escritas; ficam aqui só para
#: documentar o contrato e alimentar a conferência de sanidade.
CELULAS_CALCULADAS: dict[str, str] = {
    "H6": "=TODAY() — data da solicitação",
    "C18": "=H7 — data de pagamento espelha o vencimento",
    "B9": "=VLOOKUP(B7;...) — CNPJ do pagante",
    "I13": "=VLOOKUP(C13;...) — CNPJ do beneficiário (variante AUTORIZAÇÃO)",
    "J22": "=C22 ou =C22+F22 — valor pago",
    "E29": "=C22 ou =J22 — valor da linha de rateio",
    "E37": "=SUM(E29:E36) — total",
}

#: Placeholders aceitos em `conta.autorizacao.descricao`.
PLACEHOLDERS_DESCRICAO = (
    "competencia_extenso",
    "mes_anterior_extenso",
    "cidade",
    "competencia",
)

_RE_VLOOKUP = re.compile(
    r"VLOOKUP\(\s*"
    r"(?P<referencia>\$?[A-Z]{1,3}\$?\d{1,7})\s*[,;]\s*"
    # A tabela consultada, em duas formas possíveis: o intervalo normal, ou
    # `#REF!` — o que o Excel deixa quando a aba de consulta é apagada. A
    # segunda forma precisa ser reconhecida, senão a célula sai #REF! no PDF
    # e nada avisa. Aconteceu no modelo de um fornecedor, com seis fórmulas.
    r"(?:"
    r"(?:(?P<aba>'[^']+'|[^!,;()]+)!)?"
    r"(?P<faixa>\$?[A-Z]{1,3}\$?\d{1,7}\s*:\s*\$?[A-Z]{1,3}\$?\d{1,7})"
    r"|(?P<quebrada>\#REF\!)"
    r")"
    # A coluna devolvida distingue um VLOOKUP do outro quando varios buscam a
    # MESMA chave: em C13 o template procura CNPJ (2), contato (3), celular
    # (4) e banco (5). Opcional de proposito — indice calculado nao casa aqui,
    # e a conferencia de sanidade nao precisa dele.
    r"(?:\s*[,;]\s*(?P<indice>\d{1,3}))?",
    re.IGNORECASE,
)


class ErroAutorizacao(RuntimeError):
    """Falha ao preparar ou preencher a planilha de autorização."""


class CelulaProtegida(ErroAutorizacao):
    """Tentativa de sobrescrever uma célula que contém fórmula."""


# --------------------------------------------------------------------------- #
# Utilitários de célula
# --------------------------------------------------------------------------- #


def ancora_mesclada(ws: Worksheet, coordenada: str) -> str:
    """
    Devolve a coordenada gravável de `coordenada`.

    Em um range mesclado só a âncora (canto superior esquerdo) aceita
    escrita — as demais são `MergedCell` somente leitura. Se a coordenada
    não estiver mesclada, ela mesma é devolvida.
    """
    linha, coluna = coordinate_to_tuple(coordenada)
    for faixa in ws.merged_cells.ranges:
        if (
            faixa.min_row <= linha <= faixa.max_row
            and faixa.min_col <= coluna <= faixa.max_col
        ):
            return f"{get_column_letter(faixa.min_col)}{faixa.min_row}"
    return coordenada.upper()


def _tem_formula(valor: Any) -> bool:
    return isinstance(valor, str) and valor.startswith("=")


def _simplificar(valor: Any) -> Any:
    """Normaliza para comparação e para exibição no painel."""
    if isinstance(valor, datetime):
        return valor.date()
    return valor


def _iguais(antigo: Any, novo: Any) -> bool:
    a, b = _simplificar(antigo), _simplificar(novo)
    if a is None and b is None:
        return True
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(float(a) - float(b)) < 0.005
    if isinstance(a, str) and isinstance(b, str):
        return a == b
    return a == b


def escrever(
    ws: Worksheet,
    coordenada: str,
    valor: Any,
    *,
    permitir_formula: bool = False,
) -> tuple[Any, Any] | None:
    """
    Escreve `valor` em `coordenada`, respeitando mesclagem e formatação.

    - resolve a âncora do range mesclado automaticamente;
    - preserva o formato numérico/data já existente na célula;
    - recusa sobrescrever fórmula (`CelulaProtegida`), salvo `permitir_formula`.

    Devolve `(valor_antigo, valor_novo)` ou `None` se nada mudou.
    """
    alvo = ancora_mesclada(ws, coordenada)
    celula = ws[alvo]
    antigo = celula.value

    if _tem_formula(antigo) and not permitir_formula:
        raise CelulaProtegida(
            f"{ws.title}!{alvo} contém a fórmula {antigo!r} — não será sobrescrita."
        )

    if _iguais(antigo, valor):
        return None

    formato = celula.number_format
    celula.value = valor
    # openpyxl troca o formato ao receber date/datetime; devolvemos o original
    # (a menos que a célula fosse "General", onde o formato novo é desejável).
    if formato and formato != "General" and celula.number_format != formato:
        celula.number_format = formato

    registrador.debug("%s!%s: %r -> %r", ws.title, alvo, antigo, valor)
    return _simplificar(antigo), _simplificar(valor)


# --------------------------------------------------------------------------- #
# Descrição com placeholders
# --------------------------------------------------------------------------- #


class _DicionarioTolerante(dict):
    """Devolve o próprio placeholder quando a chave não tem valor."""

    def __init__(self, dados: dict[str, Any]):
        super().__init__({k: v for k, v in dados.items() if v not in (None, "")})
        self.faltando: set[str] = set()

    def __missing__(self, chave: str) -> str:
        self.faltando.add(chave)
        return "{" + chave + "}"


def formatar_descricao(
    modelo: str | None,
    competencia: Competencia,
    cidade: str | None = None,
) -> tuple[str | None, set[str]]:
    """
    Resolve os placeholders da descrição do evento.

    Placeholders sem valor permanecem literais no texto (nada de `KeyError`).
    Devolve `(texto, placeholders_sem_valor)`.
    """
    if not modelo:
        return modelo, set()

    valores = _DicionarioTolerante(
        {
            "competencia_extenso": competencia.extenso,
            "mes_anterior_extenso": competencia.anterior().extenso,
            "competencia": str(competencia),
            "cidade": cidade,
        }
    )
    try:
        texto = Formatter().vformat(modelo, (), valores)
    except (ValueError, IndexError, AttributeError) as erro:
        # Texto com chaves soltas ou format spec inválido: melhor devolver o
        # original do que estourar no meio do preenchimento.
        registrador.warning("descrição não pôde ser formatada (%s): %r", erro, modelo)
        return modelo, set()
    return texto, valores.faltando


# --------------------------------------------------------------------------- #
# Meses escritos à mão dentro da descrição
# --------------------------------------------------------------------------- #

MESES_EXTENSO = (
    "janeiro", "fevereiro", "março", "abril", "maio", "junho",
    "julho", "agosto", "setembro", "outubro", "novembro", "dezembro",
)

#: `março` aparece com e sem cedilha na base.
_MES_POR_NOME = {}
for _i, _nome in enumerate(MESES_EXTENSO, start=1):
    _MES_POR_NOME[_nome] = _i
    _MES_POR_NOME[_nome.replace("ç", "c")] = _i

_RE_MES_EXTENSO = re.compile(
    r"\b(" + "|".join(sorted(_MES_POR_NOME, key=len, reverse=True)) + r")\b"
    r"(?P<sufixo>\s*(?:de\s*)?[/-]?\s*(?P<ano>\d{4}|\d{2})\b)?",
    re.IGNORECASE,
)
#: `03/2026`, `05-26`. Exige separador para não capturar número de contrato.
_RE_MES_NUMERICO = re.compile(r"\b(?P<mes>0[1-9]|1[0-2])\s*[/-]\s*(?P<ano>\d{4}|\d{2})\b")


def _preservar_caixa(molde: str, novo: str) -> str:
    """`JUNHO`→`AGOSTO`, `Junho`→`Agosto`, `junho`→`agosto`."""
    if molde.isupper():
        return novo.upper()
    if molde[:1].isupper():
        return novo.capitalize()
    return novo


def achar_mes_no_texto(texto: str | None) -> tuple[int, int | None] | None:
    """
    Primeiro mês citado no texto, como `(mes, ano)`. `ano` é `None` se ausente.

    Serve para descobrir a que período a descrição se refere — as da base
    trazem o mês escrito à mão (`"referente ao mês de junho/26"`).
    """
    if not texto:
        return None
    achado = _RE_MES_EXTENSO.search(texto)
    if achado:
        ano = achado.group("ano")
        return _MES_POR_NOME[achado.group(1).lower()], _ano_cheio(ano)
    achado = _RE_MES_NUMERICO.search(texto)
    if achado:
        return int(achado.group("mes")), _ano_cheio(achado.group("ano"))
    return None


def _ano_cheio(bruto: str | None) -> int | None:
    if not bruto:
        return None
    numero = int(bruto)
    return numero if numero > 100 else 2000 + numero


def trocar_mes_no_texto(texto: str, mes: int, ano: int) -> str:
    """
    Reescreve o primeiro mês citado, mantendo o formato original.

    `"...referente ao mês de junho/26"` com (8, 2026) vira
    `"...referente ao mês de agosto/26"`; `"ref.03/2026"` vira `"ref.08/2026"`.
    Só o primeiro: descrição com dois meses é caso de conferência humana.
    """
    nome = MESES_EXTENSO[mes - 1]

    def por_extenso(achado: re.Match) -> str:
        trocado = _preservar_caixa(achado.group(1), nome)
        sufixo = achado.group("sufixo") or ""
        if achado.group("ano"):
            curto = len(achado.group("ano")) == 2
            sufixo = sufixo.replace(
                achado.group("ano"), f"{ano % 100:02d}" if curto else str(ano)
            )
        return trocado + sufixo

    novo, trocas = _RE_MES_EXTENSO.subn(por_extenso, texto, count=1)
    if trocas:
        return novo

    def numerico(achado: re.Match) -> str:
        curto = len(achado.group("ano")) == 2
        return (
            achado.group(0)
            .replace(achado.group("mes"), f"{mes:02d}")
            .replace(achado.group("ano"), f"{ano % 100:02d}" if curto else str(ano))
        )

    novo, trocas = _RE_MES_NUMERICO.subn(numerico, texto, count=1)
    return novo if trocas else texto


def _competencia_do_modelo(conta: Conta) -> Competencia | None:
    """
    Competência do modelo — `FORNECEDOR/07-2026/TI - ... - 07-2026.xlsx` dá
    julho/2026.

    A **pasta** manda, não o nome do arquivo: há casos na base em que o xlsx
    foi copiado sem renomear (`FORNECEDOR/08-2026/UNIDADE/...- 2026 -07.xlsx`),
    e aí o nome aponta para um mês que não é o daquela autorização.

    O que o caminho mostra é o mês da PASTA. Em conta que arquiva fora de fase
    (`deslocamento_pasta`), esse não é o mês da autorização: o modelo de quem
    adianta a pasta mora em `2026-09` mas é a fatura que venceu em agosto. Desandamos o
    deslocamento para devolver a competência de verdade — é ela que serve de
    régua para o mês escrito na descrição.
    """
    pedacos = (conta.modelo_base or "").replace("\\", "/").split("/")
    for pedaco in list(reversed(pedacos[:-1])) + pedacos[-1:]:
        # "2026-09(01)" — sufixo de cópia que o Explorador põe. Sem tirar,
        # sobram três números e a leitura falha.
        limpo = re.sub(r"\([^)]*\)", " ", pedaco)
        try:
            achado = Competencia.de_texto(limpo)
        except ValueError:
            continue
        return achado.somar(-conta.deslocamento_pasta)
    return None


def sugerir_descricao(
    conta: Conta, competencia: Competencia
) -> tuple[str | None, str | None]:
    """
    Descrição do evento já com o mês acertado para `competencia`.

    A base guarda o texto do mês passado com o mês escrito à mão
    (`"referente ao mês de junho/26"`), e ele saía intacto na autorização do
    mês seguinte. Aqui o mês anda junto: o deslocamento é o mesmo que existia
    entre o texto e o modelo — se a autorização de julho falava de junho, a de
    setembro fala de agosto.

    Não adivinha sozinha quando o deslocamento parece errado; devolve também
    uma explicação, e o painel deixa você corrigir antes de gerar.

    Returns:
        `(texto, explicacao)`. `explicacao` é `None` quando não houve nada a
        ajustar.
    """
    base = conta.autorizacao.descricao
    if not base:
        return None, None

    texto, _faltando = formatar_descricao(base, competencia, conta.email.cidade)
    if any("{" + p in base for p in PLACEHOLDERS_DESCRICAO):
        return texto, "mês preenchido pelo modelo do registro"

    achado = achar_mes_no_texto(texto)
    if not achado:
        return texto, None

    mes_texto, ano_texto = achado
    modelo = _competencia_do_modelo(conta)
    if modelo is None:
        return texto, "não sei de que mês é o modelo — confira o mês do texto"

    # Distância em meses entre o que o texto dizia e o mês daquele modelo.
    if ano_texto is None:
        # Sem ano no texto: assume o ciclo mais próximo do modelo (cobre o
        # caso dezembro→janeiro sem inventar um salto de 11 meses).
        deslocamento = mes_texto - modelo.mes
        if deslocamento > 6:
            deslocamento -= 12
        elif deslocamento < -6:
            deslocamento += 12
    else:
        deslocamento = (ano_texto * 12 + mes_texto) - (modelo.ano * 12 + modelo.mes)

    absoluto = competencia.ano * 12 + competencia.mes + deslocamento
    alvo = Competencia((absoluto - 1) // 12, (absoluto - 1) % 12 + 1)
    novo = trocar_mes_no_texto(texto, alvo.mes, alvo.ano)

    if deslocamento == 0:
        como = "mesmo mês da competência"
    elif deslocamento == -1:
        como = "um mês antes da competência, como estava no modelo"
    else:
        quantos = abs(deslocamento)
        lado = "antes" if deslocamento < 0 else "depois"
        como = (
            f"{quantos} meses {lado} da competência — era essa a distância no "
            f"modelo de {modelo}. Confira se é isso mesmo"
        )
    return novo, f"{MESES_EXTENSO[alvo.mes - 1]}/{alvo.ano} · {como}"


# --------------------------------------------------------------------------- #
# Conferência das fórmulas VLOOKUP
# --------------------------------------------------------------------------- #


def _coluna_de_busca(wb: Workbook, nome_aba: str | None, faixa: str, padrao: Worksheet) -> list[Any]:
    """Primeira coluna da tabela usada por um VLOOKUP."""
    aba = padrao
    if nome_aba:
        limpo = nome_aba.strip().strip("'").replace("''", "'")
        if limpo not in wb.sheetnames:
            return []
        aba = wb[limpo]
    inicio, fim = (p.strip().replace("$", "") for p in faixa.split(":"))
    linha_ini, col_ini = coordinate_to_tuple(inicio)
    linha_fim, _ = coordinate_to_tuple(fim)
    letra = get_column_letter(col_ini)
    return [aba[f"{letra}{i}"].value for i in range(linha_ini, linha_fim + 1)]


def conferir_vlookups(
    wb: Workbook,
    ws: Worksheet,
    escritas: dict[str, Any],
) -> list[str]:
    """
    Verifica se os valores recém-escritos ainda encontram par nas tabelas
    consultadas pelos VLOOKUP da aba — é isso que evita um `#N/D` no PDF.

    Devolve a lista de avisos (vazia = tudo resolve).
    """
    avisos: list[str] = []
    for linha in ws.iter_rows():
        for celula in linha:
            if not _tem_formula(celula.value):
                continue
            for achado in _RE_VLOOKUP.finditer(str(celula.value)):
                referencia = achado.group("referencia").replace("$", "").upper()
                if referencia not in escritas:
                    continue
                procurado = escritas[referencia]
                if procurado in (None, ""):
                    continue
                if achado.group("quebrada"):
                    avisos.append(
                        f"{celula.coordinate} consulta uma tabela que não existe "
                        f"mais no modelo (#REF!) e o cadastro não tem esse dado — "
                        f"a célula vai sair como #REF! no PDF."
                    )
                    continue
                try:
                    chaves = _coluna_de_busca(wb, achado.group("aba"), achado.group("faixa"), ws)
                except Exception as erro:  # noqa: BLE001 — conferência é best-effort
                    registrador.debug("VLOOKUP em %s não pôde ser conferido: %s", celula.coordinate, erro)
                    continue
                if not chaves:
                    continue
                alvo = str(procurado)
                if any(str(c) == alvo for c in chaves if c is not None):
                    continue
                if any(str(c).casefold() == alvo.casefold() for c in chaves if c is not None):
                    continue  # Excel compara sem diferenciar maiúsculas
                avisos.append(
                    f"{celula.coordinate} usa VLOOKUP em {referencia} e o valor "
                    f"{alvo!r} não existe na tabela de consulta — a célula vai "
                    f"aparecer como #N/D no PDF."
                )
    return avisos


#: Como a aba BENEFICIÁRIOS do template está organizada: coluna do VLOOKUP
#: -> campo do nosso cadastro. A ordem é a mesma do dataclass `Beneficiario`,
#: mas isso é conveniência, não garantia: `_layout_confere` verifica antes de
#: escrever qualquer coisa.
COLUNAS_BENEFICIARIOS: dict[int, str] = {
    2: "cnpj", 3: "contato", 4: "telefone", 5: "banco", 6: "agencia", 7: "conta",
}


def _so_digitos(texto: Any) -> str:
    return re.sub(r"\D", "", str(texto or ""))


def _texto_simples(valor: Any) -> str:
    return "" if valor is None else str(valor).strip()


def _layout_confere(wb: Workbook, aba: str | None, faixa: str, ws: Worksheet) -> bool:
    """
    Confere se a tabela de beneficiários está na ordem que esperamos.

    Escrever o telefone dentro do campo de CNPJ seria pior do que deixar o
    `#N/D` à mostra. Então, antes de substituir qualquer fórmula, procuramos
    na tabela um beneficiário que TAMBÉM esteja no nosso cadastro e conferimos
    se o CNPJ dele está mesmo na coluna 2. Se não estiver — ou se não houver
    ninguém em comum para comparar — desistimos e o aviso de `#N/D` fica.
    """
    try:
        planilha = wb[aba.strip("'")] if aba else ws
        inicio, fim = (parte.strip().replace("$", "") for parte in faixa.split(":"))
        linhas = planilha[f"{inicio}:{fim}"]
    except Exception:  # noqa: BLE001 — conferência é best-effort
        return False

    # Prova direta: a própria tabela diz o que há em cada coluna. Quando o
    # cabeçalho existe, ele vale mais que qualquer cruzamento — foi o que
    # faltou num dos fornecedores, cuja tabela (`Planilha2`) não tem nenhum
    # beneficiário em comum com o nosso cadastro para servir de referência.
    if linhas and len(linhas[0]) >= 2:
        titulos = [normalizar(_texto_simples(c.value)) for c in linhas[0][:2]]
        if titulos[0].startswith("nome") and titulos[1].startswith("cnpj"):
            return True

    for linha in linhas:
        nome = linha[0].value if linha else None
        if not isinstance(nome, str) or not nome.strip():
            continue
        registrado = beneficiario_por_nome(nome)
        if not registrado or not registrado.cnpj:
            continue
        if len(linha) < 2:
            return False
        return _so_digitos(linha[1].value) == _so_digitos(registrado.cnpj)
    return False


def suprir_vlookup_quebrado(
    wb: Workbook,
    ws: Worksheet,
    escritas: dict[str, Any],
    socorros: dict[str, dict[int | None, str]],
) -> list[str]:
    """
    Escreve o valor conhecido onde o VLOOKUP não tem como encontrar.

    O caso real: a tabela `EMPRESAS` do template guarda as razões sociais
    ANTIGAS. Ao escrever o nome atual em B7, o `=VLOOKUP(B7;…)` de B9 devolve
    `#N/D` — e a autorização chega ao financeiro sem o CNPJ do pagante. Vinha
    saindo assim havia meses.

    Como o cadastro de pagantes tem o CNPJ, dá para preencher a célula
    diretamente. É a única situação em que este módulo substitui uma fórmula
    por um valor, e ela sempre vira aviso no painel — trocar cálculo por
    texto fixo é decisão que precisa ficar à vista.

    O beneficiário tem o mesmo problema com mais campos: C13 alimenta quatro
    fórmulas (CNPJ, contato, celular, banco) que só se distinguem pela coluna
    devolvida. Por isso o socorro é indexado por coluna, com `None` valendo
    para qualquer uma.

    Args:
        socorros: `{referência: {coluna do VLOOKUP: valor}}`. A chave `None`
            atende qualquer coluna — é o caso do pagante, que só tem CNPJ.
            Ex.: `{"B7": {None: cnpj}}`, `{"C13": {2: cnpj, 3: contato}}`.

    Returns:
        Avisos do que foi substituído.
    """
    avisos: list[str] = []
    for linha in ws.iter_rows():
        for celula in linha:
            if not _tem_formula(celula.value):
                continue
            for achado in _RE_VLOOKUP.finditer(str(celula.value)):
                referencia = achado.group("referencia").replace("$", "").upper()
                por_coluna = socorros.get(referencia) or {}
                bruto = achado.group("indice")
                coluna = int(bruto) if bruto else None
                socorro = por_coluna.get(coluna)
                por_indice = socorro is not None
                if socorro is None:
                    socorro = por_coluna.get(None)
                procurado = escritas.get(referencia)
                if not socorro or procurado in (None, ""):
                    continue
                # Tabela apagada: não há o que consultar nem layout que
                # conferir. A fórmula sairia #REF! para qualquer nome, então
                # o valor do cadastro é estritamente melhor que ela.
                if achado.group("quebrada"):
                    formula = celula.value
                    celula.value = socorro
                    avisos.append(
                        f"{celula.coordinate}: a fórmula apontava para uma tabela "
                        f"que não existe mais (#REF!) e sairia assim no PDF. "
                        f"Usei o valor do cadastro ({socorro})."
                    )
                    registrador.warning(
                        "%s!%s: VLOOKUP com tabela apagada (#REF!) -> %r",
                        ws.title, celula.coordinate, socorro,
                    )
                    continue

                # Socorro escolhido POR COLUNA supõe uma ordem de colunas na
                # tabela consultada. Se a suposição não se confirmar, escrever
                # seria pôr o telefone no campo do CNPJ — pior que o #N/D.
                if por_indice and not _layout_confere(
                    wb, achado.group("aba"), achado.group("faixa"), ws
                ):
                    registrador.warning(
                        "%s!%s: a tabela consultada não está na ordem esperada — "
                        "deixo a fórmula como está.", ws.title, celula.coordinate,
                    )
                    continue
                try:
                    chaves = _coluna_de_busca(
                        wb, achado.group("aba"), achado.group("faixa"), ws
                    )
                except Exception:  # noqa: BLE001 — socorro é best-effort
                    continue
                if not chaves:
                    continue
                alvo = str(procurado).casefold()
                if any(str(c).casefold() == alvo for c in chaves if c is not None):
                    continue  # o VLOOKUP resolve sozinho; não se mexe

                formula = celula.value
                celula.value = socorro
                avisos.append(
                    f"{celula.coordinate}: a tabela de consulta não tem "
                    f"{procurado!r}, então a fórmula foi trocada pelo valor do "
                    f"cadastro ({socorro}). Sem isso a célula sairia como #N/D."
                )
                registrador.info(
                    "%s!%s: %r -> %r (VLOOKUP sem correspondência)",
                    ws.title, celula.coordinate, formula, socorro,
                )
    return avisos


# --------------------------------------------------------------------------- #
# Restauração das partes que o openpyxl descarta ao salvar
# --------------------------------------------------------------------------- #

#: Partes do pacote xlsx que o openpyxl não sabe reescrever e simplesmente
#: descarta. `xl/media` + `xl/drawings` são a LOGOMARCA da empresa no
#: cabeçalho da autorização — sem elas o PDF sai sem logo.
_PREFIXOS_RESTAURAVEIS = ("xl/media/", "xl/drawings/", "xl/printerSettings/")

_NS_PLANILHA = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_NS_REL_DOC = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_NS_TIPOS = "{http://schemas.openxmlformats.org/package/2006/content-types}"
_TIPO_DESENHO = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
#: Tipo de conteúdo de uma parte de desenho, para o [Content_Types].xml.
_TIPO_CONTEUDO_DESENHO = (
    "application/vnd.openxmlformats-officedocument.drawing+xml"
)


def _mapa_abas(pacote: zipfile.ZipFile) -> dict[str, str]:
    """`{nome da aba: 'xl/worksheets/sheetN.xml'}` — o N não segue a ordem."""
    alvos = {
        rel.get("Id"): (rel.get("Target") or "")
        for rel in ET.fromstring(pacote.read("xl/_rels/workbook.xml.rels"))
    }
    mapa: dict[str, str] = {}
    livro = ET.fromstring(pacote.read("xl/workbook.xml"))
    abas = livro.find(f"{_NS_PLANILHA}sheets")
    for aba in abas if abas is not None else []:
        destino = alvos.get(aba.get(f"{_NS_REL_DOC}id"), "").lstrip("/")
        if destino and not destino.startswith("xl/"):
            destino = "xl/" + destino
        if destino:
            mapa[aba.get("name", "")] = destino
    return mapa


def _inserir_desenho(xml: bytes, rid: str) -> bytes:
    """Recoloca `<drawing r:id="..."/>` na aba, na posição exigida pelo schema."""
    texto = xml.decode("utf-8")
    if "<drawing " in texto:
        return xml

    # O openpyxl só declara xmlns:r quando precisa; sem isso o prefixo do
    # <drawing> fica solto e o XML deixa de ser válido.
    fim_raiz = texto.find(">", texto.find("<worksheet"))
    if fim_raiz != -1 and "xmlns:r=" not in texto[:fim_raiz]:
        texto = (
            texto[:fim_raiz]
            + f' xmlns:r="{_NS_REL_DOC[1:-1]}"'
            + texto[fim_raiz:]
        )

    marca = f'<drawing r:id="{rid}"/>'
    for antes in ("<tableParts", "<extLst", "</worksheet>"):
        posicao = texto.find(antes)
        if posicao != -1:
            return (texto[:posicao] + marca + texto[posicao:]).encode("utf-8")
    return texto.encode("utf-8")


def _mesclar_tipos(gerado: bytes, modelo: bytes, restauradas: set[str]) -> bytes:
    """Declara no [Content_Types].xml as partes que voltaram para o pacote."""
    extensoes = {Path(p).suffix.lstrip(".").lower() for p in restauradas}
    atual = ET.fromstring(gerado)
    ja_tem_ext = {
        (e.get("Extension") or "").lower()
        for e in atual.findall(f"{_NS_TIPOS}Default")
    }
    ja_tem_parte = {e.get("PartName") for e in atual.findall(f"{_NS_TIPOS}Override")}

    novos: list[str] = []
    for entrada in ET.fromstring(modelo):
        etiqueta = entrada.tag.replace(_NS_TIPOS, "")
        if etiqueta == "Default":
            ext = (entrada.get("Extension") or "").lower()
            if ext in extensoes and ext not in ja_tem_ext:
                novos.append(
                    f'<Default Extension="{ext}" ContentType="{entrada.get("ContentType")}"/>'
                )
                ja_tem_ext.add(ext)
        elif etiqueta == "Override":
            parte = entrada.get("PartName") or ""
            if parte.lstrip("/") in restauradas and parte not in ja_tem_parte:
                novos.append(
                    f'<Override PartName="{parte}" ContentType="{entrada.get("ContentType")}"/>'
                )
                ja_tem_parte.add(parte)

    if not novos:
        return gerado
    texto = gerado.decode("utf-8")
    return texto.replace("</Types>", "".join(novos) + "</Types>").encode("utf-8")


def _coluna(referencia: str) -> int:
    """`A`→1, `J`→10. Usado só para manter as células em ordem na linha."""
    numero = 0
    for letra in referencia:
        if not letra.isalpha():
            break
        numero = numero * 26 + (ord(letra.upper()) - 64)
    return numero


def _devolver_celulas_de_estilo(gerado: bytes, modelo: bytes) -> tuple[bytes, int]:
    """
    Recoloca as células vazias que só existiam para carregar formatação.

    O openpyxl não grava célula sem conteúdo cujo estilo é o padrão (`s="0"`).
    Parece inofensivo, mas nesta planilha **não é**: a célula sem estilo herda
    o estilo padrão da coluna, que aqui é branco com 25% de escurecimento. O
    resultado no PDF é uma **tarja cinza** atrás da logomarca e no bloco de
    DATA/VENCIMENTO — 48 células num dos modelos da base.

    A devolução é cirúrgica: só entram referências que sumiram e que estavam
    vazias no modelo, então nada que a automação escreveu é tocado.
    """
    try:
        arvore_modelo = ET.fromstring(modelo)
        arvore_gerada = ET.fromstring(gerado)
    except ET.ParseError:
        return gerado, 0

    vazias = {
        cel.get("r"): cel.attrib
        for linha in arvore_modelo.iter(f"{_NS_PLANILHA}row")
        for cel in linha
        if cel.get("r") and len(list(cel)) == 0
    }
    presentes = {
        cel.get("r")
        for linha in arvore_gerada.iter(f"{_NS_PLANILHA}row")
        for cel in linha
    }
    faltando: dict[int, list[tuple[int, str]]] = {}
    for referencia, atributos in vazias.items():
        if referencia in presentes:
            continue
        linha = int("".join(c for c in referencia if c.isdigit()) or 0)
        estilo = atributos.get("s")
        marca = f'<c r="{referencia}"' + (f' s="{estilo}"' if estilo else "") + "/>"
        faltando.setdefault(linha, []).append((_coluna(referencia), marca))
    if not faltando:
        return gerado, 0

    texto = gerado.decode("utf-8")
    total = 0
    for numero, celulas in faltando.items():
        padrao = re.compile(
            rf'<row\b(?P<atrib>[^>]*\br="{numero}"[^>]*?)(?P<fecho>/>|>(?P<corpo>.*?)</row>)',
            re.S,
        )
        achado = padrao.search(texto)
        if not achado:
            continue  # linha inexistente no gerado: não vale inventar

        corpo = achado.group("corpo") or ""
        for coluna, marca in sorted(celulas):
            posicao = len(corpo)
            for existente in re.finditer(r'<c\b[^>]*\br="([A-Z]+\d+)"', corpo):
                if _coluna(existente.group(1)) > coluna:
                    posicao = existente.start()
                    break
            corpo = corpo[:posicao] + marca + corpo[posicao:]
            total += 1

        novo = f'<row{achado.group("atrib")}>{corpo}</row>'
        texto = texto[: achado.start()] + novo + texto[achado.end() :]

    return texto.encode("utf-8"), total


def _caminho_no_pacote(destino: str) -> str:
    """
    Normaliza o Target de um .rels de aba para o nome dentro do zip.

    As três formas aparecem na base: `../drawings/drawing1.xml` (Excel),
    `/xl/drawings/drawing1.xml` (absoluta) e `xl/drawings/drawing1.xml`
    (openpyxl). Sem tratar as três, sai `xl/xl/drawings/...`.
    """
    destino = (destino or "").strip().replace("\\", "/")
    if not destino:
        return ""
    if destino.startswith("/"):
        return destino.lstrip("/")
    if destino.startswith("../"):
        return "xl/" + destino[3:]
    if destino.startswith("xl/"):
        return destino
    return "xl/" + destino


def _alvo_do_desenho(rels: bytes | None) -> str | None:
    """`xl/drawings/drawing1.xml` a partir do .rels de uma aba."""
    if not rels:
        return None
    for rel in ET.fromstring(rels):
        if rel.get("Type") == _TIPO_DESENHO:
            return _caminho_no_pacote(rel.get("Target") or "") or None
    return None


def _devolver_desenhos_originais(
    conteudo: dict[str, bytes],
    extras: dict[str, bytes],
    abas_modelo: dict[str, str],
    abas_geradas: dict[str, str],
    rels_modelo: dict[str, bytes],
    novos_desenhos: set[str],
    dono: dict[str, str],
) -> list[str]:
    """
    Troca os desenhos reescritos pelo openpyxl pelos do modelo.

    Aqui não é descarte, é reescrita com perda: o openpyxl remonta
    `xl/drawings/drawingN.xml` e deixa de fora as propriedades da forma —
    entre elas o `<a:noFill/>` da imagem. Sem esse `noFill`, o Excel pinta o
    preenchimento padrão atrás da logomarca e o PDF sai com uma **tarja
    cinza** sobre o cabeçalho. Copiar o desenho do modelo resolve, e é seguro:
    a automação não mexe em imagem nenhuma.

    O `.rels` do desenho vai junto — é ele que liga o `rId` ao arquivo em
    `xl/media/`, e separar os dois deixaria a imagem sem origem.
    """
    trocados: list[str] = []
    for nome_aba, arquivo_modelo in abas_modelo.items():
        arquivo_gerado = abas_geradas.get(nome_aba)
        if not arquivo_gerado:
            continue

        de = _alvo_do_desenho(rels_modelo.get(
            f"xl/worksheets/_rels/{Path(arquivo_modelo).name}.rels"))
        para = _alvo_do_desenho(conteudo.get(
            f"xl/worksheets/_rels/{Path(arquivo_gerado).name}.rels"))
        if not de or not para or de not in extras:
            continue

        # Duas abas com o MESMO desenho: o openpyxl renumerou e colidiu. Sem
        # separar, a segunda sobrescreve a primeira e uma das abas exibe a
        # imagem da outra — foi assim que a logo grande cobriu a autorização.
        if para in dono and dono[para] != nome_aba:
            novo = _proximo_desenho(conteudo)
            registrador.warning(
                "as abas %r e %r saíram apontando para %s; separando %r em %s",
                dono[para], nome_aba, Path(para).name, nome_aba, Path(novo).name,
            )
            if not _reapontar_desenho(conteudo, arquivo_gerado, para, novo):
                continue
            para = novo
            novos_desenhos.add(para)

        dono[para] = nome_aba
        if conteudo.get(para) == extras[de]:
            continue  # já é o original

        conteudo[para] = extras[de]
        trocados.append(para)

        rels_de = f"xl/drawings/_rels/{Path(de).name}.rels"
        rels_para = f"xl/drawings/_rels/{Path(para).name}.rels"
        if rels_de in extras:
            conteudo[rels_para] = extras[rels_de]
            trocados.append(rels_para)
    return trocados


def _proximo_desenho(conteudo: dict[str, bytes]) -> str:
    """Primeiro `xl/drawings/drawingN.xml` livre no pacote."""
    usados = {
        int(m.group(1))
        for n in conteudo
        if (m := re.match(r"xl/drawings/drawing(\d+)\.xml$", n))
    }
    n = 1
    while n in usados:
        n += 1
    return f"xl/drawings/drawing{n}.xml"


def _reapontar_desenho(
    conteudo: dict[str, bytes], aba: str, antigo: str, novo: str
) -> bool:
    """
    Faz o `.rels` da aba apontar para outro arquivo de desenho.

    Devolve False sem tocar em nada se o `.rels` não tiver o desenho esperado
    — remendar por adivinhação em pacote OOXML é como se corrompe o arquivo.
    """
    caminho = f"xl/worksheets/_rels/{Path(aba).name}.rels"
    bruto = conteudo.get(caminho)
    if not bruto:
        return False
    try:
        raiz = ET.fromstring(bruto)
    except ET.ParseError:
        return False

    alvo_antigo = Path(antigo).name
    for rel in raiz:
        if rel.get("Type") != _TIPO_DESENHO:
            continue
        if Path((rel.get("Target") or "").replace("\\", "/")).name != alvo_antigo:
            continue
        rel.set("Target", f"../drawings/{Path(novo).name}")
        conteudo[caminho] = ET.tostring(raiz, encoding="utf-8", xml_declaration=True)
        return True
    return False


def _declarar_desenho(tipos: bytes, parte: str) -> bytes:
    """Acrescenta o Override do tipo de desenho para `parte`, se faltar."""
    alvo = "/" + parte.lstrip("/")
    raiz = ET.fromstring(tipos)
    for filho in raiz:
        if filho.get("PartName") == alvo:
            return tipos
    ET.SubElement(
        raiz,
        f"{_NS_TIPOS}Override",
        {"PartName": alvo, "ContentType": _TIPO_CONTEUDO_DESENHO},
    )
    return ET.tostring(raiz, encoding="utf-8", xml_declaration=True)


def restaurar_recursos(modelo: Path, gerado: Path, aba: str | None = None) -> list[str]:
    """
    Devolve ao xlsx gerado as partes que o openpyxl joga fora ao salvar.

    Sem Pillow instalado o openpyxl descarta todas as imagens do arquivo — e a
    logomarca da GOL vive justamente aí. Como o arquivo gerado é uma cópia do
    modelo, dá para reinjetar as partes originais (mídia, desenhos e
    configuração de impressora) sem tocar em nada que o openpyxl escreveu.

    Devolve a lista de partes restauradas. Qualquer imprevisto é registrado e
    o arquivo do openpyxl é mantido como está (vale mais um xlsx sem logo do
    que um xlsx corrompido).
    """
    try:
        with zipfile.ZipFile(modelo) as pacote_modelo:
            nomes = set(pacote_modelo.namelist())
            extras = {
                n: pacote_modelo.read(n)
                for n in nomes
                if n.startswith(_PREFIXOS_RESTAURAVEIS)
            }
            if not extras:
                return []
            abas_modelo = _mapa_abas(pacote_modelo)
            rels_modelo = {
                n: pacote_modelo.read(n)
                for n in nomes
                if n.startswith("xl/worksheets/_rels/")
            }
            tipos_modelo = pacote_modelo.read("[Content_Types].xml")
            aba_modelo = abas_modelo.get(aba or "")
            xml_aba_modelo = (
                pacote_modelo.read(aba_modelo)
                if aba_modelo and aba_modelo in nomes
                else None
            )

        with zipfile.ZipFile(gerado) as pacote_gerado:
            conteudo = {n: pacote_gerado.read(n) for n in pacote_gerado.namelist()}
            abas_geradas = _mapa_abas(pacote_gerado)

        faltando = {n: b for n, b in extras.items() if n not in conteudo}
        conteudo.update(faltando)

        # Desenhos são reescritos pelo openpyxl, não descartados — e a
        # reescrita perde propriedades da forma. Por isso vão à parte, sempre.
        novos_desenhos: set[str] = set()
        # `dono` sai daqui preenchido: qual aba ficou com cada arquivo de
        # desenho. O laço mais abaixo precisa dele para não entregar a mesma
        # parte a duas abas.
        dono_do_desenho: dict[str, str] = {}
        redesenhados = _devolver_desenhos_originais(
            conteudo, extras, abas_modelo, abas_geradas, rels_modelo,
            novos_desenhos, dono_do_desenho,
        )

        # Células vazias que só carregavam formatação, idem.
        celulas_devolvidas = 0
        aba_gerada = abas_geradas.get(aba or "")
        if xml_aba_modelo and aba_gerada and aba_gerada in conteudo:
            conteudo[aba_gerada], celulas_devolvidas = _devolver_celulas_de_estilo(
                conteudo[aba_gerada], xml_aba_modelo
            )

        if not faltando and not redesenhados and not celulas_devolvidas:
            return []

        for nome_aba, arquivo_modelo in abas_modelo.items():
            arquivo_gerado = abas_geradas.get(nome_aba)
            if not arquivo_gerado or arquivo_gerado not in conteudo:
                continue
            rels_origem = f"xl/worksheets/_rels/{Path(arquivo_modelo).name}.rels"
            rels_destino = f"xl/worksheets/_rels/{Path(arquivo_gerado).name}.rels"
            if rels_origem not in rels_modelo or rels_destino in conteudo:
                continue

            conteudo[rels_destino] = rels_modelo[rels_origem]
            for rel in ET.fromstring(rels_modelo[rels_origem]):
                if rel.get("Type") != _TIPO_DESENHO:
                    continue
                conteudo[arquivo_gerado] = _inserir_desenho(
                    conteudo[arquivo_gerado], rel.get("Id", "")
                )
                # O .rels veio do modelo e cita o desenho pelo nome que ele
                # tinha LÁ. No pacote novo esse nome pode já pertencer a outra
                # aba — o openpyxl renumera quando descarta algum desenho. Se
                # pertencer, esta aba ganha uma parte só dela.
                de = _caminho_no_pacote(rel.get("Target") or "")
                if not de or de not in extras:
                    continue
                if dono_do_desenho.get(de) not in (None, nome_aba):
                    para = _proximo_desenho(conteudo)
                    registrador.warning(
                        "%r ficaria com o desenho de %r (%s); separando em %s",
                        nome_aba, dono_do_desenho[de], Path(de).name, Path(para).name,
                    )
                    if not _reapontar_desenho(conteudo, arquivo_gerado, de, para):
                        continue
                    novos_desenhos.add(para)
                else:
                    para = de
                conteudo[para] = extras[de]
                dono_do_desenho[para] = nome_aba
                rels_de = f"xl/drawings/_rels/{Path(de).name}.rels"
                if rels_de in extras:
                    conteudo[f"xl/drawings/_rels/{Path(para).name}.rels"] = extras[rels_de]

        conteudo["[Content_Types].xml"] = _mesclar_tipos(
            conteudo["[Content_Types].xml"], tipos_modelo, set(faltando)
        )
        # Desenho criado agora para desfazer colisão não existe no modelo, e
        # portanto não vem em `tipos_modelo`: o override é declarado aqui.
        for parte in sorted(novos_desenhos):
            conteudo["[Content_Types].xml"] = _declarar_desenho(
                conteudo["[Content_Types].xml"], parte
            )

        # Antes de trocar o arquivo bom por um remendado, confere que todo XML
        # que tocamos continua bem formado.
        for nome, dados in conteudo.items():
            if nome.endswith(".xml") or nome.endswith(".rels"):
                ET.fromstring(dados)

        temporario = gerado.with_suffix(gerado.suffix + ".tmp")
        with zipfile.ZipFile(temporario, "w", zipfile.ZIP_DEFLATED) as saida:
            for nome, dados in conteudo.items():
                saida.writestr(nome, dados)
        temporario.replace(gerado)

    except Exception as erro:  # noqa: BLE001 — reparo é best-effort
        registrador.warning(
            "não foi possível restaurar imagens/desenhos em %s (%s) — o arquivo "
            "continua válido, mas pode sair sem a logomarca.",
            gerado.name, erro,
        )
        return []

    total = sorted(set(faltando) | set(redesenhados))
    registrador.info(
        "%s: reparo pós-openpyxl — %d parte(s) recolocada(s), %d desenho(s) "
        "devolvido(s) ao original, %d célula(s) de formatação recriada(s).",
        gerado.name, len(faltando), len(redesenhados), celulas_devolvidas,
    )
    if celulas_devolvidas:
        total.append(f"{celulas_devolvidas} célula(s) de formatação")
    return total


# --------------------------------------------------------------------------- #
# Preenchimento
# --------------------------------------------------------------------------- #


def _resolver_modelo(conta: Conta, amb: Ambiente) -> Path:
    if not conta.modelo_base:
        raise ErroAutorizacao(
            f"conta {conta.id!r} não tem 'modelo_base' no registro de fornecedores — "
            f"sem ele não há xlsx do mês anterior para copiar."
        )
    bruto = Path(conta.modelo_base)
    modelo = bruto if bruto.is_absolute() else (amb.caminhos.autorizacoes / bruto)
    if not modelo.is_file():
        raise ErroAutorizacao(f"modelo base não encontrado: {modelo}")
    return modelo


def _resolver_destino(
    conta: Conta,
    competencia: Competencia,
    destino: Path | str | None,
    amb: Ambiente,
) -> Path:
    nome = conta.nome_base_arquivo(competencia) + ".xlsx"
    if destino is None:
        return amb.pasta_trabalho(conta.id, competencia) / nome
    alvo = Path(destino)
    if alvo.is_dir() or not alvo.suffix:
        alvo = alvo / nome
    if alvo.suffix.lower() != ".xlsx":
        raise ErroAutorizacao(f"destino precisa terminar em .xlsx: {alvo}")
    return alvo


def _normalizar_vencimento(vencimento: Any) -> date | None:
    if vencimento is None:
        return None
    if isinstance(vencimento, datetime):
        return vencimento.date()
    if isinstance(vencimento, date):
        return vencimento
    if isinstance(vencimento, str):
        texto = vencimento.strip()
        for molde in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
            try:
                return datetime.strptime(texto, molde).date()
            except ValueError:
                continue
    raise ErroAutorizacao(f"vencimento não reconhecido: {vencimento!r}")


def preencher(
    conta: Conta,
    competencia: Competencia,
    *,
    valor: float | None,
    vencimento: date | datetime | str | None,
    numero_documento: str | None = None,
    descricao: str | None = None,
    pagante: str | None = None,
    beneficiario: str | None = None,
    meio_pgto: str | None = None,
    forma_pgto: str | None = None,
    destino: Path | str | None = None,
    ambiente_: Ambiente | None = None,
) -> ResultadoEtapa:
    """
    Gera a autorização de pagamento de `conta` para `competencia`.

    Copia o `modelo_base` (xlsx do mês anterior) para a área local de
    trabalho com o nome do mês novo e reescreve apenas os campos de valor.

    `detalhes` traz:
      - `alteracoes`: `{coordenada: (valor_antigo, valor_novo)}` de tudo que
        mudou — é o que o painel mostra antes de publicar;
      - `campos`: `{campo: valor}` já resolvido;
      - `avisos`: pendências para olho humano (valor faltando, VLOOKUP quebrado…);
      - `protegidas`: campos pulados por haver fórmula na célula;
      - `modelo`, `aba`, `arquivo`.
    """
    amb = ambiente_ or ambiente()
    dados = conta.autorizacao
    avisos: list[str] = []

    try:
        modelo = _resolver_modelo(conta, amb)
        arquivo = _resolver_destino(conta, competencia, destino, amb)
    except ErroAutorizacao as erro:
        registrador.error("%s: %s", conta.id, erro)
        return ResultadoEtapa.erro(Etapa.AUTORIZACAO, str(erro))

    if arquivo.exists() and arquivo.resolve() == modelo.resolve():
        return ResultadoEtapa.erro(
            Etapa.AUTORIZACAO,
            f"destino é o próprio modelo base ({arquivo}) — a geração sobrescreveria "
            f"o arquivo do mês anterior.",
        )

    # Trava de segurança: nada é escrito fora de dados/ sem permissão.
    try:
        amb.exigir_permissao(arquivo)
    except PermissionError as erro:
        registrador.error("%s: %s", conta.id, erro)
        return ResultadoEtapa.erro(Etapa.AUTORIZACAO, str(erro))

    arquivo.parent.mkdir(parents=True, exist_ok=True)

    try:
        shutil.copy2(modelo, arquivo)
    except OSError as erro:
        mensagem = f"não foi possível copiar o modelo para {arquivo}: {erro}"
        registrador.error("%s: %s", conta.id, mensagem)
        return ResultadoEtapa.erro(Etapa.AUTORIZACAO, mensagem)

    registrador.info("%s: modelo %s copiado para %s", conta.id, modelo.name, arquivo)

    try:
        wb = openpyxl.load_workbook(arquivo)  # sem data_only: preserva fórmulas
    except Exception as erro:  # noqa: BLE001 — openpyxl lança de tudo
        mensagem = f"não foi possível abrir {arquivo.name}: {erro}"
        registrador.error("%s: %s", conta.id, mensagem)
        return ResultadoEtapa.erro(Etapa.AUTORIZACAO, mensagem, artefatos=[arquivo])

    if conta.aba_autorizacao not in wb.sheetnames:
        wb.close()
        mensagem = (
            f"a aba {conta.aba_autorizacao!r} não existe em {arquivo.name}. "
            f"Abas disponíveis: {wb.sheetnames}. Corrija 'aba_autorizacao' no "
            f"registro da conta {conta.id!r}."
        )
        registrador.error("%s: %s", conta.id, mensagem)
        return ResultadoEtapa.erro(Etapa.AUTORIZACAO, mensagem, artefatos=[arquivo])

    ws = wb[conta.aba_autorizacao]

    # ---- valores a escrever ------------------------------------------------ #
    venc = _normalizar_vencimento(vencimento)
    if venc is None and conta.vencimento_dia:
        venc = competencia.dia_vencimento(int(conta.vencimento_dia))
        avisos.append(
            f"vencimento não informado — usado o dia {conta.vencimento_dia} da "
            f"competência ({venc:%d/%m/%Y}). Confira."
        )
    elif venc is None:
        avisos.append("vencimento não informado — a célula H7 ficou como estava.")

    if descricao is not None:
        # Texto revisado no painel: entra como está, sem reprocessar o mês.
        texto_descricao, faltando = descricao, set()
    else:
        texto_descricao, como = sugerir_descricao(conta, competencia)
        faltando = set()
        if como:
            avisos.append(f"observações ajustadas para {como}")

    if faltando:
        avisos.append(
            "placeholders sem valor na descrição, mantidos literais: "
            + ", ".join(sorted(faltando))
        )

    valor_ok = valor is not None and float(valor) > 0
    if not valor_ok:
        avisos.append(
            "valor da fatura ausente ou inválido — VALOR ORIGINAL (C22) não foi "
            "escrito. Informe o valor antes de publicar."
        )

    # Escolhas feitas no painel valem sobre o registro da conta.
    nome_pagante = pagante or dados.pagante
    nome_beneficiario = beneficiario or dados.beneficiario

    campos: dict[str, Any] = {
        "departamento": dados.departamento,
        "pagante": nome_pagante,
        "vencimento": venc,
        "beneficiario": nome_beneficiario,
        "forma_pgto": forma_pgto or dados.forma_pgto,
        "meio_pgto": meio_pgto or dados.meio_pgto,
        "valor": float(valor) if valor_ok else None,
        "setor": dados.setor,
        "motivo": dados.motivo,
        "regional": dados.regional,
        "rotulo_documento": dados.rotulo_documento,
        "numero_documento": numero_documento,
        "centro_custo": dados.centro_custo,
        "natureza": dados.natureza,
        "cooperativa": dados.cooperativa,
        "descricao": texto_descricao,
        "pgto_previsto": dados.pgto_previsto,
    }

    alteracoes: dict[str, tuple[Any, Any]] = {}
    protegidas: dict[str, str] = {}
    escritas: dict[str, Any] = {}

    for campo, conteudo in campos.items():
        if conteudo is None or conteudo == "":
            continue
        coordenada = CELULAS[campo]
        try:
            mudanca = escrever(
                ws,
                coordenada,
                conteudo,
                permitir_formula=campo in CAMPOS_VALOR,
            )
        except CelulaProtegida as erro:
            protegidas[campo] = str(erro)
            avisos.append(f"campo {campo!r} não escrito: {erro}")
            registrador.warning("%s: %s", conta.id, erro)
            continue
        escritas[ancora_mesclada(ws, coordenada)] = conteudo
        if mudanca is not None:
            alteracoes[ancora_mesclada(ws, coordenada)] = mudanca

    # Pagante do cadastro: se a tabela do template não conhece o nome atual,
    # o CNPJ entra no lugar da fórmula em vez de virar #N/D.
    socorros: dict[str, dict[int | None, str]] = {}
    registrado = pagante_por_nome(nome_pagante)
    if registrado and registrado.cnpj:
        socorros[ancora_mesclada(ws, CELULAS["pagante"])] = {None: registrado.cnpj}

    # Mesma história do lado de quem recebe, e com mais campos: trocar o
    # beneficiário por um que a tabela do template não conhece deixava CNPJ,
    # contato, celular e banco em #N/D. Os dados estão no cadastro; entram no
    # lugar das fórmulas que não teriam como resolver.
    recebedor = beneficiario_por_nome(nome_beneficiario)
    if recebedor:
        por_coluna = {
            coluna: str(getattr(recebedor, campo))
            for coluna, campo in COLUNAS_BENEFICIARIOS.items()
            if getattr(recebedor, campo, None)
        }
        if por_coluna:
            socorros[ancora_mesclada(ws, CELULAS["beneficiario"])] = por_coluna

    if socorros:
        avisos.extend(suprir_vlookup_quebrado(wb, ws, escritas, socorros))

    avisos.extend(conferir_vlookups(wb, ws, escritas))

    # Garante que o Excel recalcule tudo ao abrir (VLOOKUP, TODAY, SUM).
    try:
        wb.calculation.fullCalcOnLoad = True
    except AttributeError:  # pragma: no cover — workbook sem calcPr
        registrador.debug("workbook sem CalcProperties; recálculo fica a cargo do Excel")

    try:
        wb.save(arquivo)
    except OSError as erro:
        wb.close()
        mensagem = (
            f"não foi possível salvar {arquivo.name}: {erro}. "
            f"O arquivo pode estar aberto no Excel."
        )
        registrador.error("%s: %s", conta.id, mensagem)
        return ResultadoEtapa.erro(Etapa.AUTORIZACAO, mensagem, artefatos=[arquivo])
    finally:
        wb.close()

    # O openpyxl salva sem as imagens (a logomarca some do cabeçalho);
    # devolvemos essas partes ao pacote.
    restauradas = restaurar_recursos(modelo, arquivo, conta.aba_autorizacao)

    detalhes = {
        "alteracoes": alteracoes,
        "partes_restauradas": restauradas,
        "campos": {k: v for k, v in campos.items() if v not in (None, "")},
        "celulas": dict(CELULAS),
        "avisos": avisos,
        "protegidas": protegidas,
        "modelo": str(modelo),
        "aba": conta.aba_autorizacao,
        "arquivo": str(arquivo),
    }

    resumo = f"{len(alteracoes)} célula(s) atualizada(s) em {arquivo.name}"
    registrador.info("%s: %s", conta.id, resumo)

    if not valor_ok or avisos:
        return ResultadoEtapa.atencao(
            Etapa.AUTORIZACAO,
            f"{resumo}; {len(avisos)} ponto(s) para conferir.",
            detalhes=detalhes,
            artefatos=[arquivo],
        )
    return ResultadoEtapa.sucesso(
        Etapa.AUTORIZACAO, resumo, detalhes=detalhes, artefatos=[arquivo]
    )


def ler_campos(caminho_xlsx: Path | str, aba: str) -> dict:
    """
    Lê de volta o que ficou preenchido, para o painel exibir.

    Devolve `{campo: valor}` usando o mapa `CELULAS`. Fórmulas são devolvidas
    como texto (`'=VLOOKUP(...)'`) — o valor calculado só existe depois que o
    Excel abre o arquivo.
    """
    caminho = Path(caminho_xlsx)
    if not caminho.is_file():
        raise ErroAutorizacao(f"planilha não encontrada: {caminho}")

    wb = openpyxl.load_workbook(caminho)
    try:
        if aba not in wb.sheetnames:
            raise ErroAutorizacao(
                f"a aba {aba!r} não existe em {caminho.name}. "
                f"Abas disponíveis: {wb.sheetnames}."
            )
        ws = wb[aba]
        return {
            campo: _simplificar(ws[ancora_mesclada(ws, coordenada)].value)
            for campo, coordenada in CELULAS.items()
        }
    finally:
        wb.close()


# --------------------------------------------------------------------------- #
# Teste manual
# --------------------------------------------------------------------------- #

if __name__ == "__main__":  # pragma: no cover
    import sys

    from automacao.nucleo.config import DIR_CONFIG, carregar_contas

    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")

    conta_id = sys.argv[1] if len(sys.argv) > 1 else "blip"
    comp = Competencia.de_texto(sys.argv[2]) if len(sys.argv) > 2 else Competencia.atual()

    registro = DIR_CONFIG / "fornecedores.yaml"
    if not registro.is_file():
        registro = DIR_CONFIG / "fornecedores.gerado.yaml"
    alvo = next(c for c in carregar_contas(registro) if c.id == conta_id)

    saida = preencher(
        alvo,
        comp,
        valor=1234.56,
        vencimento=comp.dia_vencimento(alvo.vencimento_dia or 10),
        numero_documento="000000",
    )
    registrador.info("situação: %s — %s", saida.situacao.value, saida.mensagem)
    for coord, (antes, depois) in saida.detalhes.get("alteracoes", {}).items():
        registrador.info("  %s: %r -> %r", coord, antes, depois)
    for aviso in saida.detalhes.get("avisos", []):
        registrador.warning("  ! %s", aviso)
