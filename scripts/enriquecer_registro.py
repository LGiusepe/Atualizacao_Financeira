"""
Enriquece config/fornecedores.gerado.yaml com o que só a planilha sabe.

Preenche, para cada conta:
  * planilha_contas.chaves  — o texto exato da coluna FORNECEDOR, para o
    checklist saber qual linha pintar de verde;
  * vencimento_dia          — a coluna VENCIMENTO;
  * ativo                   — false no que está DESCONTINUADO / cancelado;
  * email.eh_operadora e email.cidade — nas contas de operadora, para o corpo
    do e-mail citar a unidade atendida.

Casamento incerto (< 0,80) fica marcado em 'revisar' em vez de virar chute.
Nada é escrito na planilha nem no OneDrive.

Saída: config/fornecedores.enriquecido.yaml
"""

from __future__ import annotations

import re
import sys
import unicodedata
from datetime import date, datetime
from difflib import SequenceMatcher
from pathlib import Path

import yaml

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from automacao.entrega import planilha_contas  # noqa: E402

from automacao.nucleo.config import ambiente

RAIZ = Path(__file__).resolve().parents[1]
ENTRADA = RAIZ / "config" / "fornecedores.gerado.yaml"
SAIDA = RAIZ / "config" / "fornecedores.enriquecido.yaml"
# Lido do settings.yaml: o caminho real da empresa mora lá, não aqui.
RAIZ_PAGAMENTOS = ambiente().caminhos.autorizacoes

DADOS_LOCAIS = RAIZ / "config" / "dados-locais.yaml"


def _locais() -> dict:
    """
    As tabelas deste script, lidas de `config/dados-locais.yaml`.

    Elas nomeiam os fornecedores da empresa e a razão social de vários deles.
    Ficavam escritas aqui; saíram quando o repositório virou público. Ausente
    o arquivo, o script roda com tabelas vazias: o casamento por semelhança
    continua funcionando, só perde os apelidos declarados à mão.

    O modelo está em `config/dados-locais.exemplo.yaml`.
    """
    if not DADOS_LOCAIS.is_file():
        return {}
    bruto = yaml.safe_load(DADOS_LOCAIS.read_text(encoding="utf-8")) or {}
    return bruto.get("enriquecer_registro") or {}


_LOCAIS = _locais()

# Fornecedores cujo e-mail precisa citar a unidade/cidade atendida.
OPERADORAS = set(_LOCAIS.get("operadoras") or ())

# Contas que saem num e-mail só, com um anexo por fatura: há fornecedor com
# quatro faturas no mês cujo financeiro recebe tudo numa mensagem única.
# Para agrupar outro, basta acrescentar o nome da pasta ao arquivo local.
GRUPOS_EMAIL = dict(_LOCAIS.get("grupos_email") or {})

LIMIAR_BOM = 0.80
LIMIAR_MINIMO = 0.55

# Situações da coluna VALOR que significam "não processar mais".
SITUACOES_INATIVAS = {"descontinuado", "cancelado"}


def sem_acento(texto: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFKD", texto or "")
        if not unicodedata.combining(c)
    )


def normalizar(texto: str) -> str:
    return re.sub(r"\s+", " ", sem_acento(texto).lower()).strip()


def so_digitos(texto: str) -> str:
    return re.sub(r"\D", "", texto or "")


def escore(a: str, b: str) -> float:
    return SequenceMatcher(None, normalizar(a), normalizar(b)).ratio()


def cidade_de(subunidade: str | None) -> str | None:
    """
    Rótulo legível da unidade atendida, para o corpo do e-mail.

    'Campinas'                 -> 'Campinas'
    'Móvel - 000000000'        -> 'Móvel'
    'COOPERATIVA - 000000000'  -> 'Cooperativa'
    'DDG EXECUTIVO'            -> 'DDG Executivo'   (mantém a sigla)
    'CPD'                      -> 'CPD'
    """
    if not subunidade:
        return None
    limpo = re.split(r"\s*-\s*\d", subunidade)[0].strip(" -")
    limpo = re.sub(r"\s*\d{6,}\s*$", "", limpo).strip()
    if not limpo or not limpo.isupper():
        return limpo or None

    # Tudo em caixa alta. Palavra sem vogal é sigla e continua em caixa alta
    # (DDG, CPD, MG); o resto vira capitalizado (VOZ -> Voz, LINK -> Link).
    vogais = set("AEIOUÁÂÃÀÉÊÍÓÔÕÚ")
    return " ".join(
        p if not (set(p) & vogais) else p.capitalize() for p in limpo.split()
    )


# Casos que semelhança de texto nunca acerta: o nome da pasta e o nome do
# fornecedor na planilha não têm parentesco, ou duas linhas têm texto idêntico.
# Valor pode ser o texto da coluna FORNECEDOR ou o NÚMERO DA LINHA no Excel
# (necessário quando duas linhas têm exatamente o mesmo texto — acontece com
# fornecedor que vende dois produtos e repete a razão social nas duas linhas).
#
# A tabela é razão social de verdade, então mora no arquivo local.
APELIDOS: dict[str, str | int] = dict(_LOCAIS.get("apelidos") or {})

# Palavras que aparecem em quase todo nome e não ajudam a distinguir.
RUIDO = {
    "ltda", "sa", "s", "a", "eireli", "me", "epp", "do", "da", "de", "dos",
    "das", "e", "em", "servicos", "solucoes", "tecnologia", "sistemas",
    "informatica", "brasil", "software",
}


def fichas(texto: str) -> set[str]:
    """Palavras significativas de um nome, sem ruído e sem números."""
    palavras = re.findall(r"[a-z0-9]+", normalizar(texto))
    return {p for p in palavras if p not in RUIDO and not p.isdigit() and len(p) > 1}


def numeros(texto: str) -> set[str]:
    """Números de conta/linha (6+ dígitos), que são o discriminador real."""
    return set(re.findall(r"\d{6,}", so_digitos(texto)))


# Pasta sem arquivo novo há muito tempo é conta que parou: cancelada,
# descontinuada ou aguardando renovação. Como o ciclo é mensal, cada 30 dias
# parados é um mês sem fatura.
DIAS_PARA_INATIVAR = 120   # 4 ciclos perdidos — trata como inativa
DIAS_PARA_ALERTAR = 65     # 2 ciclos perdidos — só sinaliza


class PastaSumiu(Exception):
    """A pasta do fornecedor não existe mais — provável renomeação."""


def dias_sem_movimento(
    conta: dict, raiz_pagamentos: Path
) -> tuple[int | None, str | None, int | None]:
    """
    Há quantos dias nada é gravado na pasta desta conta.

    Devolve `(dias_da_conta, data_do_último, dias_do_fornecedor_inteiro)`.

    O terceiro valor é rede de segurança. Se a sub-unidade parece parada mas o
    fornecedor está recebendo arquivo, o mais provável é que a sub-unidade
    tenha sido reorganizada — não que a conta tenha acabado. Foi o que
    aconteceu com um fornecedor cujos arquivos saíram da subpasta com o nome
    dele e passaram a ficar soltos no mês: a conta viva foi dada como morta.
    """
    pasta = raiz_pagamentos / conta["pasta"]
    if not pasta.is_dir():
        raise PastaSumiu(conta["pasta"])

    subunidade = conta.get("subunidade")
    recente_conta = 0.0
    recente_pasta = 0.0
    for arquivo in pasta.rglob("*"):
        if not arquivo.is_file():
            continue
        try:
            quando = arquivo.stat().st_mtime
        except OSError:
            continue
        recente_pasta = max(recente_pasta, quando)
        if not subunidade or subunidade in str(arquivo):
            recente_conta = max(recente_conta, quando)

    hoje = date.today()
    dias_pasta = (
        (hoje - datetime.fromtimestamp(recente_pasta).date()).days
        if recente_pasta
        else None
    )
    if not recente_conta:
        return None, None, dias_pasta

    ultima = datetime.fromtimestamp(recente_conta).date()
    return (hoje - ultima).days, ultima.isoformat(), dias_pasta


def numero_de_contrato(conta: dict) -> str:
    """
    Número do contrato/linha guardado na célula C29 da última autorização.

    Só serve quando é mesmo o número do contrato. Quando o texto diz "FATURA",
    aquilo é o número do documento daquele mês — muda toda competência e, se
    usado para casar, atrapalha (foi o que jogou uma conta de operadora na
    linha errada). Nesses casos devolvemos vazio.
    """
    bruto = str((conta.get("referencia") or {}).get("ultimo_numero_documento") or "")
    return "" if "fatura" in normalizar(bruto) else bruto


def afinidade(conta: dict, linha) -> tuple[float, str]:
    """
    Quanto esta conta combina com esta linha da planilha, de 0 a 1.

    O número da conta manda: é o que separa entre si as quatro contas irmãs de
    um fornecedor, e as seis de outro. Quando os dois lados têm número e eles
    não batem, o par é
    descartado — não adianta o nome ser parecido, é outra linha.
    """
    pasta = conta.get("pasta") or ""
    subunidade = conta.get("subunidade") or ""
    nome_conta = f"{pasta} {subunidade}".strip()

    apelido = APELIDOS.get(conta["id"])
    if apelido is not None:
        acertou = (
            apelido == linha.linha
            if isinstance(apelido, int)
            else normalizar(str(apelido)) == normalizar(linha.fornecedor)
        )
        return (1.0, "vínculo declarado na tabela de apelidos") if acertou else (0.0, "")

    num_conta = numeros(f"{pasta} {subunidade}") | numeros(numero_de_contrato(conta))
    num_linha = numeros(linha.fornecedor)

    if num_conta & num_linha:
        return 1.0, f"número {sorted(num_conta & num_linha)[0]} bate"

    # Penalidade, e não corte, porque o número em C29 às vezes é o da FATURA
    # do mês (que muda) e não o do contrato — descartar o par por isso
    # eliminava casamentos corretos, como os das contas de operadora.
    if num_conta and num_linha:
        peso = 0.40
        nota_numero = " (números não batem)"
    elif num_conta or num_linha:
        peso = 0.85
        nota_numero = " (só um dos lados tem número)"
    else:
        peso = 1.0
        nota_numero = ""

    fichas_conta = fichas(nome_conta)
    fichas_linha = fichas(linha.fornecedor)
    if not fichas_conta or not fichas_linha:
        return 0.0, ""

    comuns = fichas_conta & fichas_linha
    jaccard = len(comuns) / len(fichas_conta | fichas_linha)
    cobertura = len(comuns) / min(len(fichas_conta), len(fichas_linha))
    textual = escore(nome_conta, linha.fornecedor)

    valor = (0.45 * cobertura + 0.30 * jaccard + 0.25 * textual) * peso
    if not comuns:
        valor = min(valor, 0.35)

    motivo = (
        f"palavras em comum: {', '.join(sorted(comuns))}{nota_numero}"
        if comuns
        else f"só semelhança textual{nota_numero}"
    )
    return valor, motivo


def casar_tudo(contas: list[dict], linhas) -> dict[str, tuple[object, float, str]]:
    """
    Atribuição global: a melhor combinação conta↔linha no conjunto todo.

    Feito de uma vez, e não conta por conta, porque cada linha da planilha só
    pode pertencer a uma conta. Resolvendo isoladamente, as quatro contas
    irmãs de um mesmo fornecedor reivindicavam todas a mesma linha.
    """
    pares = []
    for conta in contas:
        for linha in linhas:
            valor, motivo = afinidade(conta, linha)
            if valor >= LIMIAR_MINIMO:
                pares.append((valor, conta["id"], linha, motivo))

    pares.sort(key=lambda p: -p[0])

    resultado: dict[str, tuple[object, float, str]] = {}
    linhas_usadas: set[int] = set()
    for valor, conta_id, linha, motivo in pares:
        if conta_id in resultado or linha.linha in linhas_usadas:
            continue
        resultado[conta_id] = (linha, valor, motivo)
        linhas_usadas.add(linha.linha)
    return resultado


def main() -> None:
    if not ENTRADA.is_file():
        raise SystemExit(
            f"{ENTRADA.name} não existe. Rode antes: "
            "python scripts/gerar_registro_fornecedores.py"
        )

    doc = yaml.safe_load(ENTRADA.read_text(encoding="utf-8"))
    linhas = planilha_contas.ler_contas_fixas()
    casamentos = casar_tudo(doc["contas"], linhas)
    usadas: dict[int, str] = {}

    casadas = duvidosas = 0

    for conta in doc["contas"]:
        conta.setdefault("revisar", [])
        linha, valor, motivo = casamentos.get(conta["id"], (None, 0.0, ""))

        if linha is None:
            conta["revisar"].append(
                "não achei a linha correspondente em CONTAS FIXAS — preencha "
                "'planilha_contas.chaves' com o texto exato da coluna FORNECEDOR, "
                "senão o checklist não será marcado."
            )
        else:
            conta["planilha_contas"] = {"chaves": [linha.fornecedor]}
            conta["vencimento_dia"] = (
                int(linha.vencimento)
                if isinstance(linha.vencimento, (int, float))
                else None
            )
            situacao = normalizar(linha.situacao)
            if situacao in SITUACOES_INATIVAS:
                conta["ativo"] = False
                conta["observacao"] = f"{linha.situacao} na planilha CONTAS FIXAS"
            elif situacao:
                conta["observacao"] = f"situação na planilha: {linha.situacao}"

            if valor >= LIMIAR_BOM:
                casadas += 1
            else:
                duvidosas += 1
                conta["revisar"].append(
                    f"casei com a linha {linha.linha} ({linha.fornecedor!r}) por "
                    f"{motivo}, mas a semelhança é só {valor:.0%}. Confirme."
                )

            anterior = usadas.get(linha.linha)
            if anterior:
                conta["revisar"].append(
                    f"a linha {linha.linha} da planilha já foi atribuída a "
                    f"'{anterior}'. Duas contas não podem apontar para a mesma linha."
                )
            usadas[linha.linha] = conta["id"]

        # Pasta parada = conta que provavelmente acabou. Sinal independente da
        # planilha, e que pega o que a planilha ainda não marcou.
        try:
            dias, ultima, dias_pasta = dias_sem_movimento(conta, RAIZ_PAGAMENTOS)
        except PastaSumiu:
            conta["ativo"] = False
            conta["referencia"]["dias_sem_movimento"] = None
            conta["revisar"].append(
                f"a pasta '{conta['pasta']}' não existe mais no OneDrive — "
                "foi renomeada, movida ou excluída depois que este registro foi "
                "gerado. Rode 'python scripts/gerar_registro_fornecedores.py' de "
                "novo para reler a estrutura atual."
            )
            conta["email"] = {
                "eh_operadora": normalizar(conta["pasta"]) in OPERADORAS,
                "cidade": None,
                "nome_empresa": conta.get("pasta"),
                "grupo": GRUPOS_EMAIL.get(normalizar(conta["pasta"])),
            }
            continue

        conta.setdefault("referencia", {})
        conta["referencia"]["dias_sem_movimento"] = dias
        conta["referencia"]["ultimo_arquivo_em"] = ultima

        # Fornecedor recebendo arquivo enquanto a sub-unidade parece parada =
        # reorganização de pasta, não conta encerrada. Nunca inativar nesse caso.
        pasta_viva = (
            conta.get("subunidade")
            and dias_pasta is not None
            and dias_pasta < DIAS_PARA_ALERTAR
        )

        if dias is None:
            conta["ativo"] = False
            conta["revisar"].append(
                "a pasta existe mas está vazia — nunca houve movimento. "
                "Marquei como inativa."
            )
        elif pasta_viva and dias >= DIAS_PARA_ALERTAR:
            conta["revisar"].append(
                f"a sub-pasta '{conta['subunidade']}' está sem arquivo novo há "
                f"{dias} dias, mas a pasta '{conta['pasta']}' recebeu arquivo há "
                f"{dias_pasta} dias. Isso costuma ser reorganização de pasta, não "
                "conta encerrada — por isso NÃO inativei. Confira se a "
                "sub-unidade ainda existe com esse nome."
            )
        elif dias >= DIAS_PARA_INATIVAR and conta.get("ativo"):
            conta["ativo"] = False
            conta["revisar"].append(
                f"sem nenhum arquivo novo há {dias} dias (último em {ultima}) — "
                f"são {dias // 30} ciclos mensais parados. Marquei como INATIVA. "
                "Se a conta continua viva, mude 'ativo' para true."
            )
        elif dias >= DIAS_PARA_ALERTAR and conta.get("ativo"):
            conta["revisar"].append(
                f"sem arquivo novo há {dias} dias (último em {ultima}). "
                "Confira se não foi cancelada ou está aguardando renovação."
            )

        # Operadora: o corpo do e-mail precisa dizer qual unidade é atendida.
        eh_operadora = normalizar(conta["pasta"]) in OPERADORAS
        conta["email"] = {
            "eh_operadora": eh_operadora,
            "cidade": cidade_de(conta.get("subunidade")) if eh_operadora else None,
            "nome_empresa": conta.get("pasta"),
            # Contas com o mesmo grupo saem num e-mail só, um anexo por fatura.
            "grupo": GRUPOS_EMAIL.get(normalizar(conta["pasta"])),
        }
        if eh_operadora and not conta["email"]["cidade"]:
            conta["revisar"].append(
                "é conta de operadora mas não consegui deduzir a cidade/unidade — "
                "preencha 'email.cidade'."
            )

    doc["_meta"]["enriquecido_por"] = "scripts/enriquecer_registro.py"
    with SAIDA.open("w", encoding="utf-8") as fh:
        yaml.safe_dump(doc, fh, allow_unicode=True, sort_keys=False, width=120)

    orfas = [linha for linha in linhas if linha.linha not in usadas and linha.ativa]
    pendentes = [c for c in doc["contas"] if c["revisar"]]

    print(f"{len(doc['contas'])} contas · {casadas} casadas com segurança · "
          f"{duvidosas} duvidosas · {len(pendentes)} precisam de revisão\n")

    print(f"{'ID':<36} {'LINHA':>5} {'VENC':>4} {'ATIVO':>5}  FORNECEDOR NA PLANILHA")
    print("-" * 112)
    for c in doc["contas"]:
        chaves = (c.get("planilha_contas") or {}).get("chaves") or []
        print(
            f"{'!' if c['revisar'] else ' '}{c['id']:<35} "
            f"{'—':>5} {str(c.get('vencimento_dia') or '—'):>4} "
            f"{'sim' if c.get('ativo') else 'NÃO':>5}  {(chaves[0] if chaves else '— sem vínculo —')[:44]}"
        )

    if orfas:
        print(f"\n{len(orfas)} linha(s) ATIVAS da planilha sem pasta correspondente:")
        for linha in orfas:
            print(f"   linha {linha.linha:>2}  {linha.fornecedor}")

    print(f"\nSalvo em {SAIDA}")
    print("Revise e depois: copy config\\fornecedores.enriquecido.yaml config\\fornecedores.yaml")


if __name__ == "__main__":
    main()
