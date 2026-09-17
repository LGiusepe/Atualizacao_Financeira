"""
Gera o rascunho de config/fornecedores.yaml a partir da base real (SOMENTE LEITURA).

Para cada fornecedor (e cada sub-unidade: cidade / linha / conta), abre a
Autorização de Pagamento mais recente já preenchida e extrai os valores que
se repetem todo mês (pagante, beneficiário, centro de custo, natureza,
descrição, etc). Esses valores viram os defaults da automação.

O YAML gerado é um RASCUNHO para revisão humana — nada é aplicado sozinho.

Saída: config/fornecedores.gerado.yaml
"""

from __future__ import annotations

import re
import sys
import unicodedata
import warnings
from datetime import date, datetime
from pathlib import Path

import openpyxl
import yaml

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
sys.stdout.reconfigure(encoding="utf-8")

from automacao.nucleo.config import ambiente

RAIZ = Path(__file__).resolve().parents[1]
# Lido do settings.yaml: o caminho real da empresa mora lá, não aqui.
ONEDRIVE = ambiente().caminhos.onedrive
RAIZ_PAGAMENTOS = ONEDRIVE / "Autorizações de pagamento"
PLANILHA_CONTAS = ONEDRIVE / "CONTAS E ACESSOS.xlsx"
SAIDA = RAIZ / "config" / "fornecedores.gerado.yaml"

PADROES_MES = [
    ("MM-AAAA", re.compile(r"^(?P<m>\d{2})\s*[-_]\s*(?P<a>\d{4})$")),
    ("AAAA-MM", re.compile(r"^(?P<a>\d{4})\s*[-_]\s*(?P<m>\d{2})$")),
]

# Onde cada informação mora na aba de autorização.
# Vale para as duas variantes de template (MODELO antigo e AUTORIZAÇÃO novo).
CELULAS = {
    "departamento": "B6",
    "pagante": "B7",
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
    "vencimento": "H7",
}

ABAS_AUTORIZACAO = ("AUTORIZAÇÃO", "MODELO", "AUTORIZACAO")


def sem_acento(texto: str) -> str:
    return "".join(
        c
        for c in unicodedata.normalize("NFKD", texto)
        if not unicodedata.combining(c)
    )


def gerar_id(*partes: str) -> str:
    bruto = "-".join(p for p in partes if p)
    bruto = sem_acento(bruto).lower()
    bruto = re.sub(r"[^a-z0-9]+", "-", bruto).strip("-")
    return re.sub(r"-{2,}", "-", bruto)


def classificar_mes(nome: str) -> tuple[str, int, int] | None:
    limpo = nome.strip()
    for padrao, rx in PADROES_MES:
        m = rx.match(limpo)
        if not m:
            continue
        ano, mes = int(m.group("a")), int(m.group("m"))
        if 1 <= mes <= 12:
            return padrao, ano, mes
    return None


def detectar_subunidades(pasta_mes: Path) -> list[str] | None:
    """
    Diz se este mês está dividido em sub-unidades (cidade / linha / serviço).

    O sinal é onde mora a autorização, e não a mera existência de subpasta:

    * autorização SOLTA no mês  -> conta única, mesmo que haja subpasta ao lado
      (há fornecedor com uma pasta `recarga` ao lado que é só anexo);
    * autorização DENTRO das subpastas -> cada subpasta é uma conta
      (o arranjo das operadoras).

    Devolve `None` quando o mês ainda não tem autorização nenhuma — aí quem
    decide é o mês anterior. Conferido contra oito fornecedores da base, de
    arranjos diferentes: acerta os oito.
    """
    if not pasta_mes.is_dir():
        return None

    if encontrar_xlsx(pasta_mes) is not None:
        return []  # autorização no próprio mês: conta única

    com_autorizacao = sorted(
        p.name
        for p in pasta_mes.iterdir()
        if p.is_dir() and encontrar_xlsx(p) is not None
    )
    return com_autorizacao or None


def encontrar_xlsx(pasta: Path) -> Path | None:
    candidatos = [
        p
        for p in pasta.glob("*.xlsx")
        if not p.name.startswith("~$")
    ]
    return candidatos[0] if candidatos else None


def ler_autorizacao(caminho: Path) -> dict:
    """Extrai os campos fixos da autorização. Ignora fórmulas (pega o cache)."""
    dados: dict = {}
    try:
        wb = openpyxl.load_workbook(caminho, data_only=True)
    except Exception as exc:  # arquivo corrompido / protegido
        return {"_erro": f"{type(exc).__name__}: {exc}"}

    aba = next((a for a in ABAS_AUTORIZACAO if a in wb.sheetnames), None)
    if aba is None:
        wb.close()
        return {"_erro": f"sem aba de autorização (abas: {wb.sheetnames})"}

    ws = wb[aba]
    dados["_aba"] = aba
    for campo, celula in CELULAS.items():
        valor = ws[celula].value
        if valor is None:
            continue
        if isinstance(valor, datetime):
            valor = valor.date()
        if isinstance(valor, date):
            dados[campo] = valor.isoformat()
        elif isinstance(valor, str):
            limpo = valor.strip()
            if limpo:
                dados[campo] = limpo
        else:
            dados[campo] = valor
    wb.close()
    return dados


def sufixo_data(formato_mes: str) -> str:
    """O trecho de data no mesmo formato da pasta do mês daquele fornecedor."""
    return "{AAAA}-{MM}" if formato_mes == "AAAA-MM" else "{MM}-{AAAA}"


def padrao_nome_arquivo(
    nome_arquivo: str, ano: int, mes: int, formato_mes: str
) -> tuple[str, list[str]]:
    """
    Transforma 'TI - FORNECEDOR - 2026-07' em 'TI - FORNECEDOR - {AAAA}-{MM}'.

    Devolve também os avisos encontrados. A base real tem arquivo copiado do mês
    anterior sem renomear e arquivo sem data nenhuma no nome — os dois casos
    viram aviso, para o humano decidir, em vez de virar padrão errado.
    """
    base = Path(nome_arquivo).stem
    avisos: list[str] = []

    tinha_ano = f"{ano}" in base
    base = base.replace(f"{ano}", "{AAAA}")

    tinha_mes = bool(re.search(rf"(?<!\d){mes:02d}(?!\d)", base))
    base = re.sub(rf"(?<!\d){mes:02d}(?!\d)", "{MM}", base)

    original = Path(nome_arquivo).stem

    if not tinha_ano and not tinha_mes:
        # Nome sem data nenhuma. Mantém o prefixo que o time já reconhece e
        # acrescenta a data no mesmo formato da pasta do mês.
        sugerido = f"{original.strip().rstrip('-').strip()} - {sufixo_data(formato_mes)}"
        avisos.append(
            f"o nome '{original}' não tem data. Sugeri "
            f"'{sugerido}' — se preferir outro padrão, edite 'padrao_nome_arquivo'."
        )
        return sugerido, avisos

    if not tinha_mes:
        # Provável arquivo do mês anterior copiado e não renomeado.
        outro = re.search(r"(?<!\d)(0[1-9]|1[0-2])(?!\d)", base)
        achado = f" (encontrei o mês {outro.group(1)})" if outro else ""
        corrigido = re.sub(r"(?<!\d)(0[1-9]|1[0-2])(?!\d)", "{MM}", base, count=1)
        avisos.append(
            f"o arquivo está na pasta do mês {mes:02d} mas o nome não bate{achado} "
            "— provavelmente foi copiado do mês anterior sem renomear. "
            f"Corrigi para '{corrigido}'; confira antes de usar."
        )
        return corrigido, avisos

    if not tinha_ano:
        avisos.append(
            f"o nome não contém o ano {ano} — confira 'padrao_nome_arquivo'."
        )

    return base, avisos


def linhas_contas_fixas() -> list[dict]:
    """
    Lê a aba CONTAS FIXAS.

    Delega para `automacao.planilha_contas`, que sabe ler mesmo com a planilha
    aberta no Excel (copia para um temporário quando o Windows nega a leitura).
    """
    sys.path.insert(0, str(RAIZ))
    from automacao.entrega import planilha_contas

    return [
        {
            "linha": l.linha,
            "empresa": l.empresa,
            "fornecedor": l.fornecedor,
            "vencimento": l.vencimento,
            "gerar_boleto": l.gerar_boleto,
            "situacao": l.situacao,
            "contato": l.contato,
            "observacao": l.observacao,
        }
        for l in planilha_contas.ler_contas_fixas()
    ]


def coletar_contas() -> list[dict]:
    """Uma 'conta' = uma pasta que recebe autorização todo mês."""
    contas: list[dict] = []

    for pasta_forn in sorted(p for p in RAIZ_PAGAMENTOS.iterdir() if p.is_dir()):
        # meses diretos da pasta do fornecedor
        meses = []
        for item in pasta_forn.iterdir():
            if not item.is_dir():
                continue
            info = classificar_mes(item.name)
            if info:
                meses.append((info[0], info[1], info[2], item))
        if not meses:
            continue

        meses.sort(key=lambda t: (t[1], t[2]))
        formato_mes = meses[-1][0]

        # Cada sub-unidade (cidade / linha / serviço) é uma conta independente.
        # A decisão tem duas partes, e as duas importam:
        #
        #   1. SE existem sub-unidades, quem diz é o mês mais recente que já
        #      tem autorização. Layout antigo não vale — há fornecedor que
        #      usava subpasta até fevereiro e hoje guarda tudo solto no mês.
        #   2. QUAIS são elas, é a união dos últimos meses. O mês corrente
        #      costuma estar pela metade: em agosto uma operadora só tinha 4
        #      das 6 cidades, e olhar só para ele fazia sumir duas contas.
        subunidades: list[str | None] = [None]
        recentes = [pasta_mes for _, _, _, pasta_mes in reversed(meses)]
        decisivos = [
            (p, v) for p in recentes if (v := detectar_subunidades(p)) is not None
        ]

        if decisivos and decisivos[0][1]:
            # A lista sai das SUBPASTAS do mês mais novo, sem exigir que já
            # exista autorização dentro. A sub-unidade sem autorização é
            # justamente o trabalho que falta fazer — exigir o xlsx sumia com
            # duas cidades de uma operadora em agosto, que eram justamente as
            # pendentes. Esconder o pendente é o pior erro possível aqui.
            subunidades = sorted(p.name for p in recentes[0].iterdir() if p.is_dir())

        for subunidade in subunidades:
            # Procura, do mês mais recente para trás, a última pasta que
            # realmente tem uma autorização preenchida. Pastas de mês em
            # andamento (criadas mas ainda vazias) são puladas.
            escolhido = None
            for _, ano, mes, pasta_mes in reversed(meses):
                alvo = (pasta_mes / subunidade) if subunidade else pasta_mes
                if not alvo.is_dir():
                    continue
                xlsx = encontrar_xlsx(alvo)
                if xlsx:
                    escolhido = (ano, mes, pasta_mes, alvo, xlsx)
                    break

            diagnostico: list[str] = []
            if escolhido is None:
                ano_ref, mes_ref, pasta_mes_ref = meses[-1][1], meses[-1][2], meses[-1][3]
                pasta_alvo = (pasta_mes_ref / subunidade) if subunidade else pasta_mes_ref
                xlsx = None
                padrao = None
                extraido = {}
                diagnostico.append(
                    "nenhuma autorização preenchida encontrada nesta pasta — "
                    "a automação não tem de onde copiar. Preencha uma vez na mão."
                )
            else:
                ano_ref, mes_ref, pasta_mes_ref, pasta_alvo, xlsx = escolhido
                padrao, diagnostico = padrao_nome_arquivo(
                    xlsx.name, ano_ref, mes_ref, formato_mes
                )
                extraido = ler_autorizacao(xlsx)
                if extraido.get("_erro"):
                    diagnostico.append(f"leitura da planilha falhou: {extraido['_erro']}")
                if not extraido.get("descricao"):
                    diagnostico.append(
                        "a célula A32 (descrição do evento) está vazia na planilha base."
                    )

            contas.append(
                {
                    "id": gerar_id(pasta_forn.name, subunidade or ""),
                    "pasta": pasta_forn.name,
                    "subunidade": subunidade,
                    "formato_mes": formato_mes,
                    "mes_referencia": f"{ano_ref:04d}-{mes_ref:02d}",
                    "pasta_mes_exemplo": pasta_mes_ref.name,
                    "xlsx_exemplo": xlsx.name if xlsx else None,
                    "padrao_nome": padrao,
                    "extraido": extraido,
                    "diagnostico": diagnostico,
                    "arquivos_exemplo": sorted(
                        p.name for p in pasta_alvo.iterdir() if p.is_file()
                    )
                    if pasta_alvo.is_dir()
                    else [],
                }
            )

    return contas


def montar_yaml(contas: list[dict], contas_fixas: list[dict]) -> dict:
    saida = {
        "_meta": {
            "gerado_em": date.today().isoformat(),
            "origem": "scripts/gerar_registro_fornecedores.py",
            "aviso": (
                "RASCUNHO gerado a partir da última autorização preenchida de cada "
                "pasta. Revise antes de usar: campos podem estar vazios, "
                "desatualizados ou herdados de um mês atípico."
            ),
        },
        "contas": [],
    }

    for c in contas:
        ex = c["extraido"]
        saida["contas"].append(
            {
                "id": c["id"],
                "ativo": True,
                "pasta": c["pasta"],
                "subunidade": c["subunidade"],
                "formato_mes": c["formato_mes"],
                "padrao_nome_arquivo": c["padrao_nome"],
                "aba_autorizacao": ex.get("_aba"),
                "modelo_base": (
                    f"{c['pasta']}/{c['pasta_mes_exemplo']}"
                    + (f"/{c['subunidade']}" if c["subunidade"] else "")
                    + f"/{c['xlsx_exemplo']}"
                    if c["xlsx_exemplo"]
                    else None
                ),
                "autorizacao": {
                    "departamento": ex.get("departamento"),
                    "pagante": ex.get("pagante"),
                    "beneficiario": ex.get("beneficiario"),
                    "forma_pgto": ex.get("forma_pgto"),
                    "meio_pgto": ex.get("meio_pgto"),
                    "setor": ex.get("setor"),
                    "motivo": ex.get("motivo"),
                    "regional": ex.get("regional"),
                    "cooperativa": ex.get("cooperativa"),
                    "rotulo_documento": ex.get("rotulo_documento"),
                    "centro_custo": ex.get("centro_custo"),
                    "natureza": ex.get("natureza"),
                    "pgto_previsto": ex.get("pgto_previsto"),
                    "descricao": ex.get("descricao"),
                },
                "referencia": {
                    "ultimo_mes_processado": c["mes_referencia"],
                    "ultimo_valor": ex.get("valor"),
                    "ultimo_vencimento": ex.get("vencimento"),
                    # C29 — em várias contas guarda o número do contrato/linha,
                    # que é o que amarra a conta à linha da planilha.
                    "ultimo_numero_documento": ex.get("numero_documento"),
                    "arquivos": c["arquivos_exemplo"],
                    "erro_leitura": ex.get("_erro"),
                },
                # Pontos que precisam de decisão sua. Enquanto houver item aqui,
                # trate os defaults desta conta com desconfiança.
                "revisar": c["diagnostico"],
                # preenchido na revisão humana
                "email": {"eh_operadora": None, "cidade": None},
                "coleta": {"remetentes": [], "assunto_contem": []},
                "planilha_contas": {"chaves": []},
            }
        )

    saida["contas_fixas_planilha"] = contas_fixas
    return saida


def main() -> None:
    contas = coletar_contas()
    fixas = linhas_contas_fixas()
    doc = montar_yaml(contas, fixas)

    SAIDA.parent.mkdir(parents=True, exist_ok=True)
    with SAIDA.open("w", encoding="utf-8") as fh:
        yaml.safe_dump(doc, fh, allow_unicode=True, sort_keys=False, width=120)

    precisam_revisao = [c for c in doc["contas"] if c["revisar"]]

    print(f"{len(contas)} contas detectadas / {len(fixas)} linhas em CONTAS FIXAS")
    print(f"{len(precisam_revisao)} conta(s) precisam da sua revisão\n")

    print(f"{'ID':<38} {'PADRÃO DE NOME':<38} DESCRIÇÃO")
    print("-" * 118)
    for c in doc["contas"]:
        desc = c["autorizacao"]["descricao"] or "—"
        marca = "!" if c["revisar"] else " "
        print(
            f"{marca}{c['id']:<37} {str(c['padrao_nome_arquivo'] or '—'):<38} "
            f"{desc[:40]}"
        )

    if precisam_revisao:
        print("\n" + "=" * 118)
        print("PRECISAM DA SUA DECISÃO")
        print("=" * 118)
        for c in precisam_revisao:
            print(f"\n{c['id']}")
            for aviso in c["revisar"]:
                print(f"   • {aviso}")

    print(f"\nRascunho salvo em {SAIDA}")


if __name__ == "__main__":
    main()
