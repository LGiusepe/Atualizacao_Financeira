"""
Aplica ao registro a lista oficial de contas ativas do mês.

A lista veio do responsável e vale mais que qualquer heurística: ela decide
quem está ativo e qual é o texto exato do FORNECEDOR na planilha (o que o
checklist usa para achar a linha e pintar de verde).

Conta que não está na lista é desativada — sem apagar nada. Basta voltar
`ativo: true` no YAML para reativar.

Uso:  python scripts/aplicar_lista_ativas.py
"""

from __future__ import annotations

import sys
from pathlib import Path

import yaml

sys.stdout.reconfigure(encoding="utf-8")

RAIZ = Path(__file__).resolve().parents[1]
REGISTRO = RAIZ / "config" / "fornecedores.yaml"
DADOS_LOCAIS = RAIZ / "config" / "dados-locais.yaml"


def _carregar_lista() -> tuple[str, list[tuple[str | None, str, str, str]]]:
    """
    A lista oficial, lida de `config/dados-locais.yaml`.

    Ela já foi escrita aqui dentro. Saiu porque nomeia a razão social de doze
    fornecedores, treze números de contrato e, em dois casos, a pessoa dona da
    linha telefônica — e o repositório é público. O modelo do arquivo está em
    `config/dados-locais.exemplo.yaml`.

    Devolve (competência, lista), com a lista na ordem do checklist e no
    formato (id da conta, texto do FORNECEDOR na planilha, contrato/conta,
    observação). Id vazio vira `None`: é a conta que existe na planilha mas
    ainda não tem pasta de autorização no OneDrive, registrada aqui para não
    sumir do radar.
    """
    if not DADOS_LOCAIS.is_file():
        raise SystemExit(
            f"lista oficial não encontrada: {DADOS_LOCAIS}\n"
            "Copie config/dados-locais.exemplo.yaml para config/dados-locais.yaml "
            "e preencha o bloco 'aplicar_lista_ativas'."
        )
    bruto = yaml.safe_load(DADOS_LOCAIS.read_text(encoding="utf-8")) or {}
    bloco = bruto.get("aplicar_lista_ativas") or {}
    itens = bloco.get("ativas") or []
    return str(bloco.get("competencia") or "").strip(), [
        (
            item.get("id") or None,
            str(item.get("fornecedor") or ""),
            str(item.get("contrato") or ""),
            str(item.get("observacao") or ""),
        )
        for item in itens
    ]


COMPETENCIA, ATIVAS = _carregar_lista()


def main() -> None:
    if not REGISTRO.is_file():
        raise SystemExit(f"registro não encontrado: {REGISTRO}")

    doc = yaml.safe_load(REGISTRO.read_text(encoding="utf-8"))
    contas = {c["id"]: c for c in doc["contas"]}

    da_lista = {id_ for id_, *_ in ATIVAS if id_}
    desconhecidos = sorted(da_lista - set(contas))
    sem_pasta = [(texto, obs) for id_, texto, _, obs in ATIVAS if id_ is None]

    ativadas: list[str] = []
    for id_conta, fornecedor, contrato, observacao in ATIVAS:
        if not id_conta or id_conta not in contas:
            continue
        conta = contas[id_conta]
        conta["ativo"] = True
        # O texto do FORNECEDOR com o contrato é como o leitor da planilha
        # monta a chave (fornecedor + " - " + contrato quando há contrato).
        # Há fornecedor cujo texto do FORNECEDOR já traz o número do contrato
        # dentro; concatenar de novo repetiria o número na chave.
        chave = (
            f"{fornecedor} - {contrato}"
            if contrato and contrato not in fornecedor
            else fornecedor
        )
        conta["planilha_contas"] = {"chaves": [chave]}
        if observacao:
            conta["observacao"] = observacao
        conta["revisar"] = [
            aviso
            for aviso in (conta.get("revisar") or [])
            # o vínculo com a planilha deixou de ser dúvida
            if "CONTAS FIXAS" not in aviso and "planilha_contas" not in aviso
        ]
        if observacao.startswith("CONFIRMAR"):
            conta["revisar"].append(observacao)
        ativadas.append(id_conta)

    desativadas = []
    for id_conta, conta in contas.items():
        if id_conta in da_lista:
            continue
        if conta.get("ativo"):
            desativadas.append(id_conta)
        conta["ativo"] = False
        conta["observacao"] = (
            f"fora da lista oficial de contas ativas de {COMPETENCIA} — "
            "reative mudando 'ativo' para true"
        )

    doc["_meta"]["lista_ativas_aplicada"] = f"{COMPETENCIA} (informada pelo responsável)"
    with REGISTRO.open("w", encoding="utf-8") as fh:
        yaml.safe_dump(doc, fh, allow_unicode=True, sort_keys=False, width=120)

    print(f"{len(ativadas)} conta(s) ativadas · {len(desativadas)} desativadas\n")

    print("ATIVAS")
    print("-" * 96)
    for id_conta, fornecedor, contrato, obs in ATIVAS:
        if not id_conta:
            continue
        marca = "!" if obs.startswith("CONFIRMAR") else " "
        alvo = f"{fornecedor} - {contrato}" if contrato else fornecedor
        print(f"{marca}{id_conta:<32} {alvo[:56]}")

    if sem_pasta:
        print(f"\nNA LISTA MAS SEM PASTA NO ONEDRIVE ({len(sem_pasta)})")
        print("-" * 96)
        for texto, obs in sem_pasta:
            print(f"   {texto:<52} {obs}")
        print("   → crie a pasta em 'Autorizações de pagamento' e rode")
        print("     gerar_registro_fornecedores.py + enriquecer_registro.py de novo.")

    if desativadas:
        print(f"\nDESATIVADAS ({len(desativadas)})")
        print("-" * 96)
        print("   " + ", ".join(sorted(desativadas)))

    if desconhecidos:
        print(f"\nIDS DA LISTA QUE NÃO EXISTEM NO REGISTRO: {desconhecidos}")

    print(f"\nRegistro atualizado: {REGISTRO}")


if __name__ == "__main__":
    main()
