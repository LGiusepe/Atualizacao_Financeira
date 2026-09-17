"""
Varre a pasta "Autorizações de pagamento" (SOMENTE LEITURA) e produz um raio-x
da estrutura: para cada fornecedor, quais formatos de pasta de mês existem,
quais sub-unidades (cidade/linha) existem e como os arquivos são nomeados.

Saída: dados/raio_x_estrutura.json  (+ resumo no stdout)

Nada é escrito no OneDrive.
"""

from __future__ import annotations

import json
import re
import sys
from collections import Counter
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

from automacao.nucleo.config import ambiente

# Lido do settings.yaml: o caminho real da empresa mora lá, não aqui.
RAIZ_PAGAMENTOS = ambiente().caminhos.autorizacoes
SAIDA = Path(__file__).resolve().parents[1] / "dados" / "raio_x_estrutura.json"

# Padrões de nome de pasta de mês observados na base real.
PADROES_MES = [
    ("MM-AAAA", re.compile(r"^(?P<m>\d{2})\s*-\s*(?P<a>\d{4})$")),
    ("AAAA-MM", re.compile(r"^(?P<a>\d{4})\s*-\s*(?P<m>\d{2})$")),
    ("MM_AAAA", re.compile(r"^(?P<m>\d{2})_(?P<a>\d{4})$")),
    ("AAAA", re.compile(r"^(?P<a>\d{4})$")),
]


def classificar_pasta_mes(nome: str) -> tuple[str, int | None, int | None]:
    """Devolve (padrao, ano, mes). padrao == "" quando não é pasta de mês."""
    limpo = nome.strip()
    for padrao, rx in PADROES_MES:
        m = rx.match(limpo)
        if not m:
            continue
        ano = int(m.group("a"))
        mes = int(m.group("m")) if "m" in m.groupdict() and m.group("m") else None
        if mes is not None and not (1 <= mes <= 12):
            continue
        return padrao, ano, mes
    return "", None, None


def tipo_arquivo(nome: str) -> str:
    """Heurística inicial de classificação — refinada depois pelo classificador."""
    n = nome.lower()
    if n.endswith((".xlsx", ".xls", ".xlsm")):
        return "autorizacao_xlsx"
    if not n.endswith(".pdf"):
        return "outro"
    if "boleto" in n:
        return "boleto"
    if any(t in n for t in ("nfse", "nfs_", "nfs-", "nota", " nf ", "-nf-", "danfe")):
        return "nota_fiscal"
    if any(t in n for t in ("extrato", "demonstrativo", "detalhe", "fatura_extrato")):
        return "demonstrativo"
    if "autoriza" in n:
        return "autorizacao_pdf"
    if "fatura" in n or "conta" in n:
        return "fatura"
    return "pdf_indefinido"


def varrer() -> dict:
    if not RAIZ_PAGAMENTOS.is_dir():
        raise SystemExit(f"Pasta não encontrada: {RAIZ_PAGAMENTOS}")

    fornecedores: dict[str, dict] = {}

    for pasta_forn in sorted(p for p in RAIZ_PAGAMENTOS.iterdir() if p.is_dir()):
        info: dict = {
            "pasta": pasta_forn.name,
            "padroes_mes": Counter(),
            "meses": [],
            "tem_subunidades": False,
            "subunidades": set(),
            "exemplos_arquivo": [],
            "tipos_encontrados": Counter(),
            "arquivos_na_raiz": [],
        }

        for item in sorted(pasta_forn.iterdir()):
            if item.is_file():
                info["arquivos_na_raiz"].append(item.name)
                continue

            padrao, ano, mes = classificar_pasta_mes(item.name)
            if not padrao:
                # subpasta que não é mês (ex.: "2025" já tratado, ou avulsa)
                info["subunidades"].add(item.name)
                continue

            info["padroes_mes"][f"{padrao}|{item.name}"] = 0
            info["padroes_mes"][padrao] += 1
            info["meses"].append(
                {"pasta": item.name, "padrao": padrao, "ano": ano, "mes": mes}
            )

            filhos = sorted(item.iterdir())
            subdirs = [c for c in filhos if c.is_dir()]
            arquivos = [c for c in filhos if c.is_file()]

            if subdirs:
                info["tem_subunidades"] = True
                for sd in subdirs:
                    info["subunidades"].add(sd.name)
                    for arq in sorted(sd.iterdir()):
                        if arq.is_file():
                            info["tipos_encontrados"][tipo_arquivo(arq.name)] += 1
                            if len(info["exemplos_arquivo"]) < 14:
                                info["exemplos_arquivo"].append(
                                    f"{item.name}/{sd.name}/{arq.name}"
                                )
            for arq in arquivos:
                info["tipos_encontrados"][tipo_arquivo(arq.name)] += 1
                if len(info["exemplos_arquivo"]) < 14:
                    info["exemplos_arquivo"].append(f"{item.name}/{arq.name}")

        # limpa chaves auxiliares com contagem 0
        info["padroes_mes"] = {
            k: v for k, v in info["padroes_mes"].items() if v and "|" not in k
        }
        info["subunidades"] = sorted(info["subunidades"])
        info["tipos_encontrados"] = dict(info["tipos_encontrados"])
        info["meses"].sort(key=lambda m: (m["ano"] or 0, m["mes"] or 0))
        fornecedores[pasta_forn.name] = info

    return fornecedores


def main() -> None:
    fornecedores = varrer()
    SAIDA.parent.mkdir(parents=True, exist_ok=True)
    SAIDA.write_text(
        json.dumps(fornecedores, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    print(f"{len(fornecedores)} pastas de fornecedor\n")
    print(f"{'PASTA':<34} {'PADRÃO':<12} {'MESES':>5}  {'SUB':>3}  SUBUNIDADES / TIPOS")
    print("-" * 118)
    for nome, info in fornecedores.items():
        padroes = ",".join(info["padroes_mes"]) or "-"
        subs = ", ".join(info["subunidades"][:4])
        if len(info["subunidades"]) > 4:
            subs += f" (+{len(info['subunidades']) - 4})"
        print(
            f"{nome:<34} {padroes:<12} {len(info['meses']):>5}  "
            f"{'sim' if info['tem_subunidades'] else '-':>3}  {subs}"
        )

    print(f"\nRaio-x salvo em {SAIDA}")


if __name__ == "__main__":
    main()
