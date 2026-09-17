"""
Confere o que o git levaria para o GitHub antes de você publicar.

Não altera nada: só lê os arquivos que o git considera versionados e procura
identificadores da empresa — CNPJ, e-mail, telefone, conta bancária, número de
contrato, caminho da máquina, nome da empresa, nomes de pessoas e nomes de
fornecedor.

O que procurar vem de `config/dados-locais.yaml`, que não é versionado. Antes
as listas ficavam escritas aqui dentro, e este arquivo — versionado — acabava
publicando exatamente os nomes que existe para proteger. Sem o arquivo local,
a conferência ainda roda: os padrões genéricos (CNPJ, e-mail, telefone,
contrato) não dependem dele, e o aviso diz o que ficou sem cobertura.

    python scripts/conferir_publicacao.py

Sai com código 1 se achar algo, para poder ser usado em hook de pre-commit.
"""

from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

import yaml

RAIZ = Path(__file__).resolve().parents[1]

# `python scripts/conferir_publicacao.py` põe só `scripts/` no path. O
# `import automacao.modelos` de `e_marcador` (dígito verificador de CNPJ)
# só acontece quando algo é encontrado — então, sem esta linha, o
# conferidor quebrava exatamente no caso em que ele tem algo a dizer.
if str(RAIZ) not in sys.path:
    sys.path.insert(0, str(RAIZ))

DADOS_LOCAIS = RAIZ / "config" / "dados-locais.yaml"


def _listas_locais() -> dict[str, list[str]]:
    """
    Os nomes a procurar, lidos de `config/dados-locais.yaml`.

    Devolve dicionário vazio quando o arquivo não existe: um clone novo não
    tem nomes da empresa para procurar, e quebrar aqui faria a pessoa pular a
    conferência inteira — que é justamente o oposto do que se quer.
    """
    if not DADOS_LOCAIS.is_file():
        return {}
    try:
        bruto = yaml.safe_load(DADOS_LOCAIS.read_text(encoding="utf-8")) or {}
    except (OSError, yaml.YAMLError):
        return {}
    bloco = bruto.get("conferir_publicacao") or {}
    return {
        chave: [str(v) for v in (bloco.get(chave) or []) if str(v).strip()]
        for chave in ("nome_da_empresa", "nomes_de_pessoa", "fornecedores")
    }


def _padrao_de_nomes(nomes: list[str]) -> str:
    """
    Um padrão que acha cada nome inteiro, com espaço livre entre as palavras.

    O espaço livre é o conserto de um furo real. O padrão do nome da empresa
    era escrito à mão e exigia um espaço entre as duas palavras; a grafia
    usada no código juntava as duas, com a segunda em minúsculas. Não batia, e
    a senha de primeiro acesso viajou para o repositório público sem ninguém
    ser avisado. Agora o nome é declarado uma vez, com espaço, e o padrão
    aceita qualquer junção — combinado com `re.IGNORECASE`, qualquer caixa.
    """
    partes = [
        r"\b" + r"\s*".join(re.escape(p) for p in nome.split()) + r"\b"
        for nome in nomes
        if nome.strip()
    ]
    return "|".join(partes)


def _padrao_de_nomes_proprios(nomes: list[str]) -> str:
    """
    O mesmo, mas só nas grafias em que um nome próprio aparece de verdade:
    como foi escrito, tudo em caixa alta e capitalizado.

    A forma toda minúscula fica de fora de propósito. `claro` e `vivo` são
    palavras comuns do português — "está claro", "a conta continua viva" — e
    acusá-las enche o relatório de ruído, que é exatamente como um vazamento
    de verdade passa despercebido: a pessoa aprende a ignorar a saída.

    O nome da empresa não entra nesta regra. Esse é procurado sem olhar caixa,
    porque é o que não pode escapar de jeito nenhum, e não colide com nenhuma
    palavra do idioma.
    """
    variantes = sorted({v for nome in nomes if nome.strip()
                        for v in (nome, nome.upper(), nome.title())})
    return _padrao_de_nomes(variantes)


_LISTAS = _listas_locais()

#: O que procurar: rótulo -> (padrão, flags). Tudo aqui é grave. Nome de
#: fornecedor já foi "decisão sua", porque o repositório era privado; num
#: repositório público ele revela com quem a empresa contrata, e junto do
#: número do contrato é material de golpe de fatura falsa.
PADROES: dict[str, tuple[str, int]] = {
    "CNPJ": (r"\d{2}\.\d{3}\.\d{3}[/.]\d{4}-\d{2}", 0),
    "e-mail": (r"[\w.+-]+@[\w-]+\.[\w.]{2,}", 0),
    "telefone": (r"\(\d{2}\)\s?\d{4,5}-\d{4}", 0),
    "conta bancária": (r"\b\d{4,6}-\d\b", 0),
    # Contrato de operadora: três dígitos, barra, mais nove; ou o número solto
    # de 9 a 14 dígitos. Não havia padrão nenhum para isto, e treze contratos
    # reais ficaram públicos porque nenhuma das outras regras os alcançava.
    "número de contrato": (r"\b\d{3}/\d{6,12}\b|\b\d{9,14}\b", 0),
    "caminho da máquina": (r"C:\+Users\+[A-Za-z0-9._-]+", 0),
}

if _LISTAS.get("nome_da_empresa"):
    PADROES["nome da empresa"] = (
        _padrao_de_nomes(_LISTAS["nome_da_empresa"]), re.IGNORECASE
    )
if _LISTAS.get("nomes_de_pessoa"):
    PADROES["nome de pessoa"] = (
        _padrao_de_nomes_proprios(_LISTAS["nomes_de_pessoa"]), 0
    )
if _LISTAS.get("fornecedores"):
    PADROES["fornecedor real"] = (
        _padrao_de_nomes_proprios(_LISTAS["fornecedores"]), 0
    )

#: Domínios que só existem em exemplo. Endereço com um destes não é vazamento.
DOMINIOS_DE_EXEMPLO = (
    "exemplo.com", "exemplo.com.br", "exemplo.invalido", "example.com",
    "example.org", "suaempresa.com.br", "empresa.com.br", "x.br",
    "seudominio.com.br", "invalid", "invalido", "teste.com",
)

#: Valores fictícios usados nos exemplos e nas docstrings do projeto.
#: `11.222.333/0001-81` é o CNPJ de teste clássico: precisa ter dígito
#: verificador válido para a docstring de `formatar_cnpj` dizer a verdade.
#: Os quatro números longos são códigos de erro do COM do Excel
#: (`-2147221005` e parentes), não contrato de ninguém.
MARCADORES = {
    "11.222.333/0001-81", "11222333000181",
    "00.000.000/0001-00", "1.234,56", "1.234,50",
    "2147221005", "2147221164", "2146959355", "2147417846",
}


def e_marcador(rotulo: str, termo: str) -> bool:
    """
    Distingue exemplo de dado real.

    Um conferidor que acusa os próprios marcadores acostuma a pessoa a ignorar
    o resultado — e aí, no dia em que algo real escapar, ninguém olha.
    """
    if termo in MARCADORES:
        return True
    if rotulo == "e-mail":
        return any(termo.lower().endswith(d) for d in DOMINIOS_DE_EXEMPLO)
    if rotulo == "CNPJ":
        # Dígito verificador que não fecha é número inventado.
        from automacao.nucleo.modelos import cnpj_confere

        return not cnpj_confere(termo)
    if rotulo in ("telefone", "conta bancária", "número de contrato"):
        # Só zeros (e a pontuação) é preenchimento de formulário.
        return set(c for c in termo if c.isdigit()) <= {"0"}
    return False


def arquivos_versionados() -> list[str]:
    """
    O que o git levaria, respeitando o .gitignore.

    Exige repositório: `git check-ignore` só funciona dentro de uma árvore de
    trabalho. Sem isso a conferência devolveria a pasta inteira — inclusive o
    que o .gitignore barra — e daria um susto falso.
    """
    listados = subprocess.run(
        ["git", "ls-files", "--cached", "--others", "--exclude-standard"],
        cwd=RAIZ, capture_output=True, text=True, encoding="utf-8",
    )
    if listados.returncode != 0:
        print("Este diretório ainda não é um repositório git.")
        print("Rode `git init` primeiro — só assim dá para saber o que o")
        print(".gitignore está barrando de verdade.")
        raise SystemExit(2)
    return [l for l in listados.stdout.splitlines() if l.strip()]


def main() -> int:
    binarios = {".pdf", ".xlsx", ".xls", ".db", ".png", ".jpg", ".jpeg",
                ".dat", ".bin", ".pyc"}
    vazamentos: dict[str, list[tuple[str, str]]] = {}
    marcadores = 0

    versionados = arquivos_versionados()
    for relativo in versionados:
        caminho = RAIZ / relativo
        if caminho.suffix.lower() in binarios:
            continue
        try:
            texto = caminho.read_text(encoding="utf-8", errors="replace")
        except OSError:
            continue
        for rotulo, (padrao, flags) in PADROES.items():
            for termo in sorted(set(re.findall(padrao, texto, flags))):
                if e_marcador(rotulo, termo):
                    marcadores += 1
                else:
                    vazamentos.setdefault(relativo, []).append((rotulo, termo))

    print(f"Conferindo {len(versionados)} arquivo(s) que o git levaria." + "\n")

    faltando = [c for c in ("nome_da_empresa", "nomes_de_pessoa", "fornecedores")
                if not _LISTAS.get(c)]
    if faltando:
        print(f"AVISO: sem {DADOS_LOCAIS.name}, não procurei por "
              f"{', '.join(faltando)}." + "\n")

    if vazamentos:
        print("DADO REAL — não deveria sair:" + "\n")
        for relativo, itens in sorted(vazamentos.items()):
            print(f"  {relativo}")
            for rotulo, termo in itens:
                print(f"      {rotulo:18} {termo[:60]}")
            print()
    else:
        print("Nenhum dado real encontrado.")

    print(f"{marcadores} valor(es) de exemplo ignorados "
          f"(CNPJ zerado, @exemplo.com.br e afins).")

    print()
    if vazamentos:
        print("Corrija antes de publicar.")
        return 1
    print("Pode publicar.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
