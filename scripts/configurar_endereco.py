"""
Faz `organizacao.financeira.local` apontar para esta máquina.

O nome não existe em DNS nenhum — e nem deve. Ele é resolvido pelo arquivo
`hosts` do Windows, que só vale aqui: o painel continua inalcançável de
qualquer outra máquina, como sempre foi.

    python scripts/configurar_endereco.py            # mostra o que faria
    python scripts/configurar_endereco.py --aplicar  # grava (pede admin)

Editar o `hosts` exige privilégio de administrador. Sem ele o script explica
o que fazer e não tenta forçar — arquivo de sistema escrito pela metade é
problema pior que um endereço feio.

Desfazer é simples: rode com `--remover`, ou apague a linha à mão.
"""

from __future__ import annotations

import ctypes
import shutil
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

HOSTS = Path(r"C:\Windows\System32\drivers\etc\hosts")
NOME = "organizacao.financeira.local"
ENDERECO = "127.0.0.1"
MARCA = "# Automação Financeira — painel de autorizações de pagamento"
LINHA = f"{ENDERECO}\t{NOME}\t{MARCA}"


def e_administrador() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:  # noqa: BLE001 — fora do Windows, ou API indisponível
        return False


def ja_configurado(texto: str) -> bool:
    for linha in texto.splitlines():
        limpa = linha.strip()
        if limpa.startswith("#"):
            continue
        partes = limpa.split()
        if len(partes) >= 2 and NOME in partes[1:]:
            return True
    return False


def sem_a_linha(texto: str) -> str:
    """O arquivo sem nenhuma entrada para o nosso nome."""
    mantidas = []
    for linha in texto.splitlines():
        partes = linha.strip().split()
        if len(partes) >= 2 and not linha.strip().startswith("#") and NOME in partes[1:]:
            continue
        mantidas.append(linha)
    return "\n".join(mantidas).rstrip() + "\n"


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    aplicar = "--aplicar" in sys.argv
    remover = "--remover" in sys.argv

    if not HOSTS.is_file():
        print(f"Não achei o arquivo hosts em {HOSTS}.")
        return 1

    texto = HOSTS.read_text(encoding="utf-8", errors="replace")
    configurado = ja_configurado(texto)

    print(f"arquivo   : {HOSTS}")
    print(f"nome      : {NOME} -> {ENDERECO}")
    print(f"situação  : {'já configurado' if configurado else 'ainda não configurado'}")
    print(f"admin     : {'sim' if e_administrador() else 'NÃO'}")
    print()

    if not (aplicar or remover):
        print("Nada foi alterado. Para gravar:")
        print("   1. abra o Terminal/PowerShell COMO ADMINISTRADOR")
        print("   2. cd para a pasta do projeto")
        print("   3. python scripts/configurar_endereco.py --aplicar")
        print()
        print("A linha que será acrescentada ao final do arquivo:")
        print(f"   {LINHA}")
        return 0

    if not e_administrador():
        print("Precisa de administrador para escrever no arquivo hosts.")
        print("Feche, abra o Terminal como administrador e rode de novo.")
        return 1

    # Cópia datada antes de qualquer escrita: o hosts é arquivo de sistema.
    copia = HOSTS.with_name(f"hosts.antes-da-automacao-{datetime.now():%Y%m%d-%H%M%S}")
    shutil.copy2(HOSTS, copia)
    print(f"cópia de segurança: {copia}")

    if remover:
        HOSTS.write_text(sem_a_linha(texto), encoding="utf-8")
        print(f"linha de {NOME} removida. O painel segue em http://127.0.0.1.")
        return 0

    if configurado:
        print("Já estava configurado — nada a fazer.")
        return 0

    novo = texto if texto.endswith("\n") else texto + "\n"
    HOSTS.write_text(novo + LINHA + "\n", encoding="utf-8")
    print(f"pronto: http://{NOME} passa a abrir o painel desta máquina.")
    print("Se o navegador insistir em buscar na internet, feche e abra de novo.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
