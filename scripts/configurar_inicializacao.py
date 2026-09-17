"""
Faz o painel subir sozinho quando você entra no Windows.

O atalho vai para a pasta Inicializar do SEU usuário — nada de tarefa
agendada e nada de pedir administrador. Isso importa por três motivos:

* o painel precisa do perfil do usuário montado para enxergar o OneDrive
  sincronizado; subir antes do logon não adiantaria nada;
* tarefa no Agendador exige elevação e some do lugar onde alguém procuraria;
  o atalho está em `shell:startup`, visível no Explorador, e apagar o arquivo
  desfaz tudo;
* o alvo é o `pythonw.exe`, que **não abre janela**. Sem console preto a cada
  boot, e sem navegador abrindo sozinho — você entra no painel pelo endereço
  quando precisar.

Uso:

    python scripts/configurar_inicializacao.py             # como está hoje
    python scripts/configurar_inicializacao.py --aplicar   # passa a subir sozinho
    python scripts/configurar_inicializacao.py --remover   # volta ao manual

Para encerrar um painel que subiu assim, o `parar-painel.bat` continua
servindo: ele procura pela porta, não pela janela.
"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

RAIZ = Path(__file__).resolve().parents[1]

#: Nome do atalho. Sem acento de propósito: é o que aparece no Explorador e
#: em janelas antigas do Windows, que ainda tropeçam em codificação.
NOME_DO_ATALHO = "Automacao Financeira - Painel.lnk"


def pasta_de_inicializacao() -> Path:
    """
    `shell:startup` do usuário atual.

    Vem de `%APPDATA%`, e não de um caminho fixo: perfil em outra letra de
    disco ou pasta redirecionada pela empresa continuam funcionando.
    """
    appdata = os.environ.get("APPDATA")
    if not appdata:
        raise RuntimeError("variável APPDATA não definida — não sei onde é o perfil")
    return Path(appdata) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"


def caminho_do_atalho() -> Path:
    return pasta_de_inicializacao() / NOME_DO_ATALHO


def interpretador_sem_janela() -> Path:
    """
    O `pythonw.exe` que acompanha este Python.

    Se ele não existir (instalação mínima, alguns ambientes virtuais), caímos
    no `python.exe` — o painel sobe igual, só que com a janela do console
    aparecendo. É degradação visível, e é melhor que recusar a configuração.
    """
    sem_janela = Path(sys.executable).with_name("pythonw.exe")
    return sem_janela if sem_janela.is_file() else Path(sys.executable)


def situacao() -> str:
    atalho = caminho_do_atalho()
    if not atalho.is_file():
        return f"NÃO configurado — o painel só sobe quando você abre o painel.bat.\n  (o atalho ficaria em {atalho})"
    return (
        f"Configurado: {atalho}\n"
        f"  alvo   : {interpretador_sem_janela()} -m painel\n"
        f"  pasta  : {RAIZ}\n"
        "  O painel sobe sozinho no seu próximo logon, sem janela."
    )


def aplicar() -> Path:
    """Cria (ou refaz) o atalho. Refazer é de propósito: caminho muda."""
    import win32com.client

    destino = pasta_de_inicializacao()
    destino.mkdir(parents=True, exist_ok=True)
    atalho = destino / NOME_DO_ATALHO

    shell = win32com.client.Dispatch("WScript.Shell")
    # Toda chamada COM deste projeto é posicional: em late binding o pywin32
    # descarta argumento nomeado em silêncio. Aqui as propriedades são
    # atribuídas uma a uma, que é a forma que não depende disso.
    link = shell.CreateShortCut(str(atalho))
    link.TargetPath = str(interpretador_sem_janela())
    link.Arguments = "-m painel"
    link.WorkingDirectory = str(RAIZ)
    link.Description = "Painel de autorizações de pagamento (sobe sem janela)"
    link.WindowStyle = 7  # minimizado; com pythonw não há janela nenhuma mesmo
    link.save()
    return atalho


def remover() -> bool:
    atalho = caminho_do_atalho()
    if not atalho.is_file():
        return False
    # Este arquivo é do perfil do usuário, não da bancada nem da produção — a
    # trava de exclusão do `Ambiente` não se aplica e nem deveria: ela existe
    # para arquivo do processo, e um atalho criado por este script é do
    # script. Ainda assim, só este nome exato é apagado.
    atalho.unlink()
    return True


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Painel subindo sozinho com o Windows (sem janela, sem admin)."
    )
    parser.add_argument("--aplicar", action="store_true", help="passa a subir sozinho")
    parser.add_argument("--remover", action="store_true", help="volta a subir só na mão")
    args = parser.parse_args()

    if args.aplicar and args.remover:
        print("Escolha um: --aplicar ou --remover.")
        return 2

    if args.aplicar:
        try:
            atalho = aplicar()
        except Exception as exc:  # noqa: BLE001 — a mensagem é o produto aqui
            print(f"Não consegui criar o atalho: {type(exc).__name__}: {exc}")
            return 1
        print(f"Pronto. Atalho criado em:\n  {atalho}\n")
        print("No próximo logon o painel sobe sozinho, sem janela nenhuma.")
        print("Para conferir se está no ar: http://127.0.0.1 (ou a porta em logs/porta.txt).")
        print("Para encerrar: parar-painel.bat.")
        return 0

    if args.remover:
        if remover():
            print("Atalho removido. O painel volta a subir só pelo painel.bat.")
        else:
            print("Não havia atalho para remover.")
        return 0

    print(situacao())
    print("\nPara mudar:  --aplicar  ou  --remover")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
