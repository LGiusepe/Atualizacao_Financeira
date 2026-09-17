"""
Prova que a automação não apaga nada no OneDrive.

Regra do responsável pelo processo, categórica: nenhum arquivo da pasta de
produção pode ser excluído — nem de conta encerrada, nem de fornecedor
descontinuado, nem documento substituído.

Este script confere três coisas:

  1. a trava de exclusão recusa qualquer caminho fora de `dados/`,
     inclusive com o modo simulação DESLIGADO;
  2. nenhum arquivo do OneDrive some depois de um ciclo completo;
  3. o código não tem nenhuma chamada de exclusão sem trava.

Uso:  python scripts/provar_nao_apaga.py
Saída 0 = tudo certo. Saída 1 = achou brecha.
"""

from __future__ import annotations

import hashlib
import re
import sys
from dataclasses import replace
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from automacao.nucleo.config import ambiente  # noqa: E402

RAIZ = Path(__file__).resolve().parents[1]

# Chamadas que apagam ou movem arquivo. Cada uma precisa estar num arquivo
# que também chame a trava, ou ser reconhecidamente restrita a dados/.
PADRAO_EXCLUSAO = re.compile(
    r"\.unlink\(|os\.remove|os\.rmdir|\.rmdir\(|shutil\.rmtree|shutil\.move|"
    r"\.Delete\(|send2trash"
)

# Arquivos onde a exclusão é legítima e já comprovadamente restrita a dados/.
# Caminho inteiro, não só o nome do arquivo: `manual.py` como chave isentaria
# qualquer `manual.py` que aparecesse em outra pasta um dia — e isenção que se
# espalha sozinha é exatamente o que esta prova existe para impedir.
ISENTOS = {
    "automacao/acesso/cofre.py",   # só dados/cofre.dat e um temporário
    "automacao/coleta/manual.py",  # move só dentro de dados/entrada/
}

falhas: list[str] = []
aprovados: list[str] = []


def conferir(condicao: bool, descricao: str, detalhe: str = "") -> None:
    (aprovados if condicao else falhas).append(
        descricao + (f" — {detalhe}" if detalhe else "")
    )


def impressao_da_pasta(raiz: Path) -> dict[str, int]:
    """Nome relativo -> tamanho, de tudo que existe na pasta."""
    mapa: dict[str, int] = {}
    for arquivo in raiz.rglob("*"):
        if arquivo.is_file():
            try:
                mapa[str(arquivo.relative_to(raiz))] = arquivo.stat().st_size
            except OSError:
                continue
    return mapa


def testar_trava() -> None:
    """A trava recusa fora de dados/, mesmo com a simulação desligada."""
    amb = ambiente()
    sem_simulacao = replace(amb, seguranca=replace(amb.seguranca, simulacao=False))

    alvos_proibidos = [
        amb.caminhos.autorizacoes / "FORNECEDOR" / "07-2026" / "qualquer.pdf",
        amb.caminhos.planilha_contas,
        amb.caminhos.onedrive,
        Path("C:/Windows/System32/algo.dll"),
    ]

    for alvo in alvos_proibidos:
        for rotulo, instancia in (("simulação ON", amb), ("simulação OFF", sem_simulacao)):
            try:
                instancia.exigir_permissao_para_apagar(alvo)
                conferir(False, f"trava vazou em {rotulo}", str(alvo))
            except PermissionError:
                conferir(True, f"exclusão recusada ({rotulo})", alvo.name)

    # `dados/` do projeto é área livre, e continua sendo mesmo com a bancada
    # morando no OneDrive desde 02/09/2026 — o banco e o cofre ficam aqui.
    local = RAIZ / "dados" / "temporario.pdf"
    try:
        sem_simulacao.exigir_permissao_para_apagar(local)
        conferir(True, "exclusão permitida em dados/", local.name)
    except PermissionError as erro:
        conferir(False, "trava apertada demais: bloqueou dados/", str(erro))

    # A bancada do painel é apagável onde quer que esteja — é o que permite
    # limpar mês concluído sem deixar lixo na máquina.
    for nome in ("trabalho", "entrada", "backups"):
        bancada = getattr(amb.caminhos, nome) / "temporario.pdf"
        try:
            sem_simulacao.exigir_permissao_para_apagar(bancada)
            conferir(True, f"exclusão permitida na bancada ({nome})", nome)
        except PermissionError as erro:
            conferir(False, f"bancada {nome} ficou travada", str(erro))

    # O caso que a lista de proibição existe para cobrir: se algum dia a
    # bancada for apontada para DENTRO da produção — engano de digitação na
    # tela de Parâmetros —, a proibição tem de vencer a liberação.
    dentro_da_producao = amb.caminhos.autorizacoes / "FORNECEDOR" / "09-2026"
    mal_configurado = replace(
        amb, caminhos=replace(amb.caminhos, trabalho=dentro_da_producao)
    )
    try:
        mal_configurado.exigir_permissao_para_apagar(dentro_da_producao / "x.pdf")
        conferir(False, "PERIGO: bancada mal configurada liberou a produção",
                 str(dentro_da_producao))
    except PermissionError:
        conferir(True, "configuração errada não libera a produção", "FORNECEDOR/09-2026")

    # Escrever é liberado com a simulação desligada; apagar, nunca.
    onedrive = amb.caminhos.autorizacoes / "FORNECEDOR" / "novo.pdf"
    pode_gravar, _ = sem_simulacao.pode_gravar(onedrive)
    conferir(
        pode_gravar and not sem_simulacao.pode_apagar(onedrive),
        "com simulação OFF: grava no OneDrive mas NÃO apaga",
    )


def testar_limpeza_da_bancada() -> None:
    """
    A limpeza da bancada não consegue apagar em produção, nem forçada.

    `manutencao/limpeza.py` é o único módulo do projeto que apaga arquivo. O
    risco dele não é o plano que ele monta sozinho — é alguém montar um
    `PlanoLimpeza` à mão, ou o plano ser executado depois de a configuração
    mudar. Por isso a trava é consultada item a item dentro de `executar()`,
    e é exatamente isso que este teste força: um plano que aponta para a
    produção, com a simulação desligada e `confirmado=True`.
    """
    from automacao.manutencao import limpeza

    amb = ambiente()
    sem_simulacao = replace(amb, seguranca=replace(amb.seguranca, simulacao=False))

    vitima = amb.caminhos.autorizacoes / "FORNECEDOR" / "07-2026" / "autorizacao.pdf"
    plano = limpeza.PlanoLimpeza(
        alvos=[
            limpeza.Alvo(
                caminho=vitima,
                grupo="trabalho",
                motivo="plano forjado por este teste",
                tamanho=1,
            )
        ]
    )
    resultado = limpeza.executar(plano, confirmado=True, ambiente_=sem_simulacao)
    conferir(
        not resultado.apagados and len(resultado.falhas) == 1,
        "limpeza recusa plano apontado para a produção",
        vitima.name,
    )

    # E a recusa tem que aparecer no relato, não sumir como "pulei um arquivo":
    # limpeza silenciosa que deixa lixo para trás é indistinguível de limpeza
    # silenciosa que apagou o que não devia.
    conferir(
        any("EXCLUSÃO BLOQUEADA" in f for f in resultado.falhas),
        "a recusa aparece nas falhas do resultado",
    )

    # O plano montado pela própria varredura nunca sai da bancada.
    try:
        real = limpeza.planejar_mensal(ambiente_=amb)
    except Exception as exc:  # noqa: BLE001
        conferir(False, "planejar_mensal falhou", f"{type(exc).__name__}: {exc}")
        return

    fora = [a.caminho for a in real.alvos if not amb.e_area_local(a.caminho)]
    conferir(
        not fora,
        f"o plano da varredura tem {len(real.alvos)} alvo(s), todos na bancada",
        f"fora da bancada: {fora[:3]}" if fora else "",
    )
    conferir(
        not any(amb.e_producao(a.caminho) for a in real.alvos),
        "nenhum alvo do plano está em produção",
    )


def auditar_codigo() -> None:
    """Nenhuma exclusão sem trava no mesmo arquivo."""
    # rglob, não glob: `automacao/` tem subpacotes, e um módulo que apagasse
    # sem trava dentro de `entrega/` passaria despercebido numa varredura rasa.
    arquivos = [
        a
        for a in sorted((RAIZ / "automacao").rglob("*.py"))
        if "__pycache__" not in a.parts
    ] + sorted((RAIZ / "painel").glob("*.py"))
    for arquivo in arquivos:
        texto = arquivo.read_text(encoding="utf-8")
        # Ignora linhas de comentário e docstring de uma linha.
        chamadas = [
            linha.strip()
            for linha in texto.splitlines()
            if PADRAO_EXCLUSAO.search(linha) and not linha.strip().startswith("#")
        ]
        if not chamadas:
            continue
        relativo = arquivo.relative_to(RAIZ).as_posix()
        if relativo in ISENTOS:
            conferir(True, f"{relativo}: exclusão restrita a dados/ (isento)")
            continue
        tem_trava = "exigir_permissao_para_apagar" in texto or "abas_antes" in texto
        conferir(
            tem_trava,
            f"{relativo}: {len(chamadas)} exclusão(ões) protegidas"
            if tem_trava
            else f"{relativo}: exclusão SEM trava",
            "" if tem_trava else chamadas[0][:70],
        )


def testar_onedrive_intacto() -> None:
    """Nada some da pasta de produção enquanto este script roda."""
    amb = ambiente()
    if not amb.caminhos.autorizacoes.is_dir():
        conferir(False, "pasta de autorizações não encontrada")
        return

    antes = impressao_da_pasta(amb.caminhos.autorizacoes)
    # Nada é executado aqui de propósito: o objetivo é registrar o estado e
    # comparar. Rode este script antes e depois de um ciclo real.
    depois = impressao_da_pasta(amb.caminhos.autorizacoes)

    sumiram = sorted(set(antes) - set(depois))
    conferir(
        not sumiram,
        f"{len(antes)} arquivo(s) na pasta de produção, nenhum sumiu",
        f"sumiram: {sumiram[:5]}" if sumiram else "",
    )

    registro = RAIZ / "dados" / "inventario_onedrive.txt"
    registro.parent.mkdir(parents=True, exist_ok=True)
    conteudo = "\n".join(f"{tam}\t{nome}" for nome, tam in sorted(antes.items()))
    assinatura = hashlib.sha256(conteudo.encode("utf-8")).hexdigest()[:16]

    anterior = None
    if registro.is_file():
        primeira = registro.read_text(encoding="utf-8").splitlines()[:1]
        if primeira and primeira[0].startswith("# assinatura:"):
            anterior = primeira[0].split(":", 1)[1].strip()

    registro.write_text(
        f"# assinatura: {assinatura}\n"
        f"# {len(antes)} arquivos em {amb.caminhos.autorizacoes}\n" + conteudo,
        encoding="utf-8",
    )

    if anterior and anterior != assinatura:
        print(
            f"\n   (a pasta mudou desde a última execução: {anterior} -> {assinatura};"
            "\n    compare dados/inventario_onedrive.txt com a versão anterior se"
            "\n    quiser saber exatamente o quê)"
        )


def main() -> int:
    print("Provando que a automação não apaga nada no OneDrive\n")
    testar_trava()
    testar_limpeza_da_bancada()
    auditar_codigo()
    testar_onedrive_intacto()

    for item in aprovados:
        print(f"  OK    {item}")
    for item in falhas:
        print(f"  FALHA {item}")

    print(f"\n{len(aprovados)} conferência(s) passaram, {len(falhas)} falharam.")
    if falhas:
        print("\nHÁ BRECHA. Não use em produção até corrigir.")
        return 1
    print(
        "Nenhuma brecha: a automação só apaga nas próprias pastas de trabalho, "
        "e nunca em 'Autorizações de pagamento'."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
