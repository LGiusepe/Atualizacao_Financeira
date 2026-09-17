"""
Publicação no OneDrive — o ÚNICO módulo autorizado a escrever fora de `dados/`.

Regra combinada com o usuário: nada é alterado na pasta de trabalho sem
explicação prévia. Por isso toda operação tem duas fases:

    1. `planejar()`  -> devolve exatamente o que SERIA feito. Não toca em nada.
    2. `publicar(confirmado=True)` -> executa, com backup e trilha de auditoria.

O painel sempre mostra a fase 1 antes de oferecer a fase 2.
"""

from __future__ import annotations

import hashlib
import logging
import shutil
from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from pathlib import Path

from automacao.nucleo import estado
from automacao.nucleo.config import Ambiente, ambiente
from automacao.nucleo.modelos import Competencia, Conta, Etapa, ResultadoEtapa, Situacao

log = logging.getLogger("automacao.entrega.publicador")


class Acao(str, Enum):
    CRIAR_PASTA = "criar_pasta"
    COPIAR = "copiar"
    SOBRESCREVER = "sobrescrever"
    RENOMEAR = "renomear"

    @property
    def rotulo(self) -> str:
        return {
            "criar_pasta": "Criar pasta",
            "copiar": "Copiar arquivo novo",
            "sobrescrever": "SOBRESCREVER arquivo existente",
            "renomear": "Renomear arquivo",
        }[self.value]


@dataclass
class Operacao:
    """Uma alteração pretendida no OneDrive."""

    acao: Acao
    destino: Path
    origem: Path | None = None
    tamanho: int | None = None
    # Só preenchido em SOBRESCREVER: como está o arquivo hoje.
    tamanho_atual: int | None = None
    observacao: str = ""

    @property
    def arriscada(self) -> bool:
        return self.acao is Acao.SOBRESCREVER

    def descrever(self) -> str:
        if self.acao is Acao.CRIAR_PASTA:
            return f"{self.acao.rotulo}: {self.destino}"
        tam = f" ({self.tamanho / 1024:.0f} KB)" if self.tamanho else ""
        if self.acao is Acao.SOBRESCREVER:
            atual = (
                f" — substitui arquivo de {self.tamanho_atual / 1024:.0f} KB"
                if self.tamanho_atual
                else ""
            )
            return f"{self.acao.rotulo}: {self.destino.name}{tam}{atual}"
        return f"{self.acao.rotulo}: {self.destino.name}{tam}"


@dataclass
class Plano:
    """O conjunto de alterações pretendidas, pronto para revisão humana."""

    conta_id: str
    competencia: Competencia
    destino: Path
    operacoes: list[Operacao] = field(default_factory=list)
    impedimentos: list[str] = field(default_factory=list)

    @property
    def viavel(self) -> bool:
        return not self.impedimentos and bool(self.operacoes)

    @property
    def tem_sobrescrita(self) -> bool:
        return any(o.arriscada for o in self.operacoes)

    def resumo(self) -> list[str]:
        return [o.descrever() for o in self.operacoes]


def _tamanho(caminho: Path) -> int | None:
    try:
        return caminho.stat().st_size
    except OSError:
        return None


def _hash(caminho: Path) -> str:
    h = hashlib.sha256()
    with caminho.open("rb") as fh:
        for bloco in iter(lambda: fh.read(1 << 20), b""):
            h.update(bloco)
    return h.hexdigest()


# --------------------------------------------------------------------------- #
# Fase 1 — planejar (não escreve nada)
# --------------------------------------------------------------------------- #


def planejar(
    conta: Conta,
    competencia: Competencia,
    arquivos: list[Path],
    *,
    ambiente_: Ambiente | None = None,
) -> Plano:
    """
    Monta o plano de publicação: quais pastas criar e quais arquivos copiar.

    Não toca em disco. Serve para o painel mostrar ao usuário, palavra por
    palavra, o que vai acontecer.
    """
    amb = ambiente_ or ambiente()
    destino = amb.destino_onedrive(conta, competencia)
    plano = Plano(conta_id=conta.id, competencia=competencia, destino=destino)

    raiz_forn = amb.pasta_do_fornecedor(conta)
    if not raiz_forn.is_dir():
        plano.impedimentos.append(
            f"a pasta do fornecedor não existe: {raiz_forn}. "
            "Crie manualmente ou corrija 'pasta' no registro da conta."
        )
        return plano

    if not arquivos:
        plano.impedimentos.append("nenhum arquivo para publicar.")
        return plano

    # Pastas que precisam nascer (da mais rasa para a mais profunda).
    faltantes = []
    cursor = destino
    while not cursor.exists() and cursor != raiz_forn.parent:
        faltantes.append(cursor)
        cursor = cursor.parent
    for pasta in reversed(faltantes):
        plano.operacoes.append(Operacao(Acao.CRIAR_PASTA, pasta))

    for arq in arquivos:
        if not arq.is_file():
            plano.impedimentos.append(f"arquivo sumiu da área de trabalho: {arq}")
            continue
        alvo = destino / arq.name
        existe = alvo.exists()
        plano.operacoes.append(
            Operacao(
                acao=Acao.SOBRESCREVER if existe else Acao.COPIAR,
                destino=alvo,
                origem=arq,
                tamanho=_tamanho(arq),
                tamanho_atual=_tamanho(alvo) if existe else None,
                observacao=(
                    "já existe um arquivo com este nome no destino"
                    if existe
                    else ""
                ),
            )
        )

    return plano


# --------------------------------------------------------------------------- #
# Fase 2 — publicar (escreve, com backup e auditoria)
# --------------------------------------------------------------------------- #


def publicar(
    plano: Plano,
    *,
    confirmado: bool = False,
    ambiente_: Ambiente | None = None,
) -> ResultadoEtapa:
    """
    Executa o plano.

    Sem `confirmado=True` devolve `Situacao.PENDENTE` com o plano detalhado —
    e não escreve absolutamente nada.
    """
    amb = ambiente_ or ambiente()

    if plano.impedimentos:
        return ResultadoEtapa.erro(
            Etapa.PUBLICACAO,
            "publicação bloqueada: " + "; ".join(plano.impedimentos),
            detalhes={"impedimentos": plano.impedimentos},
        )

    if not confirmado:
        return ResultadoEtapa(
            etapa=Etapa.PUBLICACAO,
            situacao=Situacao.PENDENTE,
            mensagem=(
                f"{len(plano.operacoes)} alteração(ões) aguardando sua confirmação"
                + (" — ATENÇÃO: há sobrescrita" if plano.tem_sobrescrita else "")
            ),
            detalhes={
                "destino": str(plano.destino),
                "operacoes": plano.resumo(),
                "tem_sobrescrita": plano.tem_sobrescrita,
            },
        )

    permitido, motivo = amb.pode_gravar(plano.destino)
    if not permitido:
        if amb.simulando:
            # Em simulação isto NÃO é erro: é o modo funcionando como deveria.
            # Devolvemos PULADO para o assistente liberar as etapas seguintes —
            # senão o rascunho do e-mail nunca fica acessível para conferência.
            for op in plano.operacoes:
                estado.auditar(
                    acao=op.acao.value,
                    destino=op.destino,
                    origem=op.origem,
                    conta_id=plano.conta_id,
                    competencia=plano.competencia,
                    simulado=True,
                )
            return ResultadoEtapa.pulado(
                Etapa.PUBLICACAO,
                f"SIMULADO — {len(plano.operacoes)} alteração(ões) NÃO foram "
                f"aplicadas em {plano.destino.name}. Os arquivos continuam só "
                "na área local. Desligue seguranca.simulacao para gravar de verdade.",
                detalhes={
                    "simulacao": True,
                    "destino": str(plano.destino),
                    "operacoes": plano.resumo(),
                    "tem_sobrescrita": plano.tem_sobrescrita,
                },
            )
        return ResultadoEtapa.erro(
            Etapa.PUBLICACAO,
            f"escrita não autorizada em {plano.destino}: {motivo}",
            detalhes={"motivo": motivo, "simulacao": amb.simulando},
        )

    realizadas: list[str] = []
    falhas: list[str] = []
    artefatos: list[Path] = []

    for op in plano.operacoes:
        try:
            if op.acao is Acao.CRIAR_PASTA:
                op.destino.mkdir(parents=False, exist_ok=True)
                estado.auditar(
                    acao="criar_pasta",
                    destino=op.destino,
                    conta_id=plano.conta_id,
                    competencia=plano.competencia,
                )
                realizadas.append(op.descrever())
                continue

            backup = None
            if op.acao is Acao.SOBRESCREVER and amb.seguranca.backup_antes_de_sobrescrever:
                backup = _fazer_backup(op.destino, amb)

            shutil.copy2(op.origem, op.destino)
            artefatos.append(op.destino)
            estado.auditar(
                acao=op.acao.value,
                destino=op.destino,
                origem=op.origem,
                backup=backup,
                conta_id=plano.conta_id,
                competencia=plano.competencia,
                detalhes={"sha256": _hash(op.destino)},
            )
            realizadas.append(op.descrever())

        except OSError as exc:
            falhas.append(f"{op.destino.name}: {exc}")
            log.exception("falha ao publicar %s", op.destino)

    if falhas and not realizadas:
        return ResultadoEtapa.erro(
            Etapa.PUBLICACAO,
            "nenhuma alteração foi aplicada: " + "; ".join(falhas),
            detalhes={"falhas": falhas},
        )
    if falhas:
        return ResultadoEtapa.atencao(
            Etapa.PUBLICACAO,
            f"{len(realizadas)} alteração(ões) aplicadas, {len(falhas)} falharam",
            detalhes={"realizadas": realizadas, "falhas": falhas},
            artefatos=artefatos,
        )

    return ResultadoEtapa.sucesso(
        Etapa.PUBLICACAO,
        f"{len(realizadas)} alteração(ões) aplicadas em {plano.destino}",
        detalhes={"realizadas": realizadas, "destino": str(plano.destino)},
        artefatos=artefatos,
    )


def _fazer_backup(original: Path, amb: Ambiente) -> Path:
    """Copia o arquivo prestes a ser sobrescrito para dados/backups/."""
    carimbo = datetime.now().strftime("%Y%m%d-%H%M%S")
    pasta = amb.caminhos.backups / carimbo
    pasta.mkdir(parents=True, exist_ok=True)
    destino = pasta / original.name
    shutil.copy2(original, destino)
    log.info("backup: %s -> %s", original.name, destino)
    return destino


def restaurar_backup(backup: Path, destino: Path, *, confirmado: bool = False) -> str:
    """Desfaz uma sobrescrita. Também exige confirmação."""
    if not backup.is_file():
        raise FileNotFoundError(f"backup não encontrado: {backup}")
    if not confirmado:
        return f"restauraria {backup} sobre {destino} (nada foi feito)"
    shutil.copy2(backup, destino)
    estado.auditar(acao="restaurar_backup", destino=destino, origem=backup)
    return f"restaurado: {destino}"


if __name__ == "__main__":  # teste manual — só planeja, nunca publica
    from automacao import configurar_log
    from automacao.nucleo.config import conta_por_id, contas

    configurar_log()
    amb = ambiente()
    alvo = contas()[0]
    p = planejar(alvo, Competencia(2026, 8), [], ambiente_=amb)
    print(f"conta: {alvo.rotulo}")
    print(f"destino: {p.destino}")
    print(f"impedimentos: {p.impedimentos}")
