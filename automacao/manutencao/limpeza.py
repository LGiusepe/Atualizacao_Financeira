"""
Limpeza da bancada ("Painel Temp"): o que já cumpriu o papel pode sair.

A bancada mora no OneDrive desde 02/09/2026 e cresce todo mês — cada conta
deixa lá o xlsx, o PDF da autorização, o PDF final e as faturas de origem.
Tudo isso é *cópia*: o que vale de verdade foi publicado em "Autorizações de
pagamento" na etapa 6. Sem recolher, o espaço do OneDrive vai embora
guardando duas vezes a mesma coisa.

Duas frentes, com critérios diferentes:

* **Ao concluir uma conta** — as oito etapas fecharam, então o arquivo da
  entrada já está no destino. Some da entrada, e só ele.
* **Uma vez por mês** — varre a bancada inteira atrás de competência velha,
  backup antigo e resto de entrada.

As três coisas que este módulo nunca faz, e que são a razão de ele existir
como módulo separado em vez de um `unlink()` solto em qualquer lugar:

1. **Não apaga nada que não esteja provado no destino.** "Provado" é o
   conteúdo, não o nome: o arquivo da entrada foi copiado para a bancada e
   pode ter ganhado sufixo (`boleto_2.pdf`) no caminho. A conferência é por
   SHA-256, byte a byte.
2. **Não apaga nada em produção.** Todo caminho passa por
   `Ambiente.exigir_permissao_para_apagar`, que recusa "Autorizações de
   pagamento" e a planilha de controle sob qualquer configuração.
3. **Não apaga sem plano.** `planejar()` não escreve; `executar()` exige
   `confirmado=True`. É o mesmo desenho do publicador, e pelo mesmo motivo:
   quem opera precisa ler a lista antes.

Cada exclusão entra na trilha de auditoria (`acao="apagar"`), com o motivo e
o arquivo do destino que serviu de prova.
"""

from __future__ import annotations

import hashlib
import logging
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

from automacao.nucleo import estado
from automacao.nucleo.config import RAIZ_PROJETO, Ambiente, ambiente, contas
from automacao.nucleo.modelos import Competencia, Conta

log = logging.getLogger("automacao.manutencao.limpeza")

__all__ = [
    "Regras",
    "Alvo",
    "Mantido",
    "PlanoLimpeza",
    "planejar_conta",
    "planejar_mensal",
    "executar",
    "limpar_entrada_da_conta",
    "descartar_da_bancada",
    "ocupacao_da_bancada",
    "rodar_limpeza_mensal_se_for_a_hora",
    "ultima_limpeza",
]

#: Onde fica registrada a última varredura mensal. Em `dados/`, e não na
#: bancada: a marca não pode sumir junto com o que ela mede.
MARCA_ULTIMA_LIMPEZA = RAIZ_PROJETO / "dados" / "ultima-limpeza.txt"

#: Quantos backups ficam de pé mesmo vencidos. Ficar sem nenhum backup da
#: planilha por causa de um mês parado é pior que ocupar 300 KB.
BACKUPS_INTOCAVEIS = 3

GRUPOS = ("entrada", "trabalho", "backups")


# --------------------------------------------------------------------------- #
# Regras
# --------------------------------------------------------------------------- #


@dataclass(frozen=True)
class Regras:
    """O bloco `manutencao` do settings.yaml, com os padrões do projeto."""

    limpar_entrada_ao_concluir: bool = True
    limpeza_mensal: bool = True
    #: Quantas competências ficam em `trabalho/`. 2 = o mês atual e o anterior.
    meses_de_trabalho: int = 2
    dias_de_backup: int = 90

    @classmethod
    def de_ambiente(cls, amb: Ambiente) -> Regras:
        bruto = amb.manutencao or {}

        def inteiro(chave: str, padrao: int, minimo: int) -> int:
            try:
                valor = int(bruto.get(chave, padrao))
            except (TypeError, ValueError):
                return padrao
            # Zero aqui significaria "apague o mês corrente enquanto ele
            # ainda está sendo feito". Não existe configuração que justifique.
            return max(minimo, valor)

        return cls(
            limpar_entrada_ao_concluir=bool(
                bruto.get("limpar_entrada_ao_concluir", True)
            ),
            limpeza_mensal=bool(bruto.get("limpeza_mensal", True)),
            meses_de_trabalho=inteiro("meses_de_trabalho", 2, 1),
            dias_de_backup=inteiro("dias_de_backup", 90, 7),
        )


# --------------------------------------------------------------------------- #
# Plano
# --------------------------------------------------------------------------- #


@dataclass
class Alvo:
    """Um arquivo que sairia da bancada, e por quê."""

    caminho: Path
    grupo: str
    motivo: str
    tamanho: int = 0
    conta_id: str | None = None
    competencia: Competencia | None = None
    #: O arquivo do destino que prova que este aqui é dispensável.
    prova: Path | None = None


@dataclass
class Mantido:
    """Um arquivo que ficou de fora do plano — e o motivo, que é o que importa."""

    caminho: Path
    grupo: str
    motivo: str
    tamanho: int = 0


@dataclass
class PlanoLimpeza:
    alvos: list[Alvo] = field(default_factory=list)
    mantidos: list[Mantido] = field(default_factory=list)

    @property
    def vazio(self) -> bool:
        return not self.alvos

    @property
    def bytes_a_liberar(self) -> int:
        return sum(a.tamanho for a in self.alvos)

    def por_grupo(self) -> dict[str, dict[str, int]]:
        """`{'trabalho': {'arquivos': 12, 'bytes': 3_400_000}, ...}`."""
        resumo = {g: {"arquivos": 0, "bytes": 0} for g in GRUPOS}
        for alvo in self.alvos:
            linha = resumo.setdefault(alvo.grupo, {"arquivos": 0, "bytes": 0})
            linha["arquivos"] += 1
            linha["bytes"] += alvo.tamanho
        return resumo

    def somar(self, outro: PlanoLimpeza) -> PlanoLimpeza:
        ja_listados = {a.caminho for a in self.alvos}
        self.alvos.extend(a for a in outro.alvos if a.caminho not in ja_listados)
        self.mantidos.extend(outro.mantidos)
        return self


@dataclass
class ResultadoLimpeza:
    apagados: list[Path] = field(default_factory=list)
    falhas: list[str] = field(default_factory=list)
    bytes_liberados: int = 0
    pastas_removidas: int = 0

    @property
    def ok(self) -> bool:
        return not self.falhas


# --------------------------------------------------------------------------- #
# Conferência no destino
# --------------------------------------------------------------------------- #


def _sha256(caminho: Path) -> str | None:
    try:
        h = hashlib.sha256()
        with caminho.open("rb") as fh:
            for bloco in iter(lambda: fh.read(1 << 20), b""):
                h.update(bloco)
        return h.hexdigest()
    except OSError as exc:
        log.warning("não consegui ler %s para conferir: %s", caminho, exc)
        return None


def _tamanho(caminho: Path) -> int:
    try:
        return caminho.stat().st_size
    except OSError:
        return 0


def gemeo_no_destino(arquivo: Path, destino: Path) -> Path | None:
    """
    O arquivo de `destino` com o MESMO conteúdo de `arquivo`, se existir.

    Casa por conteúdo, não por nome, de propósito: `copiar_sem_duplicar` dá
    sufixo a nome repetido (`boleto.pdf` → `boleto_2.pdf`), então o que foi
    publicado pode ter outro nome que o original da entrada. Comparar nomes
    deixaria de reconhecer a cópia e a bancada nunca esvaziaria — ou, pior,
    reconheceria a cópia errada.

    A pasta do mês tem meia dúzia de arquivos, então filtrar por tamanho antes
    de calcular o SHA basta para isto não pesar.
    """
    if not destino.is_dir():
        return None
    tamanho = _tamanho(arquivo)
    if not tamanho:
        return None

    candidatos = [
        c for c in destino.iterdir() if c.is_file() and _tamanho(c) == tamanho
    ]
    if not candidatos:
        return None

    digital = _sha256(arquivo)
    if digital is None:
        return None
    for candidato in candidatos:
        if _sha256(candidato) == digital:
            return candidato
    return None


# --------------------------------------------------------------------------- #
# Frente 1 — a conta que acabou de fechar
# --------------------------------------------------------------------------- #


def _arquivos_da_entrada(conta_id: str, amb: Ambiente) -> list[Path]:
    """Os arquivos de `entrada/<conta_id>/`, fora de `_processados/`."""
    from automacao.coleta.manual import EXTENSOES_ACEITAS, PASTA_PROCESSADOS

    pasta = amb.caminhos.entrada / conta_id
    if not pasta.is_dir():
        return []
    return sorted(
        arquivo
        for arquivo in pasta.rglob("*")
        if arquivo.is_file()
        and arquivo.suffix.lower() in EXTENSOES_ACEITAS
        and PASTA_PROCESSADOS not in arquivo.parts
    )


def planejar_conta(
    conta: Conta,
    competencia: Competencia,
    *,
    ambiente_: Ambiente | None = None,
    exigir_conclusao: bool = True,
) -> PlanoLimpeza:
    """
    O que sairia da entrada desta conta neste mês.

    Só entra no plano o arquivo que está **provado** na pasta de destino, com
    o mesmo conteúdo. O que não estiver lá vai para `mantidos`, com o motivo —
    é a diferença entre "a bancada está limpa" e "o arquivo sumiu".

    Args:
        conta: a conta.
        competencia: o mês.
        ambiente_: ambiente já carregado (útil em teste).
        exigir_conclusao: com `True` (o padrão), mês que não fechou as oito
            etapas não perde nada. Desligar só faz sentido na varredura
            mensal de meses antigos, onde a prova no destino já basta.

    Returns:
        `PlanoLimpeza`. Nada foi tocado em disco.
    """
    amb = ambiente_ or ambiente()
    plano = PlanoLimpeza()

    arquivos = _arquivos_da_entrada(conta.id, amb)
    if not arquivos:
        return plano

    proc = estado.carregar(conta.id, competencia)
    if exigir_conclusao and not proc.concluido:
        for arquivo in arquivos:
            plano.mantidos.append(
                Mantido(
                    arquivo,
                    "entrada",
                    f"{conta.rotulo} / {competencia}: o processo ainda não fechou "
                    "as oito etapas",
                    _tamanho(arquivo),
                )
            )
        return plano

    destino = amb.destino_onedrive(conta, competencia)
    for arquivo in arquivos:
        prova = gemeo_no_destino(arquivo, destino)
        if prova is None:
            plano.mantidos.append(
                Mantido(
                    arquivo,
                    "entrada",
                    f"não encontrei este arquivo em {destino}",
                    _tamanho(arquivo),
                )
            )
            continue
        plano.alvos.append(
            Alvo(
                caminho=arquivo,
                grupo="entrada",
                motivo=f"publicado em {destino.name} como {prova.name}",
                tamanho=_tamanho(arquivo),
                conta_id=conta.id,
                competencia=competencia,
                prova=prova,
            )
        )
    return plano


def limpar_entrada_da_conta(
    conta: Conta,
    competencia: Competencia,
    *,
    ambiente_: Ambiente | None = None,
) -> ResultadoLimpeza:
    """
    Apaga da entrada o que esta conta já publicou — o gancho do fim do fluxo.

    Chamado pelo orquestrador quando a última etapa fecha. Respeita
    `manutencao.limpar_entrada_ao_concluir`: desligado, não faz nada e não
    reclama.

    Reprocessar depois disso continua funcionando: `etapa_coleta` procura em
    três lugares, e o terceiro é a própria pasta de trabalho — onde as cópias
    continuam até a limpeza mensal levar a competência inteira, já publicada.
    """
    amb = ambiente_ or ambiente()
    regras = Regras.de_ambiente(amb)
    if not regras.limpar_entrada_ao_concluir:
        return ResultadoLimpeza()

    plano = planejar_conta(conta, competencia, ambiente_=amb)
    if plano.vazio:
        return ResultadoLimpeza()
    return executar(plano, confirmado=True, ambiente_=amb)


def descartar_da_bancada(
    caminho: Path,
    *,
    conta_id: str | None = None,
    competencia: Competencia | None = None,
    ambiente_: Ambiente | None = None,
) -> ResultadoLimpeza:
    """
    Tira UM arquivo da pasta de trabalho — o desfazer da escolha errada.

    O painel chama isto quando o operador marcou na etapa 1 um arquivo que não
    era desta fatura. Sem ele a escolha errada é definitiva: a cópia fica na
    pasta de trabalho e o passo 3 da coleta a traz de volta a cada execução.

    Não é a limpeza normal, e por isso não passa por `planejar_*`: ali um alvo
    só cai quando o mesmo conteúdo está provado no destino, e aqui é o
    contrário — o arquivo sai justamente porque **não** devia ser publicado.
    A trava de exclusão continua valendo, aplicada por `executar()`, e o
    `e_area_local` abaixo estreita ainda mais: só a bancada, nunca produção,
    nunca o resto do disco.

    O original em `dados/entrada/` não é tocado: se o operador errou no
    descarte, ele marca de novo na tela.
    """
    amb = ambiente_ or ambiente()
    alvo = Path(caminho)

    if not amb.e_area_local(alvo):
        raise PermissionError(
            f"descarte recusado: {alvo} está fora das pastas do painel"
        )

    plano = PlanoLimpeza()
    plano.alvos.append(
        Alvo(
            caminho=alvo,
            grupo="descarte",
            motivo="operador descartou o documento na etapa de coleta",
            tamanho=_tamanho(alvo),
            conta_id=conta_id,
            competencia=competencia,
        )
    )
    return executar(plano, confirmado=True, ambiente_=amb)


# --------------------------------------------------------------------------- #
# Frente 2 — a varredura mensal
# --------------------------------------------------------------------------- #


def _competencia_da_pasta(pasta: Path) -> Competencia | None:
    try:
        return Competencia.de_texto(pasta.name)
    except ValueError:
        return None


def _limite_de_trabalho(regras: Regras, hoje: date) -> Competencia:
    """A competência mais antiga que continua na bancada."""
    return Competencia.de_data(hoje).somar(-(regras.meses_de_trabalho - 1))


def _planejar_trabalho(amb: Ambiente, regras: Regras, hoje: date) -> PlanoLimpeza:
    """
    Competências antigas de `trabalho/`.

    Duas conversas diferentes com o mesmo mês:

    * a conta **fechou as oito etapas** — está tudo publicado, e a pasta
      inteira pode ir, inclusive o PDF intermediário da autorização, que
      nunca foi para o destino e por isso nunca teria prova;
    * a conta **não fechou** — sai só o que estiver provado no destino. O
      resto fica, listado em `mantidos`. Mês velho e inacabado costuma ser
      exatamente o que alguém vai querer reabrir.
    """
    plano = PlanoLimpeza()
    raiz = amb.caminhos.trabalho
    if not raiz.is_dir():
        return plano

    limite = _limite_de_trabalho(regras, hoje)
    por_id = {c.id: c for c in contas()}

    for pasta_mes in sorted(p for p in raiz.iterdir() if p.is_dir()):
        competencia = _competencia_da_pasta(pasta_mes)
        if competencia is None:
            plano.mantidos.append(
                Mantido(pasta_mes, "trabalho", "nome de pasta não é uma competência")
            )
            continue
        if (competencia.ano, competencia.mes) >= (limite.ano, limite.mes):
            continue  # dentro da janela que você pediu para manter

        for pasta_conta in sorted(p for p in pasta_mes.iterdir() if p.is_dir()):
            conta = por_id.get(pasta_conta.name)
            arquivos = [a for a in pasta_conta.rglob("*") if a.is_file()]
            if not arquivos:
                continue

            if conta is None:
                # Conta que saiu do registro: sem `Conta` não há como montar o
                # caminho do destino, então não há como provar nada.
                for arquivo in arquivos:
                    plano.mantidos.append(
                        Mantido(
                            arquivo,
                            "trabalho",
                            f"{pasta_conta.name} não está mais no registro de contas",
                            _tamanho(arquivo),
                        )
                    )
                continue

            proc = estado.carregar(conta.id, competencia)
            if proc.concluido:
                for arquivo in arquivos:
                    plano.alvos.append(
                        Alvo(
                            caminho=arquivo,
                            grupo="trabalho",
                            motivo=f"{competencia} concluída — está em "
                            f"{amb.destino_onedrive(conta, competencia)}",
                            tamanho=_tamanho(arquivo),
                            conta_id=conta.id,
                            competencia=competencia,
                        )
                    )
                continue

            destino = amb.destino_onedrive(conta, competencia)
            for arquivo in arquivos:
                prova = gemeo_no_destino(arquivo, destino)
                if prova is None:
                    plano.mantidos.append(
                        Mantido(
                            arquivo,
                            "trabalho",
                            f"{conta.rotulo} / {competencia} não fechou e este "
                            "arquivo não está no destino",
                            _tamanho(arquivo),
                        )
                    )
                    continue
                plano.alvos.append(
                    Alvo(
                        caminho=arquivo,
                        grupo="trabalho",
                        motivo=f"publicado em {destino.name} como {prova.name}",
                        tamanho=_tamanho(arquivo),
                        conta_id=conta.id,
                        competencia=competencia,
                        prova=prova,
                    )
                )

    return plano


def _planejar_entrada(amb: Ambiente, regras: Regras, hoje: date) -> PlanoLimpeza:
    """
    Restos da entrada: o que ficou para trás e os `_processados/` vencidos.

    A varredura por conta não olha só o mês corrente — o arquivo pode ter
    ficado na entrada desde antes de a limpeza existir. Para cada conta com
    arquivo parado, conferimos as competências que o banco conhece dela.
    """
    from automacao.coleta.manual import PASTA_PROCESSADOS

    plano = PlanoLimpeza()
    raiz = amb.caminhos.entrada
    if not raiz.is_dir():
        return plano

    limite = _limite_de_trabalho(regras, hoje)

    for conta in contas():
        arquivos = _arquivos_da_entrada(conta.id, amb)
        if not arquivos:
            continue

        meses = estado.competencias_da_conta(conta.id) or [Competencia.atual()]
        provados: dict[Path, Alvo] = {}
        for competencia in meses:
            parcial = planejar_conta(
                conta, competencia, ambiente_=amb, exigir_conclusao=True
            )
            for alvo in parcial.alvos:
                provados.setdefault(alvo.caminho, alvo)

        plano.alvos.extend(provados.values())
        # Um motivo por arquivo, e não um por competência tentada: a mesma
        # fatura apareceria quatro vezes na lista de mantidos, dizendo o
        # mesmo em quatro meses diferentes.
        for arquivo in arquivos:
            if arquivo in provados:
                continue
            plano.mantidos.append(
                Mantido(
                    arquivo,
                    "entrada",
                    f"{conta.rotulo}: nenhum mês concluído tem este arquivo no destino",
                    _tamanho(arquivo),
                )
            )

    # `_processados/<competencia>/` — o que `arquivar_entrada()` guardou.
    guarda = raiz / PASTA_PROCESSADOS
    if guarda.is_dir():
        for pasta_mes in sorted(p for p in guarda.iterdir() if p.is_dir()):
            competencia = _competencia_da_pasta(pasta_mes)
            if competencia is None:
                continue
            if (competencia.ano, competencia.mes) >= (limite.ano, limite.mes):
                continue
            for arquivo in sorted(a for a in pasta_mes.rglob("*") if a.is_file()):
                plano.alvos.append(
                    Alvo(
                        caminho=arquivo,
                        grupo="entrada",
                        motivo=f"já arquivado de {competencia}, fora da janela",
                        tamanho=_tamanho(arquivo),
                        competencia=competencia,
                    )
                )

    return plano


def _planejar_backups(amb: Ambiente, regras: Regras, hoje: date) -> PlanoLimpeza:
    """
    Backups vencidos — com o chão de `BACKUPS_INTOCAVEIS`.

    Aqui não existe prova no destino: backup é justamente a cópia do que foi
    substituído. O critério é idade, e o chão existe para que um mês sem
    gravação nenhuma não deixe a pasta zerada.
    """
    plano = PlanoLimpeza()
    raiz = amb.caminhos.backups
    if not raiz.is_dir():
        return plano

    arquivos = sorted(
        (a for a in raiz.rglob("*") if a.is_file()),
        key=lambda a: a.stat().st_mtime if a.exists() else 0,
        reverse=True,
    )
    protegidos = {a for a in arquivos[:BACKUPS_INTOCAVEIS]}
    corte = datetime.combine(hoje, datetime.min.time()).timestamp() - (
        regras.dias_de_backup * 86400
    )

    for arquivo in arquivos:
        try:
            quando = arquivo.stat().st_mtime
        except OSError:
            continue
        if arquivo in protegidos:
            plano.mantidos.append(
                Mantido(arquivo, "backups", "um dos mais recentes", _tamanho(arquivo))
            )
            continue
        if quando >= corte:
            continue
        idade = int((datetime.now().timestamp() - quando) / 86400)
        plano.alvos.append(
            Alvo(
                caminho=arquivo,
                grupo="backups",
                motivo=f"{idade} dias — passou dos {regras.dias_de_backup} configurados",
                tamanho=_tamanho(arquivo),
            )
        )
    return plano


def planejar_mensal(
    *, ambiente_: Ambiente | None = None, hoje: date | None = None
) -> PlanoLimpeza:
    """
    A varredura inteira da bancada. Não escreve nada.

    Args:
        ambiente_: ambiente já carregado (útil em teste).
        hoje: data de referência (útil em teste); padrão é `date.today()`.
    """
    amb = ambiente_ or ambiente()
    regras = Regras.de_ambiente(amb)
    quando = hoje or date.today()

    plano = PlanoLimpeza()
    plano.somar(_planejar_trabalho(amb, regras, quando))
    plano.somar(_planejar_entrada(amb, regras, quando))
    plano.somar(_planejar_backups(amb, regras, quando))
    return plano


# --------------------------------------------------------------------------- #
# Execução
# --------------------------------------------------------------------------- #


def executar(
    plano: PlanoLimpeza,
    *,
    confirmado: bool = False,
    ambiente_: Ambiente | None = None,
) -> ResultadoLimpeza:
    """
    Apaga o que está no plano. Sem `confirmado=True`, não toca em nada.

    Cada caminho passa pela trava de exclusão ANTES do `unlink()`, um a um —
    e não uma vez no começo. Um plano montado com a configuração de agora
    poderia ser executado depois de alguém ter mexido em `caminhos.trabalho`;
    a trava por item é o que garante que a produção continua fora de alcance
    mesmo nessa janela.
    """
    amb = ambiente_ or ambiente()
    resultado = ResultadoLimpeza()
    if not confirmado:
        return resultado

    pastas_tocadas: set[Path] = set()

    for alvo in plano.alvos:
        try:
            amb.exigir_permissao_para_apagar(alvo.caminho)
        except PermissionError as exc:
            # Não é só "pulei um arquivo": é a trava impedindo algo que o
            # planejamento deixou passar. Vale aparecer no relato.
            resultado.falhas.append(f"{alvo.caminho}: {exc}")
            log.error("trava de exclusão recusou %s", alvo.caminho)
            continue

        tamanho = _tamanho(alvo.caminho)
        try:
            alvo.caminho.unlink()
        except FileNotFoundError:
            continue  # alguém já tirou; o objetivo era esse mesmo
        except OSError as exc:
            resultado.falhas.append(f"{alvo.caminho.name}: {exc}")
            log.warning("não consegui apagar %s: %s", alvo.caminho, exc)
            continue

        resultado.apagados.append(alvo.caminho)
        resultado.bytes_liberados += tamanho
        pastas_tocadas.add(alvo.caminho.parent)
        estado.auditar(
            acao="apagar",
            destino=alvo.caminho,
            conta_id=alvo.conta_id,
            competencia=alvo.competencia,
            detalhes={
                "grupo": alvo.grupo,
                "motivo": alvo.motivo,
                "bytes": tamanho,
                "prova": str(alvo.prova) if alvo.prova else None,
            },
        )

    resultado.pastas_removidas = _remover_pastas_vazias(pastas_tocadas, amb)
    log.info(
        "limpeza: %d arquivo(s), %.1f MB, %d pasta(s) vazia(s) removida(s)",
        len(resultado.apagados),
        resultado.bytes_liberados / 1024 / 1024,
        resultado.pastas_removidas,
    )
    return resultado


def _remover_pastas_vazias(pastas: set[Path], amb: Ambiente) -> int:
    """
    Recolhe a casca: pasta de conta e de mês que ficaram sem nada dentro.

    Sobe enquanto a pasta estiver vazia e continuar sendo bancada. Duas
    exceções, e as duas são sobre não assustar quem opera:

    * as raízes configuradas (trabalho, entrada, backups) ficam — o painel as
      recria na subida seguinte, mas vê-las sumir do OneDrive assusta;
    * **`entrada/<conta>/` também fica**, ainda que vazia. É a caixa onde
      você larga o PDF baixado no portal, e o vínculo com a conta vem do
      nome dela. Apagar a caixa junto com o conteúdo deixaria você abrindo o
      Explorador atrás de uma pasta que a automação comeu.
    """
    raizes = {p.resolve() for p in amb.pastas_do_painel()}
    try:
        raiz_entrada = amb.caminhos.entrada.resolve()
        raizes |= {p.resolve() for p in raiz_entrada.iterdir() if p.is_dir()}
    except OSError:
        pass
    removidas = 0
    for pasta in sorted(pastas, key=lambda p: len(p.parts), reverse=True):
        cursor = pasta
        while True:
            try:
                if cursor.resolve() in raizes or not cursor.is_dir():
                    break
                if any(cursor.iterdir()):
                    break
                amb.exigir_permissao_para_apagar(cursor)
                cursor.rmdir()
            except (PermissionError, OSError):
                break
            removidas += 1
            cursor = cursor.parent
    return removidas


# --------------------------------------------------------------------------- #
# Quanto a bancada está ocupando
# --------------------------------------------------------------------------- #


def ocupacao_da_bancada(*, ambiente_: Ambiente | None = None) -> dict[str, dict]:
    """`{'trabalho': {'arquivos': n, 'bytes': n, 'caminho': Path}, ...}`."""
    amb = ambiente_ or ambiente()
    resumo: dict[str, dict] = {}
    for grupo in GRUPOS:
        raiz = getattr(amb.caminhos, grupo)
        arquivos = 0
        bytes_ = 0
        if raiz.is_dir():
            for arquivo in raiz.rglob("*"):
                if arquivo.is_file():
                    arquivos += 1
                    bytes_ += _tamanho(arquivo)
        resumo[grupo] = {"arquivos": arquivos, "bytes": bytes_, "caminho": raiz}
    return resumo


# --------------------------------------------------------------------------- #
# A varredura automática, uma vez por mês
# --------------------------------------------------------------------------- #


def ultima_limpeza() -> str:
    """A competência em que a varredura rodou pela última vez, ou ''."""
    try:
        return MARCA_ULTIMA_LIMPEZA.read_text(encoding="utf-8").strip()
    except OSError:
        return ""


def _marcar_rodada(competencia: Competencia) -> None:
    try:
        MARCA_ULTIMA_LIMPEZA.parent.mkdir(parents=True, exist_ok=True)
        MARCA_ULTIMA_LIMPEZA.write_text(str(competencia), encoding="utf-8")
    except OSError as exc:  # noqa: BLE001 — não registrar não pode derrubar o painel
        log.warning("não consegui gravar a marca da limpeza: %s", exc)


def rodar_limpeza_mensal_se_for_a_hora(
    *, ambiente_: Ambiente | None = None, hoje: date | None = None
) -> ResultadoLimpeza | None:
    """
    Roda a varredura no primeiro acesso de cada mês. Devolve `None` se não era a hora.

    Chamada na subida do painel. A marca é gravada mesmo quando nada foi
    apagado: o combinado é "uma varredura por mês", não "uma varredura por
    subida até achar alguma coisa".
    """
    amb = ambiente_ or ambiente()
    regras = Regras.de_ambiente(amb)
    if not regras.limpeza_mensal:
        return None

    agora = Competencia.de_data(hoje or date.today())
    if ultima_limpeza() == str(agora):
        return None

    plano = planejar_mensal(ambiente_=amb, hoje=hoje)
    resultado = executar(plano, confirmado=True, ambiente_=amb)
    _marcar_rodada(agora)
    if resultado.apagados:
        log.info(
            "limpeza mensal de %s: %d arquivo(s), %.1f MB liberados",
            agora,
            len(resultado.apagados),
            resultado.bytes_liberados / 1024 / 1024,
        )
    return resultado
