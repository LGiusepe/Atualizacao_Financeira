"""
Estado do processamento em SQLite.

Guarda, por (conta, competência), o que já foi feito e o que ainda falta.
É o que permite parar no meio, fechar o painel e retomar depois sem refazer
trabalho — e é a memória que evita mandar a mesma fatura duas vezes.

Também mantém uma trilha de auditoria de tudo que foi gravado no OneDrive.
"""

from __future__ import annotations

import json
import logging
import sqlite3
from contextlib import contextmanager
from dataclasses import asdict
from datetime import date, datetime
from pathlib import Path
from typing import Iterator

from automacao.nucleo.config import RAIZ_PROJETO
from automacao.nucleo.modelos import (
    Competencia,
    Documento,
    Etapa,
    OrigemDocumento,
    Processamento,
    ResultadoEtapa,
    Situacao,
    TipoDocumento,
)

log = logging.getLogger("automacao.nucleo.estado")

CAMINHO_BANCO = RAIZ_PROJETO / "dados" / "estado.db"

ESQUEMA = """
CREATE TABLE IF NOT EXISTS processamento (
    conta_id        TEXT NOT NULL,
    competencia     TEXT NOT NULL,
    valor           REAL,
    vencimento      TEXT,
    numero_documento TEXT,
    -- Observações reescritas por você no painel, só para este mês.
    descricao       TEXT,
    -- Escolhas da etapa da autorização, idem: valem só para este mês.
    pagante         TEXT,
    beneficiario    TEXT,
    meio_pgto       TEXT,
    forma_pgto      TEXT,
    xlsx            TEXT,
    pdf_final       TEXT,
    email_enviado   INTEGER NOT NULL DEFAULT 0,
    -- Data em que a fatura foi enviada ao financeiro. Preenchida quando o
    -- envio aconteceu fora da ferramenta e você registrou isso no painel.
    enviado_em      TEXT,
    enviado_por_fora INTEGER NOT NULL DEFAULT 0,
    criado_em       TEXT NOT NULL,
    atualizado_em   TEXT NOT NULL,
    PRIMARY KEY (conta_id, competencia)
);

CREATE TABLE IF NOT EXISTS etapa (
    conta_id      TEXT NOT NULL,
    competencia   TEXT NOT NULL,
    etapa         TEXT NOT NULL,
    situacao      TEXT NOT NULL,
    mensagem      TEXT NOT NULL DEFAULT '',
    detalhes      TEXT NOT NULL DEFAULT '{}',
    artefatos     TEXT NOT NULL DEFAULT '[]',
    atualizado_em TEXT NOT NULL,
    PRIMARY KEY (conta_id, competencia, etapa)
);

CREATE TABLE IF NOT EXISTS documento (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    conta_id      TEXT NOT NULL,
    competencia   TEXT NOT NULL,
    caminho       TEXT NOT NULL,
    tipo          TEXT NOT NULL,
    origem        TEXT NOT NULL,
    confianca     REAL NOT NULL DEFAULT 0,
    motivo        TEXT NOT NULL DEFAULT '',
    paginas       INTEGER,
    valor         REAL,
    vencimento    TEXT,
    numero_documento TEXT,
    remetente     TEXT,
    assunto_email TEXT,
    UNIQUE (conta_id, competencia, caminho)
);

-- Trilha de auditoria: toda escrita fora de dados/ passa por aqui.
CREATE TABLE IF NOT EXISTS auditoria (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    quando      TEXT NOT NULL,
    conta_id    TEXT,
    competencia TEXT,
    acao        TEXT NOT NULL,
    destino     TEXT NOT NULL,
    origem      TEXT,
    backup      TEXT,
    simulado    INTEGER NOT NULL DEFAULT 0,
    detalhes    TEXT NOT NULL DEFAULT '{}'
);

-- Quais contas valem em cada mês. O conjunto muda: julho teve 50 linhas no
-- checklist, agosto tem 28. Sem isto, desativar uma conta apagaria o histórico
-- dos meses em que ela existia.
CREATE TABLE IF NOT EXISTS conta_do_mes (
    competencia TEXT NOT NULL,
    conta_id    TEXT NOT NULL,
    ativa       INTEGER NOT NULL DEFAULT 1,
    definido_em TEXT NOT NULL,
    PRIMARY KEY (competencia, conta_id)
);

CREATE INDEX IF NOT EXISTS idx_etapa_comp ON etapa (competencia);
CREATE INDEX IF NOT EXISTS idx_auditoria_quando ON auditoria (quando DESC);
"""


def _agora() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _momento(bruto: str | None) -> datetime | None:
    """Texto ISO do banco de volta para `datetime`; formato estranho vira None."""
    if not bruto:
        return None
    try:
        return datetime.fromisoformat(bruto)
    except ValueError:
        return None


#: Colunas acrescentadas depois que o banco já existia. `CREATE TABLE IF NOT
#: EXISTS` não altera tabela criada antes — sem isto, quem já tinha
#: `estado.db` continuaria sem a coluna nova e a gravação falharia.
COLUNAS_NOVAS = {
    "processamento": {
        "descricao": "TEXT",
        "pagante": "TEXT",
        "beneficiario": "TEXT",
        "meio_pgto": "TEXT",
        "forma_pgto": "TEXT",
    },
}


def _acrescentar_colunas_novas(con: sqlite3.Connection) -> None:
    for tabela, colunas in COLUNAS_NOVAS.items():
        existentes = {linha["name"] for linha in con.execute(f"PRAGMA table_info({tabela})")}
        for nome, tipo in colunas.items():
            if nome not in existentes:
                con.execute(f"ALTER TABLE {tabela} ADD COLUMN {nome} {tipo}")
                log.info("banco: coluna %s.%s criada", tabela, nome)


@contextmanager
def conexao(caminho: Path | None = None) -> Iterator[sqlite3.Connection]:
    alvo = caminho or CAMINHO_BANCO
    alvo.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(alvo)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys = ON")
    try:
        con.executescript(ESQUEMA)
        _acrescentar_colunas_novas(con)
        yield con
        con.commit()
    except Exception:
        con.rollback()
        raise
    finally:
        con.close()


# --------------------------------------------------------------------------- #
# Gravação
# --------------------------------------------------------------------------- #


def salvar_etapa(
    conta_id: str, competencia: Competencia, resultado: ResultadoEtapa
) -> None:
    with conexao() as con:
        _garantir_processamento(con, conta_id, competencia)
        con.execute(
            """
            INSERT INTO etapa (conta_id, competencia, etapa, situacao, mensagem,
                               detalhes, artefatos, atualizado_em)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT (conta_id, competencia, etapa) DO UPDATE SET
                situacao = excluded.situacao,
                mensagem = excluded.mensagem,
                detalhes = excluded.detalhes,
                artefatos = excluded.artefatos,
                atualizado_em = excluded.atualizado_em
            """,
            (
                conta_id,
                str(competencia),
                resultado.etapa.value,
                resultado.situacao.value,
                resultado.mensagem,
                json.dumps(resultado.detalhes, ensure_ascii=False, default=str),
                json.dumps([str(a) for a in resultado.artefatos], ensure_ascii=False),
                _agora(),
            ),
        )
        con.execute(
            "UPDATE processamento SET atualizado_em = ? WHERE conta_id = ? AND competencia = ?",
            (_agora(), conta_id, str(competencia)),
        )


def salvar_dados(
    conta_id: str,
    competencia: Competencia,
    *,
    valor: float | None = None,
    vencimento=None,
    numero_documento: str | None = None,
    descricao: str | None = None,
    pagante: str | None = None,
    beneficiario: str | None = None,
    meio_pgto: str | None = None,
    forma_pgto: str | None = None,
    xlsx: Path | None = None,
    pdf_final: Path | None = None,
    email_enviado: bool | None = None,
) -> None:
    """Atualiza só os campos informados (None = não mexe)."""
    campos = {
        "valor": valor,
        "vencimento": vencimento.isoformat() if vencimento else None,
        "numero_documento": numero_documento,
        "descricao": descricao,
        "pagante": pagante,
        "beneficiario": beneficiario,
        "meio_pgto": meio_pgto,
        "forma_pgto": forma_pgto,
        "xlsx": str(xlsx) if xlsx else None,
        "pdf_final": str(pdf_final) if pdf_final else None,
        "email_enviado": int(email_enviado) if email_enviado is not None else None,
    }
    campos = {k: v for k, v in campos.items() if v is not None}
    if not campos:
        return

    with conexao() as con:
        _garantir_processamento(con, conta_id, competencia)
        atribuicoes = ", ".join(f"{k} = ?" for k in campos)
        con.execute(
            f"UPDATE processamento SET {atribuicoes}, atualizado_em = ? "
            "WHERE conta_id = ? AND competencia = ?",
            (*campos.values(), _agora(), conta_id, str(competencia)),
        )


def registrar_envio_externo(
    conta_id: str, competencia: Competencia, quando, observacao: str = ""
) -> None:
    """
    Marca a conta como já enviada, com envio feito fora da ferramenta.

    Existe porque o processo não nasceu na automação: boa parte do mês já foi
    despachada na mão, e sem isso o painel mostraria como pendente algo que o
    financeiro já recebeu. Todas as etapas viram PULADO, com a data no texto.
    """
    from automacao.nucleo.modelos import ORDEM_ETAPAS_PADRAO

    nota = f" — {observacao}" if observacao else ""
    mensagem = (
        f"enviada fora da ferramenta em {quando.strftime('%d/%m/%Y')}{nota}"
    )

    with conexao() as con:
        _garantir_processamento(con, conta_id, competencia)
        con.execute(
            "UPDATE processamento SET email_enviado = 1, enviado_em = ?, "
            "enviado_por_fora = 1, atualizado_em = ? "
            "WHERE conta_id = ? AND competencia = ?",
            (quando.isoformat(), _agora(), conta_id, str(competencia)),
        )

    for etapa in ORDEM_ETAPAS_PADRAO:
        salvar_etapa(
            conta_id,
            competencia,
            ResultadoEtapa.pulado(etapa, mensagem, detalhes={"envio_externo": True}),
        )
    log.info("%s/%s registrada como enviada fora da ferramenta", conta_id, competencia)


def registrar_envio_pela_ferramenta(
    conta_id: str, competencia: Competencia, quando: date
) -> None:
    """
    Marca que a mensagem saiu — despachada pela própria automação.

    Diferente de `registrar_envio_externo`: aqui `enviado_por_fora` continua 0,
    porque quem mandou foi a ferramenta. As duas gravam `enviado_em`, e é dele
    que a lista do mês tira a data de "enviada em".

    Este registro é o que impede o segundo envio: `etapa_email` consulta antes
    de falar com o Outlook. E-mail despachado não tem volta.
    """
    with conexao() as con:
        _garantir_processamento(con, conta_id, competencia)
        con.execute(
            "UPDATE processamento SET email_enviado = 1, enviado_em = ?, "
            "atualizado_em = ? WHERE conta_id = ? AND competencia = ?",
            (quando.isoformat(), _agora(), conta_id, str(competencia)),
        )
    log.info("%s/%s: e-mail despachado em %s", conta_id, competencia, quando)


def desfazer_envio_externo(conta_id: str, competencia: Competencia) -> None:
    """Volta atrás: limpa o registro e reabre as etapas."""
    with conexao() as con:
        con.execute(
            "UPDATE processamento SET email_enviado = 0, enviado_em = NULL, "
            "enviado_por_fora = 0, atualizado_em = ? "
            "WHERE conta_id = ? AND competencia = ?",
            (_agora(), conta_id, str(competencia)),
        )
    reabrir(conta_id, competencia)
    log.info("%s/%s: registro de envio externo desfeito", conta_id, competencia)


def salvar_documentos(
    conta_id: str, competencia: Competencia, documentos: list[Documento]
) -> None:
    """
    Grava a lista de documentos da competência — e só ela.

    **Substitui, não acumula.** Quem chama sempre monta a lista inteira do que
    a competência tem agora, então o que não está nela tem de sair da tabela.

    Antes era só INSERT com `ON CONFLICT DO UPDATE`: linha que ficava de fora
    da lista simplesmente permanecia. O sintoma era o documento que não some da
    etapa 1 — o arquivo era apagado do disco, a tela continuava mostrando, e
    clicar em "descartar" de novo não adiantava nada, porque o descarte já
    tinha feito a sua parte e o que sobrava era a linha órfã no banco.

    Lista vazia apaga tudo da competência, que é o certo: é o que a coleta
    grava quando não acha documento nenhum.
    """
    with conexao() as con:
        _garantir_processamento(con, conta_id, competencia)
        for d in documentos:
            con.execute(
                """
                INSERT INTO documento (conta_id, competencia, caminho, tipo, origem,
                                       confianca, motivo, paginas, valor, vencimento,
                                       numero_documento, remetente, assunto_email)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT (conta_id, competencia, caminho) DO UPDATE SET
                    tipo = excluded.tipo,
                    origem = excluded.origem,
                    confianca = excluded.confianca,
                    motivo = excluded.motivo,
                    paginas = excluded.paginas,
                    valor = excluded.valor,
                    vencimento = excluded.vencimento,
                    numero_documento = excluded.numero_documento
                """,
                (
                    conta_id,
                    str(competencia),
                    str(d.caminho),
                    d.tipo.value,
                    d.origem.value,
                    d.confianca,
                    d.motivo,
                    d.paginas,
                    d.valor,
                    d.vencimento.isoformat() if d.vencimento else None,
                    d.numero_documento,
                    d.remetente,
                    d.assunto_email,
                ),
            )

        # O que não veio na lista sai. O DELETE é por `caminho` e fica na mesma
        # transação dos INSERTs: ou a competência inteira passa a refletir a
        # lista, ou nada muda.
        caminhos = [str(d.caminho) for d in documentos]
        marcadores = ",".join("?" * len(caminhos))
        con.execute(
            "DELETE FROM documento WHERE conta_id = ? AND competencia = ?"
            + (f" AND caminho NOT IN ({marcadores})" if caminhos else ""),
            (conta_id, str(competencia), *caminhos),
        )


def esquecer_documentos(conta_id: str, competencia: Competencia) -> None:
    with conexao() as con:
        con.execute(
            "DELETE FROM documento WHERE conta_id = ? AND competencia = ?",
            (conta_id, str(competencia)),
        )


def auditar(
    *,
    acao: str,
    destino: Path | str,
    conta_id: str | None = None,
    competencia: Competencia | None = None,
    origem: Path | str | None = None,
    backup: Path | str | None = None,
    simulado: bool = False,
    detalhes: dict | None = None,
) -> None:
    """Registra uma escrita (real ou simulada) fora da área local."""
    with conexao() as con:
        con.execute(
            """
            INSERT INTO auditoria (quando, conta_id, competencia, acao, destino,
                                   origem, backup, simulado, detalhes)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                _agora(),
                conta_id,
                str(competencia) if competencia else None,
                acao,
                str(destino),
                str(origem) if origem else None,
                str(backup) if backup else None,
                int(simulado),
                json.dumps(detalhes or {}, ensure_ascii=False, default=str),
            ),
        )
    log.info(
        "auditoria: %s%s -> %s", acao, " (simulado)" if simulado else "", destino
    )


def _garantir_processamento(
    con: sqlite3.Connection, conta_id: str, competencia: Competencia
) -> None:
    con.execute(
        """
        INSERT INTO processamento (conta_id, competencia, criado_em, atualizado_em)
        VALUES (?, ?, ?, ?)
        ON CONFLICT (conta_id, competencia) DO NOTHING
        """,
        (conta_id, str(competencia), _agora(), _agora()),
    )


# --------------------------------------------------------------------------- #
# Leitura
# --------------------------------------------------------------------------- #


def _data(texto: str | None):
    from datetime import date

    return date.fromisoformat(texto) if texto else None


def carregar(conta_id: str, competencia: Competencia) -> Processamento:
    """Reconstrói o estado de uma conta/competência. Nunca devolve None."""
    proc = Processamento(conta_id=conta_id, competencia=competencia)

    with conexao() as con:
        linha = con.execute(
            "SELECT * FROM processamento WHERE conta_id = ? AND competencia = ?",
            (conta_id, str(competencia)),
        ).fetchone()
        if linha:
            proc.valor = linha["valor"]
            proc.vencimento = _data(linha["vencimento"])
            proc.numero_documento = linha["numero_documento"]
            proc.descricao = linha["descricao"]
            proc.pagante = linha["pagante"]
            proc.beneficiario = linha["beneficiario"]
            proc.meio_pgto = linha["meio_pgto"]
            proc.forma_pgto = linha["forma_pgto"]
            proc.xlsx = Path(linha["xlsx"]) if linha["xlsx"] else None
            proc.pdf_final = Path(linha["pdf_final"]) if linha["pdf_final"] else None
            proc.email_enviado = bool(linha["email_enviado"])
            proc.enviado_em = _data(linha["enviado_em"])
            proc.enviado_por_fora = bool(linha["enviado_por_fora"])

        for e in con.execute(
            "SELECT * FROM etapa WHERE conta_id = ? AND competencia = ?",
            (conta_id, str(competencia)),
        ):
            try:
                etapa = Etapa(e["etapa"])
            except ValueError:
                continue  # etapa de versão antiga do código
            proc.resultados[etapa] = ResultadoEtapa(
                etapa=etapa,
                situacao=Situacao(e["situacao"]),
                mensagem=e["mensagem"],
                detalhes=json.loads(e["detalhes"]),
                artefatos=[Path(a) for a in json.loads(e["artefatos"])],
                atualizado_em=_momento(e["atualizado_em"]),
            )

        for d in con.execute(
            "SELECT * FROM documento WHERE conta_id = ? AND competencia = ? ORDER BY id",
            (conta_id, str(competencia)),
        ):
            proc.documentos.append(
                Documento(
                    caminho=Path(d["caminho"]),
                    tipo=TipoDocumento(d["tipo"]),
                    origem=OrigemDocumento(d["origem"]),
                    confianca=d["confianca"],
                    motivo=d["motivo"],
                    paginas=d["paginas"],
                    valor=d["valor"],
                    vencimento=_data(d["vencimento"]),
                    numero_documento=d["numero_documento"],
                    remetente=d["remetente"],
                    assunto_email=d["assunto_email"],
                )
            )

    return proc


def carregar_competencia(competencia: Competencia) -> dict[str, Processamento]:
    """Todos os processamentos de um mês, indexados por conta_id."""
    with conexao() as con:
        ids = [
            r["conta_id"]
            for r in con.execute(
                "SELECT DISTINCT conta_id FROM processamento WHERE competencia = ?",
                (str(competencia),),
            )
        ]
    return {i: carregar(i, competencia) for i in ids}


def valores_anteriores(
    competencia: Competencia,
) -> dict[str, tuple[float, Competencia]]:
    """
    O último valor de fatura já registrado ANTES desta competência, por conta.

    É daqui que sai o "previsto" do mês novo: o que se pagou da última vez é
    palpite melhor que um número fixo no cadastro, que envelhece sem ninguém
    perceber — e ninguém lembra de revisar 54 contas uma a uma.

    Só conta `valor` gravado de verdade (lido do boleto ou digitado no painel).
    Mês em que a fatura nunca foi lida tem `valor` NULL e fica de fora: senão o
    palpite de um mês viraria "histórico" no seguinte e se propagaria adiante
    como se fosse fato.

    Devolve também a competência de onde o valor veio, para a tela poder dizer
    de que mês está falando. Número sem procedência é o tipo de coisa que se
    confere uma vez e nunca mais.
    """
    with conexao() as con:
        linhas = con.execute(
            # MAX() com colunas soltas no SELECT: o SQLite garante que as
            # outras colunas venham da MESMA linha do máximo. Vale para
            # min/max e para mais nada — em qualquer outro agregado isto
            # seria um bug. A competência é gravada como 'AAAA-MM', então
            # comparar texto ordena por data.
            """
            SELECT conta_id, valor, MAX(competencia) AS competencia
              FROM processamento
             WHERE competencia < ? AND valor IS NOT NULL
             GROUP BY conta_id
            """,
            (str(competencia),),
        ).fetchall()

    return {
        linha["conta_id"]: (
            float(linha["valor"]),
            Competencia.de_texto(linha["competencia"]),
        )
        for linha in linhas
    }


def valor_anterior(
    conta_id: str, competencia: Competencia
) -> tuple[float, Competencia] | None:
    """O mesmo que `valores_anteriores`, para uma conta só."""
    with conexao() as con:
        linha = con.execute(
            """
            SELECT valor, MAX(competencia) AS competencia
              FROM processamento
             WHERE conta_id = ? AND competencia < ? AND valor IS NOT NULL
            """,
            (conta_id, str(competencia)),
        ).fetchone()

    # Sem nenhuma linha que sirva, o MAX() devolve uma linha só de nulos.
    if linha is None or linha["competencia"] is None:
        return None
    return float(linha["valor"]), Competencia.de_texto(linha["competencia"])


def ultimas_auditorias(limite: int = 100) -> list[dict]:
    with conexao() as con:
        return [
            dict(r)
            for r in con.execute(
                "SELECT * FROM auditoria ORDER BY id DESC LIMIT ?", (limite,)
            )
        ]


# --------------------------------------------------------------------------- #
# Contas ativas por mês
# --------------------------------------------------------------------------- #


def definir_contas_do_mes(competencia: Competencia, ativas: set[str]) -> None:
    """
    Grava quais contas valem nesta competência.

    Substitui o conjunto inteiro: o que não vier em `ativas` fica marcado como
    inativo naquele mês, sem sumir do registro nem afetar os outros meses.
    """
    with conexao() as con:
        con.execute(
            "DELETE FROM conta_do_mes WHERE competencia = ?", (str(competencia),)
        )
        con.executemany(
            "INSERT INTO conta_do_mes (competencia, conta_id, ativa, definido_em) "
            "VALUES (?, ?, 1, ?)",
            [(str(competencia), i, _agora()) for i in sorted(ativas)],
        )
    log.info("competência %s: %d conta(s) ativas", competencia, len(ativas))


def contas_do_mes(competencia: Competencia) -> set[str] | None:
    """
    Contas ativas desta competência, ou None se o mês nunca foi configurado.

    None significa "use o `ativo` do registro" — assim um mês novo já nasce
    com o padrão em vez de nascer vazio.
    """
    with conexao() as con:
        linhas = con.execute(
            "SELECT conta_id FROM conta_do_mes WHERE competencia = ? AND ativa = 1",
            (str(competencia),),
        ).fetchall()
    return {l["conta_id"] for l in linhas} if linhas else None


def competencias_da_conta(conta_id: str) -> list[Competencia]:
    """
    Os meses em que esta conta já teve alguma etapa gravada, do novo para o velho.

    Serve à limpeza da bancada: um arquivo parado em `entrada/<conta>/` pode
    ser de qualquer mês, e é aqui que se descobre quais meses vale a pena
    conferir no destino. Vem da tabela `etapa`, e não de `conta_do_mes`, para
    enxergar também o mês que foi processado sem estar na lista de ativas.
    """
    with conexao() as con:
        linhas = con.execute(
            "SELECT DISTINCT competencia FROM etapa WHERE conta_id = ? "
            "ORDER BY competencia DESC",
            (conta_id,),
        ).fetchall()

    meses: list[Competencia] = []
    for linha in linhas:
        try:
            meses.append(Competencia.de_texto(linha["competencia"]))
        except ValueError:
            continue
    return meses


def meses_configurados() -> list[str]:
    with conexao() as con:
        return [
            r["competencia"]
            for r in con.execute(
                "SELECT DISTINCT competencia FROM conta_do_mes ORDER BY competencia DESC"
            )
        ]


def copiar_contas_do_mes(origem: Competencia, destino: Competencia) -> int:
    """Repete em `destino` o conjunto de contas de `origem`."""
    ativas = contas_do_mes(origem)
    if not ativas:
        return 0
    definir_contas_do_mes(destino, ativas)
    return len(ativas)


def reabrir(conta_id: str, competencia: Competencia, etapa: Etapa | None = None) -> None:
    """
    Apaga o resultado de uma etapa (ou de todas) para reprocessar.

    Reabrir a etapa de e-mail também limpa o registro de despacho. É de
    propósito: a trava contra envio repetido tem que ceder a um ato
    deliberado — reenviar acontece (anexo errado, valor corrigido) —, mas não
    a um duplo clique nem a um F5.
    """
    limpa_envio = etapa in (None, Etapa.EMAIL)
    with conexao() as con:
        if etapa:
            con.execute(
                "DELETE FROM etapa WHERE conta_id = ? AND competencia = ? AND etapa = ?",
                (conta_id, str(competencia), etapa.value),
            )
        else:
            con.execute(
                "DELETE FROM etapa WHERE conta_id = ? AND competencia = ?",
                (conta_id, str(competencia)),
            )
        if limpa_envio:
            # Só o envio feito pela ferramenta. O registro "enviei por fora"
            # tem botão próprio para desfazer e não some por reprocessamento.
            con.execute(
                "UPDATE processamento SET email_enviado = 0, enviado_em = NULL, "
                "atualizado_em = ? WHERE conta_id = ? AND competencia = ? "
                "AND enviado_por_fora = 0",
                (_agora(), conta_id, str(competencia)),
            )


if __name__ == "__main__":  # teste manual
    from automacao import configurar_log

    configurar_log()
    comp = Competencia(2026, 8)
    salvar_etapa(
        "teste-conta",
        comp,
        ResultadoEtapa.sucesso(Etapa.COLETA, "2 documentos encontrados"),
    )
    p = carregar("teste-conta", comp)
    print(f"situação da coleta: {p.situacao_de(Etapa.COLETA).value}")
    print(f"concluído? {p.concluido}")
    with conexao() as c:
        c.execute("DELETE FROM etapa WHERE conta_id = 'teste-conta'")
        c.execute("DELETE FROM processamento WHERE conta_id = 'teste-conta'")
    print("limpo.")
