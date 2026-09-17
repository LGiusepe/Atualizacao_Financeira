"""
Montagem do PDF único que vai anexado ao e-mail.

Ordem obrigatória (contrato do projeto):
    1º autorização → 2º demonstrativo (se houver) → 3º boleto → 4º nota fiscal.

A ordem efetiva vem de `ambiente().pdf["ordem"]`, mas o contrato acima é o
padrão e o que vale quando a configuração estiver ausente ou incompleta.

O PDF final recebe o mesmo nome do xlsx da autorização — regra do usuário,
para não duplicar arquivo no servidor. Quem resolve o nome é o chamador;
aqui só validamos que a extensão é `.pdf`.
"""

from __future__ import annotations

import logging
from pathlib import Path

from pypdf import PdfReader, PdfWriter
from pypdf.errors import PyPdfError

from automacao.nucleo.config import ambiente
from automacao.nucleo.modelos import Documento, Etapa, ResultadoEtapa, Situacao, TipoDocumento

registrador = logging.getLogger("automacao.documentos.montar_pdf")

#: Contrato do projeto. `ambiente().pdf["ordem"]` pode reordenar, nunca
#: inventar tipos novos.
ORDEM_PADRAO: tuple[TipoDocumento, ...] = (
    TipoDocumento.AUTORIZACAO,
    TipoDocumento.DEMONSTRATIVO,
    TipoDocumento.BOLETO,
    TipoDocumento.NOTA_FISCAL,
)


def ordem_configurada(ambiente_=None) -> tuple[TipoDocumento, ...]:
    """Lê a ordem de montagem do settings.yaml, caindo no padrão do contrato."""
    try:
        amb = ambiente_ or ambiente()
        bruto = list((amb.pdf or {}).get("ordem") or [])
    except Exception as erro:  # noqa: BLE001 — sem config, vale o contrato
        registrador.debug("ordem do settings.yaml indisponível (%s)", erro)
        return ORDEM_PADRAO

    ordem: list[TipoDocumento] = []
    for item in bruto:
        try:
            tipo = TipoDocumento(str(item).strip().lower())
        except ValueError:
            registrador.warning("tipo desconhecido em pdf.ordem: %r — ignorado", item)
            continue
        if tipo not in ordem:
            ordem.append(tipo)

    if TipoDocumento.AUTORIZACAO not in ordem:
        ordem.insert(0, TipoDocumento.AUTORIZACAO)
    return tuple(ordem) or ORDEM_PADRAO


def _abrir(caminho: Path, conta_id: str | None = None) -> PdfReader:
    """
    Abre o PDF, destravando o que der.

    Tenta senha vazia (resolve o boleto com restrição de impressão) e depois a
    senha guardada da conta. Na base atual, 59 dos 69 PDFs criptografados são
    faturas de operadora com senha de verdade — sem ela, o PDF final sairia
    incompleto e sem ninguém perceber.
    """
    from automacao.acesso.senhas_pdf import abrir as abrir_protegido

    return abrir_protegido(caminho, conta_id)


def montar(
    documentos: list[Documento],
    destino: Path,
    conta_id: str | None = None,
) -> ResultadoEtapa:
    """
    Junta os PDFs de `documentos` em `destino`, na ordem do contrato.

    - falta a autorização → `Situacao.ERRO`;
    - falta o boleto → `Situacao.ATENCAO` (algumas contas só têm fatura),
      mas o PDF é montado assim mesmo;
    - documentos `DESCONHECIDO` (e qualquer tipo fora da ordem, como o XML da
      NF-e) não entram; ficam listados em `detalhes["ignorados"]`;
    - PDF corrompido/protegido não derruba a montagem: vai para
      `detalhes["falhas"]` e o resto segue.

    `detalhes["indice"]` traz, para o painel:
    `[{"tipo", "arquivo", "pagina_inicial", "paginas"}]`.
    """
    alvo = Path(destino)
    if alvo.suffix.lower() != ".pdf":
        mensagem = (
            f"o destino do PDF final precisa terminar em .pdf (recebido: {alvo.name})."
        )
        registrador.error(mensagem)
        return ResultadoEtapa.erro(Etapa.PDF_FINAL, mensagem)

    ordem = ordem_configurada()
    posicao = {tipo: i for i, tipo in enumerate(ordem)}

    entram: list[Documento] = []
    ignorados: list[dict] = []
    for doc in documentos:
        if doc.tipo in posicao:
            entram.append(doc)
        else:
            ignorados.append(
                {
                    "arquivo": str(doc.caminho),
                    "tipo": doc.tipo.value,
                    "motivo": (
                        "documento não identificado — confirme o tipo no painel"
                        if doc.tipo is TipoDocumento.DESCONHECIDO
                        else f"tipo {doc.tipo.value!r} não entra no PDF final"
                    ),
                }
            )

    # Ordenação estável: mantém a ordem original dentro de cada tipo.
    entram.sort(key=lambda d: posicao[d.tipo])

    presentes = {d.tipo for d in entram}
    if TipoDocumento.AUTORIZACAO not in presentes:
        mensagem = (
            "não há autorização de pagamento entre os documentos — o PDF final "
            "não pode ser montado sem ela."
        )
        registrador.error("%s (destino %s)", mensagem, alvo.name)
        return ResultadoEtapa.erro(
            Etapa.PDF_FINAL, mensagem, detalhes={"ignorados": ignorados}
        )

    # O destino não pode ser uma das entradas: seria ler e sobrescrever ao
    # mesmo tempo (acontece quando o PDF da autorização recebe o nome final).
    resolvido = alvo.resolve()
    conflito = next(
        (d for d in entram if Path(d.caminho).resolve() == resolvido), None
    )
    if conflito is not None:
        mensagem = (
            f"o destino {alvo.name} é um dos PDFs de entrada "
            f"({conflito.tipo.value}) — escolha outro nome para o PDF final."
        )
        registrador.error(mensagem)
        return ResultadoEtapa.erro(Etapa.PDF_FINAL, mensagem)

    # Trava de segurança: nada é escrito fora de dados/ sem permissão.
    try:
        ambiente().exigir_permissao(alvo)
    except PermissionError as erro:
        registrador.error("%s", erro)
        return ResultadoEtapa.erro(Etapa.PDF_FINAL, str(erro))

    escritor = PdfWriter()
    indice: list[dict] = []
    falhas: list[dict] = []
    proxima_pagina = 1

    for doc in entram:
        caminho = Path(doc.caminho)
        if not caminho.is_file():
            falhas.append(
                {
                    "arquivo": str(caminho),
                    "tipo": doc.tipo.value,
                    "erro": "arquivo não encontrado",
                }
            )
            registrador.warning("%s: arquivo não encontrado", caminho)
            continue
        try:
            leitor = _abrir(caminho, conta_id)
            paginas = len(leitor.pages)
            for pagina in leitor.pages:
                escritor.add_page(pagina)
        except (PyPdfError, OSError, ValueError, KeyError) as erro:
            falhas.append(
                {
                    "arquivo": str(caminho),
                    "tipo": doc.tipo.value,
                    "erro": f"{type(erro).__name__}: {erro}",
                }
            )
            registrador.warning("%s: PDF ilegível (%s)", caminho.name, erro)
            continue

        doc.paginas = paginas
        indice.append(
            {
                "tipo": doc.tipo.value,
                "arquivo": str(caminho),
                "pagina_inicial": proxima_pagina,
                "paginas": paginas,
            }
        )
        proxima_pagina += paginas

    if not indice:
        mensagem = "nenhum PDF pôde ser lido — o arquivo final não foi gerado."
        registrador.error(mensagem)
        return ResultadoEtapa.erro(
            Etapa.PDF_FINAL,
            mensagem,
            detalhes={"ignorados": ignorados, "falhas": falhas, "indice": []},
        )

    montados = {item["tipo"] for item in indice}
    if TipoDocumento.AUTORIZACAO.value not in montados:
        mensagem = (
            "a autorização de pagamento existe na lista mas não pôde ser lida — "
            "o PDF final não foi gerado."
        )
        registrador.error(mensagem)
        return ResultadoEtapa.erro(
            Etapa.PDF_FINAL,
            mensagem,
            detalhes={"ignorados": ignorados, "falhas": falhas, "indice": indice},
        )

    alvo.parent.mkdir(parents=True, exist_ok=True)
    try:
        with alvo.open("wb") as saida:
            escritor.write(saida)
    except OSError as erro:
        mensagem = (
            f"não foi possível gravar {alvo.name}: {erro}. "
            f"O PDF pode estar aberto em um leitor."
        )
        registrador.error(mensagem)
        return ResultadoEtapa.erro(
            Etapa.PDF_FINAL,
            mensagem,
            detalhes={"ignorados": ignorados, "falhas": falhas, "indice": indice},
        )
    finally:
        escritor.close()

    total = proxima_pagina - 1
    detalhes = {
        "indice": indice,
        "ignorados": ignorados,
        "falhas": falhas,
        "ordem": [t.value for t in ordem],
        "total_paginas": total,
        "arquivo": str(alvo),
    }

    avisos: list[str] = []
    if TipoDocumento.BOLETO.value not in montados:
        avisos.append("sem boleto — conta paga só com fatura/nota? Confira.")
    if TipoDocumento.NOTA_FISCAL.value not in montados:
        avisos.append("sem nota fiscal no anexo.")
    if falhas:
        avisos.append(f"{len(falhas)} arquivo(s) não puderam ser lidos.")
    if ignorados:
        avisos.append(f"{len(ignorados)} arquivo(s) ignorados por tipo.")
    detalhes["avisos"] = avisos

    resumo = f"{alvo.name}: {total} página(s) de {len(indice)} documento(s)"
    registrador.info("%s — ordem: %s", resumo, " > ".join(i["tipo"] for i in indice))

    sem_boleto = TipoDocumento.BOLETO.value not in montados
    if sem_boleto or falhas:
        motivo = "; ".join(avisos)
        return ResultadoEtapa(
            Etapa.PDF_FINAL,
            Situacao.ATENCAO,
            f"{resumo}. {motivo}",
            detalhes=detalhes,
            artefatos=[alvo],
        )
    return ResultadoEtapa.sucesso(
        Etapa.PDF_FINAL, resumo, detalhes=detalhes, artefatos=[alvo]
    )


# --------------------------------------------------------------------------- #
# Teste manual
# --------------------------------------------------------------------------- #

if __name__ == "__main__":  # pragma: no cover
    import sys

    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")

    if len(sys.argv) < 3:
        registrador.info(
            "uso: python -m automacao.montar_pdf <saida.pdf> "
            "<autorizacao.pdf> [demonstrativo.pdf] [boleto.pdf] [nf.pdf]"
        )
        raise SystemExit(0)

    saida_pdf = Path(sys.argv[1])
    tipos = [
        TipoDocumento.AUTORIZACAO,
        TipoDocumento.DEMONSTRATIVO,
        TipoDocumento.BOLETO,
        TipoDocumento.NOTA_FISCAL,
    ]
    entradas = [
        Documento(caminho=Path(arg), tipo=tipos[min(i, len(tipos) - 1)])
        for i, arg in enumerate(sys.argv[2:])
    ]

    resultado = montar(entradas, saida_pdf)
    registrador.info("situação: %s — %s", resultado.situacao.value, resultado.mensagem)
    for item in resultado.detalhes.get("indice", []):
        registrador.info(
            "  p.%s (+%s) %s — %s",
            item["pagina_inicial"], item["paginas"], item["tipo"],
            Path(item["arquivo"]).name,
        )
