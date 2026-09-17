"""
Exportação da aba de autorização para PDF, via Excel (COM).

Por que Excel e não uma biblioteca: o template usa fórmulas (`VLOOKUP`,
`TODAY`, `SUM`) e uma área de impressão calibrada à mão (`$A$2:$J$48` ou
`$A$2:$J$49`). Só o próprio Excel devolve o PDF idêntico ao que o financeiro
recebe hoje.

Cuidados adotados:
  * `DispatchEx` — instância dedicada, não sequestra o Excel aberto do usuário;
  * `CoInitialize`/`CoUninitialize` — funciona dentro de thread do servidor web;
  * `try/finally` fechando workbook e aplicação em qualquer cenário;
  * nada de mexer em configuração de página: a área de impressão do arquivo é
    respeitada como está.
"""

from __future__ import annotations

import logging
import os
from pathlib import Path

from automacao.nucleo.config import ambiente
from automacao.nucleo.modelos import Etapa, ResultadoEtapa

registrador = logging.getLogger("automacao.documentos.exportar_pdf")

#: xlTypePDF
_XL_TIPO_PDF = 0
#: xlCalculationAutomatic
_XL_CALCULO_AUTOMATICO = -4105

#: HRESULTs que aparecem com frequência, traduzidos para algo acionável.
_ERROS_CONHECIDOS: dict[int, str] = {
    -2147221005: (
        "o Excel não está instalado ou não está registrado no Windows "
        "(erro 'Invalid class string'). Instale o Microsoft Excel nesta máquina "
        "para exportar o PDF."
    ),
    -2147221164: (
        "o componente COM do Excel não está registrado. Reinstale o Office ou "
        "rode 'excel.exe /regserver' em um prompt como administrador."
    ),
    -2146959355: (
        "o Windows não conseguiu iniciar o Excel (Server execution failed). "
        "Feche instâncias travadas do EXCEL.EXE no Gerenciador de Tarefas e "
        "tente de novo."
    ),
    -2147417846: (
        "o Excel está ocupado respondendo a outra chamada (provavelmente com uma "
        "caixa de diálogo aberta). Feche os diálogos do Excel e tente de novo."
    ),
}


class ErroExportacao(RuntimeError):
    """Falha ao exportar a autorização para PDF."""


def _importar_com():
    """Importa pywin32 com mensagem clara caso não esteja disponível."""
    try:
        import pythoncom  # type: ignore[import-not-found]
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as erro:  # pragma: no cover — depende do ambiente
        raise ErroExportacao(
            "pywin32 não está instalado neste ambiente — sem ele não dá para "
            "conversar com o Excel. Instale com 'pip install pywin32'."
        ) from erro
    return pythoncom, win32com.client


def _traduzir(erro: Exception, caminho: Path | None = None) -> str:
    """Transforma um erro COM em uma frase que o usuário consegue agir."""
    codigo = getattr(erro, "hresult", None)
    if codigo is None:
        args = getattr(erro, "args", ())
        codigo = args[0] if args and isinstance(args[0], int) else None

    if codigo in _ERROS_CONHECIDOS:
        return _ERROS_CONHECIDOS[codigo]

    # A descrição útil do Excel vem no excepinfo (3º elemento de args).
    detalhe = ""
    args = getattr(erro, "args", ())
    if len(args) >= 3 and isinstance(args[2], tuple):
        partes = [str(p) for p in args[2] if isinstance(p, str) and p.strip()]
        detalhe = " ".join(partes).strip()
    detalhe = detalhe or str(erro)
    baixo = detalhe.lower()

    if "senha" in baixo or "password" in baixo:
        return (
            f"a planilha está protegida por senha e o Excel não conseguiu abri-la "
            f"sem intervenção: {detalhe}"
        )
    if "protegid" in baixo or "protected" in baixo or "read-only" in baixo:
        return (
            f"a planilha ou a aba está protegida: {detalhe}. Desproteja a aba no "
            f"Excel (Revisão > Desproteger Planilha) e tente de novo."
        )
    if "sendo usado" in baixo or "in use" in baixo or "being used" in baixo:
        alvo = f" ({caminho.name})" if caminho else ""
        return (
            f"o arquivo{alvo} está aberto ou bloqueado por outro processo. "
            f"Feche o arquivo no Excel e tente de novo."
        )
    return f"o Excel recusou a operação: {detalhe}"


def excel_disponivel() -> tuple[bool, str]:
    """
    Checagem rápida do ambiente, para o painel avisar antes de rodar.

    Devolve `(disponivel, mensagem)`.
    """
    try:
        pythoncom, cliente = _importar_com()
    except ErroExportacao as erro:
        return False, str(erro)

    pythoncom.CoInitialize()
    excel = None
    try:
        excel = cliente.DispatchEx("Excel.Application")
        versao = str(excel.Version)
        return True, f"Excel {versao} disponível para exportação."
    except Exception as erro:  # noqa: BLE001 — qualquer falha aqui é indisponibilidade
        return False, _traduzir(erro)
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:  # noqa: BLE001
                registrador.debug("falha ao encerrar o Excel da checagem", exc_info=True)
        pythoncom.CoUninitialize()


def exportar(
    caminho_xlsx: Path | str,
    aba: str,
    destino_pdf: Path | str | None = None,
) -> ResultadoEtapa:
    """
    Exporta **apenas** a aba `aba` de `caminho_xlsx` para PDF.

    A área de impressão já definida no arquivo é respeitada; nenhuma
    configuração de página é alterada.

    Sem `destino_pdf`, o PDF nasce ao lado do xlsx como
    `"<nome> - autorizacao.pdf"` — de propósito **diferente** do nome do xlsx,
    porque o nome limpo é reservado ao PDF final montado por `montar_pdf`
    (que leria e sobrescreveria este arquivo ao mesmo tempo).
    """
    origem = Path(caminho_xlsx).resolve()
    if not origem.is_file():
        mensagem = f"planilha não encontrada: {origem}"
        registrador.error(mensagem)
        return ResultadoEtapa.erro(Etapa.PDF_AUTORIZACAO, mensagem)

    destino = (
        Path(destino_pdf)
        if destino_pdf
        else origem.with_name(f"{origem.stem} - autorizacao.pdf")
    )
    destino = destino.resolve() if destino.parent.exists() else destino
    if destino.suffix.lower() != ".pdf":
        mensagem = f"destino do PDF precisa terminar em .pdf: {destino}"
        registrador.error(mensagem)
        return ResultadoEtapa.erro(Etapa.PDF_AUTORIZACAO, mensagem)

    # Trava de segurança: nada é escrito fora de dados/ sem permissão.
    try:
        ambiente().exigir_permissao(destino)
    except PermissionError as erro:
        registrador.error("%s", erro)
        return ResultadoEtapa.erro(Etapa.PDF_AUTORIZACAO, str(erro))

    destino.parent.mkdir(parents=True, exist_ok=True)
    if destino.exists():
        # Segunda trava, independente da de escrita: apagar só dentro de dados/.
        try:
            ambiente().exigir_permissao_para_apagar(destino)
        except PermissionError as erro:
            registrador.error("%s", erro)
            return ResultadoEtapa.erro(Etapa.PDF_AUTORIZACAO, str(erro))
        try:
            destino.unlink()
        except OSError as erro:
            mensagem = (
                f"não foi possível substituir {destino.name}: {erro}. "
                f"O PDF pode estar aberto em um leitor."
            )
            registrador.error(mensagem)
            return ResultadoEtapa.erro(Etapa.PDF_AUTORIZACAO, mensagem)

    try:
        pythoncom, cliente = _importar_com()
    except ErroExportacao as erro:
        registrador.error("%s", erro)
        return ResultadoEtapa.erro(Etapa.PDF_AUTORIZACAO, str(erro))

    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        try:
            excel = cliente.DispatchEx("Excel.Application")
        except Exception as erro:  # noqa: BLE001
            mensagem = _traduzir(erro, origem)
            registrador.error("não foi possível iniciar o Excel: %s", mensagem)
            return ResultadoEtapa.erro(Etapa.PDF_AUTORIZACAO, mensagem)

        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False

        try:
            # ATENÇÃO: em late binding o pywin32 IGNORA argumentos nomeados
            # sem avisar — `ReadOnly=True` viraria silenciosamente o padrão.
            # Toda chamada COM neste projeto é posicional.
            # Open(Filename, UpdateLinks, ReadOnly)
            wb = excel.Workbooks.Open(str(origem), 0, True)
        except Exception as erro:  # noqa: BLE001
            mensagem = _traduzir(erro, origem)
            registrador.error("não foi possível abrir %s: %s", origem.name, mensagem)
            return ResultadoEtapa.erro(Etapa.PDF_AUTORIZACAO, mensagem)

        abas = [str(w.Name) for w in wb.Worksheets]
        if aba not in abas:
            mensagem = (
                f"a aba {aba!r} não existe em {origem.name}. Abas disponíveis: {abas}."
            )
            registrador.error(mensagem)
            return ResultadoEtapa.erro(Etapa.PDF_AUTORIZACAO, mensagem)

        # Garante que VLOOKUP/TODAY/SUM estejam resolvidos antes de imprimir.
        # (Application.Calculation só aceita atribuição com um workbook aberto.)
        try:
            excel.Calculation = _XL_CALCULO_AUTOMATICO
            excel.CalculateFullRebuild()
        except Exception:  # noqa: BLE001 — recálculo é reforço, não requisito
            registrador.debug("recálculo forçado indisponível", exc_info=True)

        ws = wb.Worksheets(aba)
        area = str(ws.PageSetup.PrintArea or "")

        try:
            ws.ExportAsFixedFormat(_XL_TIPO_PDF, str(destino))
        except Exception as erro:  # noqa: BLE001
            mensagem = _traduzir(erro, destino)
            registrador.error("falha ao exportar %s: %s", origem.name, mensagem)
            return ResultadoEtapa.erro(Etapa.PDF_AUTORIZACAO, mensagem)

    finally:
        if wb is not None:
            try:
                wb.Close(False)  # posicional: SaveChanges
            except Exception:  # noqa: BLE001
                registrador.warning("não foi possível fechar o workbook", exc_info=True)
        if excel is not None:
            try:
                excel.ScreenUpdating = True
                excel.DisplayAlerts = True
            except Exception:  # noqa: BLE001
                registrador.debug("não foi possível restaurar flags do Excel", exc_info=True)
            try:
                excel.Quit()
            except Exception:  # noqa: BLE001
                registrador.warning("não foi possível encerrar o Excel", exc_info=True)
        pythoncom.CoUninitialize()

    if not destino.is_file() or destino.stat().st_size == 0:
        mensagem = (
            f"o Excel não reclamou, mas {destino.name} não foi gerado (ou saiu vazio). "
            f"Confira se a aba {aba!r} tem conteúdo e área de impressão."
        )
        registrador.error(mensagem)
        return ResultadoEtapa.erro(Etapa.PDF_AUTORIZACAO, mensagem)

    detalhes = {
        "origem": str(origem),
        "aba": aba,
        "area_impressao": area,
        "bytes": destino.stat().st_size,
    }
    registrador.info(
        "%s: aba %r exportada para %s (%d bytes, área %s)",
        origem.name, aba, destino.name, detalhes["bytes"], area or "padrão",
    )
    return ResultadoEtapa.sucesso(
        Etapa.PDF_AUTORIZACAO,
        f"{destino.name} gerado a partir da aba {aba!r}.",
        detalhes=detalhes,
        artefatos=[destino],
    )


# --------------------------------------------------------------------------- #
# Teste manual
# --------------------------------------------------------------------------- #

if __name__ == "__main__":  # pragma: no cover
    import sys

    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")

    disponivel, recado = excel_disponivel()
    registrador.info("Excel disponível: %s — %s", disponivel, recado)

    if len(sys.argv) > 1:
        planilha = Path(sys.argv[1])
        nome_aba = sys.argv[2] if len(sys.argv) > 2 else "AUTORIZAÇÃO"
        saida = exportar(planilha, nome_aba)
        registrador.info("situação: %s — %s", saida.situacao.value, saida.mensagem)
    else:
        registrador.info(
            "uso: python -m automacao.exportar_pdf <caminho.xlsx> [aba]  "
            "(diretório atual: %s)", os.getcwd(),
        )
