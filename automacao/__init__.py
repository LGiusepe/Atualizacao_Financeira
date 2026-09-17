"""
Automação Financeira — Autorizações de Pagamento (T.I.)

Fluxo automatizado:
    coletar fatura → organizar pasta do mês → preencher autorização →
    exportar PDF → montar PDF único → publicar no OneDrive →
    criar rascunho do e-mail → marcar verde no checklist

Princípio de operação: a automação PREPARA tudo na bancada de trabalho. Só
grava no OneDrive e no Outlook depois de confirmação explícita.

Como os módulos estão arrumados — as três primeiras pastas são o próprio fluxo,
na ordem em que ele acontece:

    nucleo/      modelos, config, estado — o que todo o resto usa
    coleta/      de onde a fatura vem e que documento ela é
    documentos/  a autorização preenchida e os PDFs
    entrega/     OneDrive, Outlook e checklist: o mundo de fora
    acesso/      quem entra no painel, até onde vai, e as senhas guardadas
    orquestrador.py

O `orquestrador` fica na raiz porque não pertence a nenhuma das etapas: ele é
quem as chama, na ordem, e decide o que já pode rodar. Cada pacote tem um
`__init__.py` explicando o que cabe ali — vale a leitura antes de acrescentar
um módulo, para ele não nascer no lugar errado.

Os módulos pesados (Excel, Outlook, PDF) são importados dentro das funções, não
no topo: é o que deixa o painel abrir numa máquina sem Office para conferir
configuração.
"""

from __future__ import annotations

import logging

__version__ = "0.1.0"

logging.getLogger(__name__).addHandler(logging.NullHandler())


def configurar_log(nivel: int = logging.INFO, arquivo: str | None = None) -> None:
    """Liga o log no console (e opcionalmente em arquivo). Chame só na ponta."""
    import sys
    from pathlib import Path

    # Sob `pythonw.exe` — que é como o painel sobe junto com o Windows — não
    # existe console: `sys.stderr` é None e o StreamHandler quebraria a cada
    # linha de log. O arquivo continua sendo escrito normalmente.
    handlers: list[logging.Handler] = []
    if sys.stderr is not None:
        handlers.append(logging.StreamHandler())
    if arquivo:
        caminho = Path(arquivo)
        caminho.parent.mkdir(parents=True, exist_ok=True)
        handlers.append(logging.FileHandler(caminho, encoding="utf-8"))

    logging.basicConfig(
        level=nivel,
        format="%(asctime)s  %(levelname)-8s %(name)-33s %(message)s",
        datefmt="%H:%M:%S",
        handlers=handlers,
        force=True,
    )
