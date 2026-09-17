"""
Senhas de PDF por conta.

Boa parte das faturas de operadora vem com senha: na base atual, 59 dos 69 PDFs
criptografados não abrem com senha vazia — as operadoras, principalmente. Sem a
senha, o pypdf não lê as páginas e o PDF final sai incompleto.

A senha é sempre a mesma para o mesmo contrato (CPF do titular, CNPJ, número da
conta), então vale guardar uma vez por conta. Fica no cofre DPAPI da máquina,
nunca em texto puro e nunca no OneDrive.
"""

from __future__ import annotations

import logging
from pathlib import Path

from pypdf import PdfReader

from automacao.acesso import cofre

log = logging.getLogger("automacao.acesso.senhas_pdf")

PREFIXO_COFRE = "pdf::"


def _chave(conta_id: str) -> str:
    return f"{PREFIXO_COFRE}{conta_id}"


def guardar(conta_id: str, senha: str) -> None:
    """Guarda a senha do PDF desta conta no cofre cifrado."""
    cofre.gravar(
        _chave(conta_id),
        login=conta_id,
        senha=senha,
        observacao="senha usada para abrir o PDF da fatura",
    )
    log.info("senha de PDF guardada para %s", conta_id)


def obter(conta_id: str) -> str | None:
    credencial = cofre.obter(_chave(conta_id))
    return credencial.senha if credencial else None


def remover(conta_id: str) -> bool:
    return cofre.remover(_chave(conta_id))


def contas_com_senha() -> list[str]:
    return [
        nome[len(PREFIXO_COFRE):]
        for nome in cofre.listar()
        if nome.startswith(PREFIXO_COFRE)
    ]


def esta_travado(caminho: Path) -> bool:
    """True quando o PDF é criptografado e não abre com senha vazia."""
    try:
        leitor = PdfReader(str(caminho))
    except Exception:
        return False
    if not leitor.is_encrypted:
        return False
    try:
        if leitor.decrypt(""):
            _ = len(leitor.pages)
            return False
    except Exception:
        pass
    return True


def travados(caminhos: list[Path]) -> list[Path]:
    """Quais destes PDFs precisam de senha."""
    return [c for c in caminhos if c.suffix.lower() == ".pdf" and esta_travado(c)]


def abrir(caminho: Path, conta_id: str | None = None) -> PdfReader:
    """
    Abre o PDF, tentando senha vazia e depois a senha guardada da conta.

    Levanta `PermissionError` com mensagem acionável quando nenhuma serve —
    aí o painel pede a senha para o usuário.
    """
    leitor = PdfReader(str(caminho))
    if not leitor.is_encrypted:
        return leitor

    for senha, origem in (("", "senha vazia"), (obter(conta_id) if conta_id else None, "senha guardada")):
        if senha is None:
            continue
        try:
            if leitor.decrypt(senha):
                _ = len(leitor.pages)
                log.debug("%s aberto com %s", caminho.name, origem)
                return leitor
        except Exception:
            continue
        # decrypt pode devolver 0 sem estourar; recomeça limpo para a próxima
        leitor = PdfReader(str(caminho))

    raise PermissionError(
        f"{caminho.name} está protegido por senha e nenhuma das senhas "
        "conhecidas abriu. Informe a senha na etapa de coleta do painel — "
        "ela fica guardada cifrada e vale para os próximos meses."
    )


def testar(caminho: Path, senha: str) -> tuple[bool, str]:
    """Confere se a senha abre o PDF, sem guardar nada."""
    try:
        leitor = PdfReader(str(caminho))
    except Exception as erro:
        return False, f"não consegui abrir o arquivo: {erro}"

    if not leitor.is_encrypted:
        return True, "este PDF não tem senha."
    try:
        if leitor.decrypt(senha):
            return True, f"senha confere — {len(leitor.pages)} página(s) legíveis."
    except Exception as erro:
        return False, f"a senha não foi aceita: {erro}"
    return False, "a senha não foi aceita."


if __name__ == "__main__":  # teste manual
    import sys

    from automacao import configurar_log

    configurar_log()
    if len(sys.argv) > 1:
        alvo = Path(sys.argv[1])
        print(f"{alvo.name}: travado={esta_travado(alvo)}")
    print("contas com senha guardada:", contas_com_senha() or "nenhuma")
