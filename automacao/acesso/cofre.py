"""
Cofre local de credenciais, protegido pelo Windows DPAPI.

O arquivo `dados/cofre.dat` guarda um JSON cifrado com
`win32crypt.CryptProtectData`, no escopo do usuário logado e com entropia
adicional fixa deste projeto. Na prática:

  - só a conta Windows que gravou consegue abrir;
  - copiar o .dat para outra máquina ou outro usuário não adianta nada;
  - mesmo dentro da máquina, sem a entropia do projeto o arquivo não abre.

O cofre NUNCA vai para o OneDrive nem para o git (ver `.gitignore`).

Importante: `importar_da_planilha()` só LÊ a aba ACESSOS E CONTATOS. A
planilha continua sendo a fonte da verdade das senhas — decisão do usuário.
O cofre é conveniência local, não substituto.
"""

from __future__ import annotations

import json
import logging
import os
import tempfile
from pathlib import Path

from automacao.nucleo.config import RAIZ_PROJETO
from automacao.entrega.planilha_contas import Credencial

log = logging.getLogger("automacao.acesso.cofre")

# Onde o cofre mora. Fora do OneDrive, de propósito.
CAMINHO_COFRE = RAIZ_PROJETO / "dados" / "cofre.dat"

# Entropia adicional do projeto. Não é segredo (está no código-fonte), mas
# amarra o blob a esta aplicação: outro programa rodando com o mesmo usuário
# não decifra o arquivo por acidente.
ENTROPIA = b"AutomacaoFinanceira::cofre::v1"

DESCRICAO = "Automação Financeira — cofre de credenciais"

VERSAO_FORMATO = 1


class ErroCofre(RuntimeError):
    """Falha ao abrir, decifrar ou gravar o cofre."""


# --------------------------------------------------------------------------- #
# DPAPI
# --------------------------------------------------------------------------- #


def _win32crypt():
    try:
        import win32crypt  # noqa: PLC0415 - import tardio, só quando precisa
    except ImportError as exc:  # pragma: no cover
        raise ErroCofre(
            f"pywin32 (win32crypt) não disponível — o cofre exige Windows DPAPI: {exc}"
        ) from exc
    return win32crypt


def _cifrar(dados: dict) -> bytes:
    bruto = json.dumps(dados, ensure_ascii=False, sort_keys=True).encode("utf-8")
    return _win32crypt().CryptProtectData(bruto, DESCRICAO, ENTROPIA, None, None, 0)


def _decifrar(blob: bytes) -> dict:
    try:
        _descricao, bruto = _win32crypt().CryptUnprotectData(
            blob, ENTROPIA, None, None, 0
        )
    except Exception as exc:
        raise ErroCofre(
            f"não foi possível abrir {CAMINHO_COFRE.name} — o arquivo pertence a "
            f"outro usuário/máquina ou está corrompido ({type(exc).__name__}: {exc})"
        ) from exc
    try:
        dados = json.loads(bruto.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ErroCofre(f"conteúdo do cofre ilegível: {exc}") from exc
    if not isinstance(dados, dict):
        raise ErroCofre("conteúdo do cofre não é um mapeamento")
    return dados


# --------------------------------------------------------------------------- #
# Persistência
# --------------------------------------------------------------------------- #


def _carregar() -> dict[str, dict]:
    """Devolve {chave_normalizada: {servico, login, senha, site, observacao}}."""
    if not CAMINHO_COFRE.is_file():
        return {}
    conteudo = _decifrar(CAMINHO_COFRE.read_bytes())
    if conteudo.get("versao") != VERSAO_FORMATO:
        log.warning(
            "cofre na versão %r (esperada %r) — lendo assim mesmo",
            conteudo.get("versao"),
            VERSAO_FORMATO,
        )
    itens = conteudo.get("credenciais") or {}
    return itens if isinstance(itens, dict) else {}


def _salvar(itens: dict[str, dict]) -> None:
    """Grava de forma atômica: escreve num temporário e substitui."""
    CAMINHO_COFRE.parent.mkdir(parents=True, exist_ok=True)
    blob = _cifrar({"versao": VERSAO_FORMATO, "credenciais": itens})

    descritor, temporario = tempfile.mkstemp(
        dir=str(CAMINHO_COFRE.parent), prefix=".cofre-", suffix=".tmp"
    )
    try:
        with os.fdopen(descritor, "wb") as fh:
            fh.write(blob)
            fh.flush()
            os.fsync(fh.fileno())
        os.replace(temporario, CAMINHO_COFRE)
    except BaseException:
        Path(temporario).unlink(missing_ok=True)
        raise
    log.info("cofre gravado (%d credenciais, %d bytes)", len(itens), len(blob))


def _chave(servico: str) -> str:
    """Chave de busca: sem acento, minúscula, espaços colapsados."""
    from automacao.nucleo.modelos import normalizar  # noqa: PLC0415 - evita import circular no topo

    return normalizar(servico)


# --------------------------------------------------------------------------- #
# API pública
# --------------------------------------------------------------------------- #


def gravar(
    servico: str,
    *,
    login: str,
    senha: str,
    site: str = "",
    observacao: str = "",
) -> None:
    """Cria ou substitui a credencial de `servico`."""
    nome = servico.strip()
    if not nome:
        raise ValueError("nome do serviço não pode ser vazio")
    itens = _carregar()
    itens[_chave(nome)] = {
        "servico": nome,
        "login": login,
        "senha": senha,
        "site": site,
        "observacao": observacao,
    }
    _salvar(itens)
    log.info("credencial de %r gravada no cofre", nome)


def obter(servico: str) -> Credencial | None:
    """Devolve a credencial de `servico`, ou None se não estiver no cofre."""
    registro = _carregar().get(_chave(servico))
    if registro is None:
        return None
    return Credencial(
        servico=registro.get("servico", servico),
        login=registro.get("login", ""),
        senha=registro.get("senha", ""),
        site=registro.get("site", ""),
        observacao=registro.get("observacao", ""),
    )


def listar() -> list[str]:
    """Nomes dos serviços guardados, em ordem. NUNCA devolve senha."""
    return sorted(
        (r.get("servico", chave) for chave, r in _carregar().items()),
        key=lambda s: s.casefold(),
    )


def remover(servico: str) -> bool:
    """Apaga a credencial. True se existia, False se não havia nada."""
    itens = _carregar()
    if itens.pop(_chave(servico), None) is None:
        return False
    _salvar(itens)
    log.info("credencial de %r removida do cofre", servico)
    return True


def importar_da_planilha(*, sobrescrever: bool = False) -> dict:
    """
    Popula o cofre com os blocos da aba ACESSOS E CONTATOS.

    NÃO altera a planilha — as senhas continuam lá também, como o usuário
    decidiu. Devolve `{"importados": [...], "ignorados": [...],
    "ja_existiam": [...]}` (só nomes de serviço, nunca senhas).
    """
    from automacao.entrega.planilha_contas import ler_acessos  # noqa: PLC0415

    itens = _carregar()
    importados: list[str] = []
    ignorados: list[str] = []
    ja_existiam: list[str] = []

    for credencial in ler_acessos():
        nome = credencial.servico.strip()
        if not nome:
            continue
        if not credencial.login and not credencial.senha:
            ignorados.append(nome)
            log.info("bloco %r sem login e sem senha — ignorado", nome)
            continue

        chave = _chave(nome)
        if chave in itens and not sobrescrever:
            ja_existiam.append(nome)
            continue

        itens[chave] = {
            "servico": nome,
            "login": credencial.login,
            "senha": credencial.senha,
            "site": credencial.site,
            "observacao": credencial.observacao,
        }
        importados.append(nome)

    if importados:
        _salvar(itens)

    resumo = {
        "importados": importados,
        "ignorados": ignorados,
        "ja_existiam": ja_existiam,
    }
    log.info(
        "importação da planilha: %d importados, %d já existiam, %d ignorados",
        len(importados),
        len(ja_existiam),
        len(ignorados),
    )
    return resumo


def existe_cofre() -> bool:
    return CAMINHO_COFRE.is_file()


# --------------------------------------------------------------------------- #
# Teste de que a senha nunca vaza em repr()/str()
# --------------------------------------------------------------------------- #


def _teste_mascara_senha() -> None:
    """
    Comprova que a senha não escapa por repr()/str()/format()/log.

    Chamado pelo bloco de teste manual. Levanta AssertionError se vazar.
    """
    segredo = "SenhaSuperSecreta#2026"
    cred = Credencial(
        servico="Serviço de Teste",
        login="fulano@exemplo.com.br",
        senha=segredo,
        site="https://exemplo.invalido",
        linha=99,
    )

    for rotulo, texto in (
        ("repr()", repr(cred)),
        ("str()", str(cred)),
        ("f-string", f"{cred}"),
        ("format()", "{}".format(cred)),
        ("%s", "%s" % (cred,)),
        ("dentro de lista", repr([cred])),
        ("dentro de dict", repr({"c": cred})),
        ("como_dict()", repr(cred.como_dict())),
    ):
        assert segredo not in texto, f"SENHA VAZOU em {rotulo}: {texto}"
        assert "***" in texto, f"máscara ausente em {rotulo}: {texto}"

    # O valor continua acessível de propósito, quando pedido explicitamente.
    assert cred.senha == segredo
    assert cred.como_dict(com_senha=True)["senha"] == segredo
    assert cred.como_dict()["senha"] == "***"

    # Credencial sem senha não inventa máscara.
    vazia = Credencial(servico="Sem Senha")
    assert vazia.mascarada() == ""
    assert "***" not in repr(vazia)


# --------------------------------------------------------------------------- #
# Teste manual: python -m automacao.cofre
# --------------------------------------------------------------------------- #

if __name__ == "__main__":
    import sys

    sys.stdout.reconfigure(encoding="utf-8")
    logging.basicConfig(
        level=logging.INFO, format="%(levelname)-7s %(name)s: %(message)s"
    )

    print(f"cofre: {CAMINHO_COFRE}  (existe={existe_cofre()})")
    # Se o cofre ainda não existia, o teste não deixa rastro no disco.
    HAVIA_COFRE = existe_cofre()

    _teste_mascara_senha()
    print("OK  máscara de senha em repr()/str()/como_dict() — nada vazou")

    SERVICO_FICTICIO = "__teste_ficticio__"
    SENHA_FICTICIA = "senha-de-teste-123"

    presentes_antes = listar()
    print(f"serviços no cofre antes: {presentes_antes}")
    assert SERVICO_FICTICIO not in presentes_antes, "sobrou lixo de um teste anterior"

    gravar(
        SERVICO_FICTICIO,
        login="ninguem@exemplo.invalido",
        senha=SENHA_FICTICIA,
        site="https://exemplo.invalido",
        observacao="credencial fictícia de teste",
    )
    print(f"OK  gravado {SERVICO_FICTICIO!r}")

    lido = obter(SERVICO_FICTICIO)
    assert lido is not None, "credencial gravada não foi encontrada"
    assert lido.senha == SENHA_FICTICIA, "senha lida diferente da gravada"
    assert lido.login == "ninguem@exemplo.invalido"
    assert SENHA_FICTICIA not in repr(lido), "senha vazou no repr da credencial lida"
    print(f"OK  lido de volta: {lido!r}")

    assert SERVICO_FICTICIO in listar()
    assert all(SENHA_FICTICIA not in n for n in listar()), "listar() expôs senha"
    print(f"OK  listar() -> {listar()}")

    assert remover(SERVICO_FICTICIO) is True
    assert remover(SERVICO_FICTICIO) is False, "remover duas vezes deveria dar False"
    assert obter(SERVICO_FICTICIO) is None
    assert listar() == presentes_antes, "o cofre não voltou ao estado anterior"
    print(f"OK  removido — cofre de volta a {listar()}")

    if not HAVIA_COFRE:
        CAMINHO_COFRE.unlink(missing_ok=True)
        print("OK  cofre.dat apagado (não existia antes do teste)")

    print("\ntodos os testes do cofre passaram")
    print(
        "\nDica: para trazer a aba ACESSOS E CONTATOS para o cofre, rode\n"
        "      python -c \"from automacao import cofre; "
        'print(cofre.importar_da_planilha())"'
    )
