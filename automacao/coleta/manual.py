"""
Coleta manual: a pasta `dados/entrada/`.

Nem toda fatura chega por e-mail. Várias são baixadas no portal do fornecedor,
e o caminho para elas é: o usuário joga o PDF em `dados/entrada/` e o painel
faz o resto.

Duas formas de usar:

* **Solta na raiz** — `dados/entrada/boleto.pdf`. O módulo classifica e usa
  `sugerir_conta()` para chutar de qual conta é; quem decide é o humano, pelo
  painel (`/entrada`), que chama `vincular()` com o que ele marcou.
* **Dentro de uma subpasta com o id da conta** — `dados/entrada/claro-guarulhos/`.
  A subpasta ajuda a saber de que conta o arquivo é, e `conta_id_da_subpasta()`
  continua lendo isso. Ela **não** é mais varrida por ninguém: até 14/09/2026 a
  pasta inteira entrava sozinha na competência aberta, e como a pasta era a
  mesma para todos os meses, o que ninguém recolheu de um mês entrou no
  seguinte — uma autorização saiu preenchida com os dados da fatura anterior.
  Hoje o caminho normal da fatura é o upload na própria etapa 1, que grava
  direto em `trabalho/<competencia>/<conta_id>/`.

Regra de ouro do ciclo de vida: `vincular()` **copia** para a pasta de trabalho.
O original só sai da entrada em `arquivar_entrada()`, que o painel chama depois
que a publicação deu certo. Se algo falhar no meio, o arquivo continua lá.

Nada aqui escreve fora de `dados/`. A única leitura do OneDrive é a aba
BENEFICIÁRIOS de um modelo de autorização, usada por `sugerir_conta()` para
saber o CNPJ de quem cobra — e ela é opcional: sem OneDrive, o palpite
continua funcionando com os outros sinais.
"""

from __future__ import annotations

import logging
import re
import shutil
from pathlib import Path
from typing import Iterable, Sequence

from automacao.coleta.classificador import RE_CNPJ, classificar, extrair_texto_primeira_pagina
from automacao.coleta.outlook import copiar_sem_duplicar, nome_livre
from automacao.nucleo.config import Ambiente, ambiente
from automacao.nucleo.modelos import (
    Competencia,
    Conta,
    Documento,
    Etapa,
    OrigemDocumento,
    ResultadoEtapa,
    normalizar,
)

log = logging.getLogger("automacao.coleta.manual")

__all__ = [
    "varrer_entrada",
    "vincular",
    "arquivar_entrada",
    "sugerir_conta",
    "conta_id_da_subpasta",
    "mapa_beneficiarios",
    "esquecer_caches",
    "PASTA_PROCESSADOS",
]

# Subpasta onde os arquivos já publicados vão descansar. Começa com "_" para
# ficar no topo da listagem e para a varredura saber que deve ignorá-la.
PASTA_PROCESSADOS = "_processados"

# Extensões que interessam na entrada manual.
EXTENSOES_ACEITAS = {".pdf", ".xml"}

# Nome da aba que, em todo modelo de autorização, mapeia beneficiário → CNPJ.
ABA_BENEFICIARIOS = "BENEFICIÁRIOS"

# Palavras que não ajudam a identificar conta nenhuma.
RUIDO = {
    "ltda", "sa", "s/a", "me", "epp", "eireli", "the", "com", "and",
    "de", "da", "do", "das", "dos", "e", "em", "no", "na", "para",
    "fatura", "boleto", "nota", "fiscal", "nf", "nfe", "nfse", "pdf",
}

# Quantos pontos equivalem a "certeza" numa sugestão de conta.
PONTOS_PARA_CERTEZA = 10.0


# --------------------------------------------------------------------------- #
# Leitura da pasta de entrada
# --------------------------------------------------------------------------- #


def _pasta_entrada(amb: Ambiente | None = None) -> Path:
    entrada = (amb or ambiente()).caminhos.entrada
    entrada.mkdir(parents=True, exist_ok=True)
    return entrada


def conta_id_da_subpasta(caminho: Path, *, ambiente_: Ambiente | None = None) -> str | None:
    """
    Id da conta deduzido da subpasta de `dados/entrada/`.

    `dados/entrada/claro-guarulhos/fatura.pdf` devolve `'claro-guarulhos'`.
    Arquivo na raiz da entrada (ou fora dela) devolve `None`.
    """
    entrada = _pasta_entrada(ambiente_)
    try:
        relativo = Path(caminho).resolve().relative_to(entrada.resolve())
    except ValueError:
        return None
    partes = relativo.parts
    if len(partes) < 2:
        return None            # está solto na raiz
    if partes[0] == PASTA_PROCESSADOS:
        return None            # já foi processado
    return partes[0]


def _arquivos_uteis(raiz: Path, entrada: Path) -> list[Path]:
    """Arquivos aceitos dentro de `raiz`, ignorando `_processados/`."""
    return sorted(
        arquivo
        for arquivo in raiz.rglob("*")
        if arquivo.is_file()
        and arquivo.suffix.lower() in EXTENSOES_ACEITAS
        and PASTA_PROCESSADOS not in arquivo.relative_to(entrada).parts
    )


def varrer_entrada(*, ambiente_: Ambiente | None = None) -> list[Documento]:
    """
    Classifica tudo que está em `dados/entrada/`, inclusive nas subpastas.

    Ignora `_processados/`. Cada `Documento` volta com
    `origem = ENTRADA_MANUAL`; quando o arquivo estava numa subpasta com id de
    conta, isso aparece no começo do `motivo`.

    Args:
        ambiente_: ambiente já carregado (útil em teste); padrão é `ambiente()`.

    Returns:
        Lista ordenada por caminho — a mesma ordem que o usuário vê no Explorer.
    """
    amb = ambiente_ or ambiente()
    entrada = _pasta_entrada(amb)

    documentos: list[Documento] = []
    for arquivo in _arquivos_uteis(entrada, entrada):
        conta_id = conta_id_da_subpasta(arquivo, ambiente_=amb)
        documento = classificar(arquivo, contexto=conta_id or "")
        documento.origem = OrigemDocumento.ENTRADA_MANUAL
        if conta_id:
            documento.motivo = f"pasta '{conta_id}' indica a conta; {documento.motivo}"
        documentos.append(documento)

    log.info("entrada manual: %d arquivo(s) em %s", len(documentos), entrada)
    return documentos


# --------------------------------------------------------------------------- #
# Aba BENEFICIÁRIOS — de quem é o CNPJ que aparece no documento
# --------------------------------------------------------------------------- #

# Lido uma vez por execução: são 50+ contas e o arquivo está no OneDrive.
_cache_beneficiarios: dict[str, str] | None = None
# Texto da 1ª página, por (caminho, tamanho, mtime) — o painel chama
# `sugerir_conta()` uma vez por documento a cada carregamento de página.
_cache_texto: dict[tuple, str] = {}


def esquecer_caches() -> None:
    """Zera os caches de leitura — usar depois de mexer nos arquivos."""
    global _cache_beneficiarios
    _cache_beneficiarios = None
    _cache_texto.clear()


def _ler_aba_beneficiarios(caminho_xlsx: Path) -> dict[str, str]:
    """`{nome normalizado do beneficiário: CNPJ}` de um modelo de autorização."""
    import openpyxl

    achados: dict[str, str] = {}
    livro = openpyxl.load_workbook(caminho_xlsx, read_only=True, data_only=True)
    try:
        if ABA_BENEFICIARIOS not in livro.sheetnames:
            return achados
        for linha in livro[ABA_BENEFICIARIOS].iter_rows(values_only=True):
            if not linha or len(linha) < 2:
                continue
            nome, cnpj = linha[0], linha[1]
            if not nome or not cnpj:
                continue
            chave = normalizar(str(nome))
            if chave and chave != "nome":
                achados.setdefault(chave, str(cnpj).strip())
    finally:
        livro.close()
    return achados


def mapa_beneficiarios(
    contas: Sequence[Conta],
    *,
    ambiente_: Ambiente | None = None,
) -> dict[str, str]:
    """
    Tabela beneficiário → CNPJ, lida da aba BENEFICIÁRIOS de um modelo real.

    A aba é a mesma em todos os modelos de autorização, então basta abrir o
    primeiro que existir. É leitura pura no OneDrive e é **opcional**: se o
    OneDrive não estiver montado, ou nenhum `modelo_base` existir, devolvemos
    um mapa vazio e o palpite de conta segue com os outros sinais.

    Args:
        contas: contas candidatas — usamos o `modelo_base` delas para achar o xlsx.
        ambiente_: ambiente já carregado (útil em teste).

    Returns:
        `{nome normalizado: CNPJ formatado}`. Vazio quando não deu para ler.
    """
    global _cache_beneficiarios
    if _cache_beneficiarios is not None:
        return _cache_beneficiarios

    amb = ambiente_ or ambiente()
    _cache_beneficiarios = {}
    for conta in contas:
        if not conta.modelo_base:
            continue
        bruto = Path(conta.modelo_base)
        modelo = bruto if bruto.is_absolute() else (amb.caminhos.autorizacoes / bruto)
        if not modelo.is_file():
            continue
        try:
            _cache_beneficiarios = _ler_aba_beneficiarios(modelo)
        except Exception as exc:  # noqa: BLE001 — sinal opcional, nunca derruba o painel
            log.debug("não consegui ler a aba %s de %s: %s", ABA_BENEFICIARIOS, modelo.name, exc)
            continue
        if _cache_beneficiarios:
            log.info(
                "aba %s lida de %s: %d beneficiário(s)",
                ABA_BENEFICIARIOS, modelo.name, len(_cache_beneficiarios),
            )
            break

    if not _cache_beneficiarios:
        log.debug("aba %s indisponível — palpite seguirá sem o CNPJ do cobrador", ABA_BENEFICIARIOS)
    return _cache_beneficiarios


# --------------------------------------------------------------------------- #
# Sugestão de conta
# --------------------------------------------------------------------------- #


def _texto_do_documento(documento: Documento) -> str:
    """Texto da 1ª página, normalizado. XML e ilegíveis devolvem string vazia."""
    caminho = Path(documento.caminho)
    if caminho.suffix.lower() != ".pdf":
        return ""
    try:
        info = caminho.stat()
        chave = (str(caminho.resolve()), info.st_size, info.st_mtime_ns)
    except OSError:
        return ""
    if chave in _cache_texto:
        return _cache_texto[chave]

    texto, _paginas, _erro = extrair_texto_primeira_pagina(caminho)
    _cache_texto[chave] = normalizar(texto)
    return _cache_texto[chave]


def _tokens(*textos: str | None) -> set[str]:
    """Palavras úteis de um texto: sem acento, sem ruído, com 3+ caracteres."""
    achados: set[str] = set()
    for texto in textos:
        if not texto:
            continue
        for pedaco in re.split(r"[^a-z0-9]+", normalizar(texto)):
            if len(pedaco) >= 3 and pedaco not in RUIDO:
                achados.add(pedaco)
    return achados


def _so_digitos(texto: str) -> str:
    return re.sub(r"\D", "", texto)


def _cnpjs_do_beneficiario(texto_normalizado: str) -> list[str]:
    """
    CNPJs que aparecem perto de quem está cobrando.

    O CNPJ do pagador (a própria associação) está em todo documento e não
    identifica conta nenhuma; o que interessa é o do beneficiário/prestador.
    """
    rotulos = ("beneficiario", "prestador", "cedente", "emitente", "fornecedor", "sacador")
    perto: list[str] = []
    for rotulo in rotulos:
        for encontro in re.finditer(rotulo, texto_normalizado):
            trecho = texto_normalizado[max(0, encontro.start() - 120): encontro.end() + 200]
            for cnpj in RE_CNPJ.findall(trecho):
                if cnpj not in perto:
                    perto.append(cnpj)
    return perto


def _pontos_da_subunidade(tokens: set[str], nome: str, texto: str) -> float:
    """
    Quanto a subunidade da conta combina com o arquivo.

    É o sinal que separa contas irmãs — as quatro de um fornecedor, as seis de
    outro, as seis de um terceiro. Palavra que aparece soma; palavra que falta
    **desconta**, senão "LINK" e "LINK MULTIMIDIA" empatam para sempre.
    """
    if not tokens:
        return 0.0
    no_nome = {t for t in tokens if t in nome}
    no_texto = {t for t in tokens if t in texto} - no_nome
    pontos = 2.0 * len(no_nome) + 1.0 * len(no_texto)
    faltando = len(tokens) - len(no_nome) - len(no_texto)
    if faltando:
        pontos -= 0.75 * faltando
    elif pontos:
        pontos += 1.5   # a subunidade inteira apareceu: é ela mesma
    return pontos


def sugerir_conta(
    documento: Documento,
    contas: Sequence[Conta],
    *,
    ambiente_: Ambiente | None = None,
) -> list[tuple[Conta, float]]:
    """
    Ranqueia quais contas combinam com o arquivo, da melhor para a pior.

    Sinais cruzados, do mais decisivo para o mais fraco:

    1. Subpasta com o id da conta — se existir, acabou: confiança 1.0.
    2. Identificadores da conta (nº da linha, nº do contrato) no nome ou no texto.
    3. CNPJ do beneficiário: o da aba BENEFICIÁRIOS para
       `conta.autorizacao.beneficiario`, e os que aparecem rotulados no PDF.
    4. Palavras da subunidade e da pasta no nome do arquivo e no texto.
    5. Textos usados na planilha CONTAS E ACESSOS e o nome do beneficiário.

    Args:
        documento: já classificado por `varrer_entrada()` ou `classificar()`.
        contas: universo de contas candidatas (normalmente `config.contas()`).
        ambiente_: ambiente já carregado (útil em teste).

    Returns:
        Lista `(conta, confianca)` ordenada da melhor para a pior, só com
        confiança acima de zero. Quando o arquivo veio de uma subpasta com o id
        da conta, devolve essa conta sozinha com confiança 1.0.
    """
    amb = ambiente_ or ambiente()

    # Caminho feliz: o usuário já disse de quem é, pondo na subpasta certa.
    conta_id = conta_id_da_subpasta(documento.caminho, ambiente_=amb)
    if conta_id:
        for conta in contas:
            if conta.id == conta_id:
                log.info("%s vinculado por subpasta a %s", documento.nome, conta.id)
                return [(conta, 1.0)]
        log.warning(
            "subpasta %r não corresponde a nenhuma conta conhecida — caindo no palpite",
            conta_id,
        )

    nome = normalizar(Path(documento.caminho).stem)
    nome_digitos = _so_digitos(nome)
    texto = _texto_do_documento(documento)
    texto_digitos = _so_digitos(texto)
    cnpjs_rotulados = _cnpjs_do_beneficiario(texto)
    cnpjs_no_texto = {_so_digitos(c) for c in RE_CNPJ.findall(texto)}
    beneficiarios = mapa_beneficiarios(contas, ambiente_=amb)

    ranking: list[tuple[Conta, float]] = []
    for conta in contas:
        pontos = 0.0

        # 1) identificadores (nº da linha / contrato) — o sinal mais decisivo
        for identificador in conta.coleta.identificadores:
            digitos = _so_digitos(identificador)
            alvo = normalizar(identificador)
            if alvo and alvo in nome:
                pontos += 3.5
            elif digitos and len(digitos) >= 5 and digitos in nome_digitos:
                pontos += 3.5
            elif alvo and alvo in texto:
                pontos += 3.0
            elif digitos and len(digitos) >= 5 and digitos in texto_digitos:
                pontos += 3.0

        # 2) CNPJ do beneficiário desta conta, segundo a aba BENEFICIÁRIOS
        cnpj_oficial = beneficiarios.get(normalizar(conta.autorizacao.beneficiario or ""))
        if cnpj_oficial and _so_digitos(cnpj_oficial) in cnpjs_no_texto:
            pontos += 3.5

        # 3) nome do arquivo × subunidade e pasta da conta
        pontos += _pontos_da_subunidade(_tokens(conta.subunidade), nome, texto)
        for token in _tokens(conta.pasta):
            if token in nome:
                pontos += 1.5
            elif token in texto:
                pontos += 1.0

        # 4) nomes usados na planilha de controle
        for chave in conta.chaves_planilha:
            chave_n = normalizar(chave)
            if not chave_n:
                continue
            if chave_n in nome:
                pontos += 2.5
            elif chave_n in texto:
                pontos += 2.0

        # 5) beneficiário da autorização escrito no documento
        tokens_beneficiario = _tokens(conta.autorizacao.beneficiario)
        if tokens_beneficiario:
            batidos = sum(1 for t in tokens_beneficiario if t in texto)
            pontos += min(3.0, batidos * 1.0)

        # 6) CNPJ rotulado no PDF citado na configuração da conta
        referencias = normalizar(
            " ".join(conta.coleta.identificadores)
            + " " + " ".join(conta.chaves_planilha)
            + " " + conta.observacao
        )
        referencias_digitos = _so_digitos(referencias)
        for cnpj in cnpjs_rotulados:
            if cnpj in referencias or _so_digitos(cnpj) in referencias_digitos:
                pontos += 3.0

        if pontos > 0:
            ranking.append((conta, round(min(0.99, pontos / PONTOS_PARA_CERTEZA), 2)))

    ranking.sort(key=lambda par: (-par[1], par[0].id))
    log.debug(
        "sugestões para %s: %s",
        documento.nome,
        [(c.id, s) for c, s in ranking[:3]],
    )
    return ranking


# --------------------------------------------------------------------------- #
# Vínculo
# --------------------------------------------------------------------------- #


def vincular(
    documentos: Iterable[Documento],
    conta: Conta,
    competencia: Competencia,
    *,
    ambiente_: Ambiente | None = None,
) -> ResultadoEtapa:
    """
    Copia os arquivos escolhidos para a pasta de trabalho da conta.

    **Copia, não move.** O original fica em `dados/entrada/` até a publicação
    dar certo — quem tira de lá é `arquivar_entrada()`. Assim, se o processo
    quebrar no meio do caminho, nada se perde.

    Os `Documento` da lista são atualizados para apontar para a cópia.

    Args:
        documentos: os que o humano confirmou que são desta conta (podem estar
            soltos na raiz de `dados/entrada/`).
        conta: conta de destino.
        competencia: mês de referência.
        ambiente_: ambiente já carregado (útil em teste).

    Returns:
        `ResultadoEtapa` de `Etapa.COLETA`, com os arquivos copiados em
        `artefatos` e os documentos atualizados em `detalhes["documentos"]`.
    """
    amb = ambiente_ or ambiente()
    etapa = Etapa.COLETA
    lista = list(documentos)

    if not lista:
        return ResultadoEtapa.pulado(
            etapa,
            "nenhum documento selecionado na entrada manual",
            detalhes={"documentos": []},
        )

    destino = amb.pasta_trabalho(conta.id, competencia)
    originais = [str(d.caminho) for d in lista]
    copiados: list[Path] = []
    falhas: list[str] = []

    for documento in lista:
        origem = Path(documento.caminho)
        if not origem.is_file():
            falhas.append(f"{origem.name}: sumiu da entrada antes de ser copiado")
            continue
        try:
            alvo, _copiou = copiar_sem_duplicar(origem, destino)
        except OSError as exc:
            falhas.append(f"{origem.name}: {exc}")
            log.warning("falha ao copiar %s: %s", origem, exc)
            continue

        # O documento passa a viver na pasta de trabalho; o original continua
        # na entrada, esperando o arquivamento.
        documento.caminho = alvo
        documento.origem = OrigemDocumento.ENTRADA_MANUAL
        copiados.append(alvo)
        log.info("copiado para a pasta de trabalho: %s", alvo.name)

    detalhes = {
        "documentos": lista,
        "destino": str(destino),
        "originais_na_entrada": originais,
        "falhas": falhas,
    }

    if not copiados:
        return ResultadoEtapa.erro(
            etapa,
            "nenhum arquivo pôde ser copiado: " + "; ".join(falhas),
            detalhes=detalhes,
        )
    if falhas:
        return ResultadoEtapa.atencao(
            etapa,
            f"{len(copiados)} arquivo(s) copiado(s) para {destino}, "
            f"mas {len(falhas)} falhou(aram): " + "; ".join(falhas),
            detalhes=detalhes,
            artefatos=copiados,
        )
    return ResultadoEtapa.sucesso(
        etapa,
        f"{len(copiados)} arquivo(s) copiado(s) da entrada manual para {destino}",
        detalhes=detalhes,
        artefatos=copiados,
    )


# --------------------------------------------------------------------------- #
# Arquivamento
# --------------------------------------------------------------------------- #


def _dentro_da_entrada(caminho: Path, entrada: Path) -> bool:
    """True se o caminho está em `dados/entrada/`, fora de `_processados`."""
    try:
        relativo = caminho.resolve().relative_to(entrada.resolve())
    except (ValueError, OSError):
        return False
    return PASTA_PROCESSADOS not in relativo.parts and caminho.is_file()


def arquivar_entrada(
    documentos: Iterable[Documento],
    competencia: Competencia | None = None,
    *,
    ambiente_: Ambiente | None = None,
) -> list[Path]:
    """
    Tira da entrada os arquivos já publicados.

    Move para `dados/entrada/_processados/<competencia>/`. Deve ser chamada
    **somente depois** que a publicação der certo — antes disso o original tem
    que continuar em `dados/entrada/`, porque é a única cópia que sobra se algo
    quebrar no meio.

    Nunca sobrescreve: se já existe um arquivo com aquele nome em
    `_processados/`, o novo ganha sufixo (`boleto.pdf`, `boleto_2.pdf`…).

    Aceita tanto os `Documento` originais da entrada quanto os já atualizados
    por `vincular()`: nesse segundo caso procuramos, dentro de
    `dados/entrada/`, o arquivo de mesmo nome. Nada fora da entrada é tocado —
    a cópia na pasta de trabalho e o OneDrive ficam como estão.

    Args:
        documentos: documentos a arquivar.
        competencia: mês de referência; sem ela, usa a competência atual.
        ambiente_: ambiente já carregado (útil em teste).

    Returns:
        Os caminhos novos, dentro de `_processados/`. Lista vazia quando não
        havia nada da entrada para arquivar.
    """
    amb = ambiente_ or ambiente()
    competencia = competencia or Competencia.atual()
    entrada = _pasta_entrada(amb)
    guarda = entrada / PASTA_PROCESSADOS / str(competencia)

    alvos: list[Path] = []
    for documento in documentos:
        caminho = Path(documento.caminho)
        if _dentro_da_entrada(caminho, entrada):
            alvos.append(caminho.resolve())
            continue
        # Documento já apontando para a cópia na pasta de trabalho: procuramos
        # o original pelo nome, inclusive nas subpastas por conta.
        alvos.extend(
            candidato.resolve()
            for candidato in entrada.rglob(caminho.name)
            if candidato.is_file()
            and PASTA_PROCESSADOS not in candidato.relative_to(entrada).parts
        )

    unicos = list(dict.fromkeys(alvos))  # sem repetir, mantendo a ordem
    if not unicos:
        log.info("nada para arquivar: nenhum destes documentos está em %s", entrada)
        return []

    guarda.mkdir(parents=True, exist_ok=True)
    movidos: list[Path] = []
    for origem in unicos:
        try:
            destino = nome_livre(guarda, origem.name)
            shutil.move(str(origem), str(destino))
        except OSError as exc:
            log.warning("não consegui arquivar %s: %s", origem, exc)
            continue
        movidos.append(destino)
        log.info("arquivado: %s -> %s", origem.name, destino)

    log.info("%d arquivo(s) movido(s) para %s", len(movidos), guarda)
    return movidos


# --------------------------------------------------------------------------- #
# Teste manual
# --------------------------------------------------------------------------- #

if __name__ == "__main__":
    from automacao.nucleo.config import DIR_CONFIG, carregar_contas

    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")

    registro = DIR_CONFIG / "fornecedores.yaml"
    if not registro.is_file():
        registro = DIR_CONFIG / "fornecedores.gerado.yaml"
    todas = carregar_contas(registro)
    print(f"{len(todas)} contas carregadas de {registro.name}\n")

    achados = varrer_entrada()
    if not achados:
        print("dados/entrada/ está vazia — jogue alguns PDFs lá e rode de novo.")
        raise SystemExit(0)

    for documento in achados:
        print(f"{documento.nome}")
        print(f"  tipo   : {documento.tipo.value} (confiança {documento.confianca:.2f})")
        print(f"  motivo : {documento.motivo[:160]}")
        palpites = sugerir_conta(documento, todas)
        if palpites:
            print("  contas sugeridas:")
            for conta, nota in palpites[:5]:
                print(f"    {nota:.2f}  {conta.id}  ({conta.rotulo})")
        else:
            print("  contas sugeridas: nenhuma — escolha na mão no painel")
        print()
