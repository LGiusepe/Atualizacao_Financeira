"""
Carrega config/settings.yaml e config/fornecedores.yaml.

Regra de ouro deste projeto: nada é escrito fora de `dados/` enquanto
`seguranca.simulacao` for true. Quem grava no OneDrive tem que passar por
`Ambiente.pode_gravar()`.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from dataclasses import asdict
from functools import lru_cache
from pathlib import Path

import yaml

import re

from automacao.nucleo.modelos import (
    Beneficiario,
    Conta,
    ConfigColeta,
    ConfigEmail,
    DadosAutorizacao,
    Pagante,
    normalizar,
)

# automacao/nucleo/config.py -> automacao/nucleo -> automacao -> a raiz.
# Este numero muda se o arquivo mudar de pasta, e dele saem TODOS os
# caminhos do projeto; errar aqui faz o painel procurar config/ no lugar
# errado e reclamar de configuracao ausente.
RAIZ_PROJETO = Path(__file__).resolve().parents[2]
DIR_CONFIG = RAIZ_PROJETO / "config"


class ErroConfiguracao(RuntimeError):
    """Configuração ausente, inválida ou incoerente."""


def _ler_yaml(caminho: Path) -> dict:
    if not caminho.is_file():
        raise ErroConfiguracao(f"arquivo de configuração não encontrado: {caminho}")
    with caminho.open(encoding="utf-8") as fh:
        dados = yaml.safe_load(fh) or {}
    if not isinstance(dados, dict):
        raise ErroConfiguracao(f"{caminho.name} deveria conter um mapeamento no topo")
    return dados


def _resolver(caminho: str | Path) -> Path:
    """Caminho relativo é resolvido a partir da raiz do projeto."""
    p = Path(str(caminho).strip())
    return p if p.is_absolute() else (RAIZ_PROJETO / p)


@dataclass(frozen=True)
class Caminhos:
    onedrive: Path
    autorizacoes: Path
    planilha_contas: Path
    trabalho: Path
    entrada: Path
    backups: Path

    def validar(self) -> list[str]:
        """Devolve a lista de problemas encontrados (vazia = tudo certo)."""
        problemas = []
        if not self.onedrive.is_dir():
            problemas.append(f"pasta do OneDrive não encontrada: {self.onedrive}")
        if not self.autorizacoes.is_dir():
            problemas.append(f"pasta de autorizações não encontrada: {self.autorizacoes}")
        if not self.planilha_contas.is_file():
            problemas.append(f"planilha não encontrada: {self.planilha_contas}")
        return problemas

    def preparar_locais(self) -> None:
        for p in (self.trabalho, self.entrada, self.backups):
            p.mkdir(parents=True, exist_ok=True)


@dataclass(frozen=True)
class Seguranca:
    simulacao: bool = True
    confirmar_antes_de_gravar: bool = True
    backup_antes_de_sobrescrever: bool = True


@dataclass
class Ambiente:
    """Configuração efetiva da execução."""

    caminhos: Caminhos
    email: dict
    pdf: dict
    checklist: dict
    outlook: dict
    seguranca: Seguranca
    #: Bloco `manutencao` do settings.yaml — a limpeza da bancada. Quem o
    #: interpreta é `manutencao.limpeza.Regras`; aqui ele só viaja, para o
    #: núcleo não precisar conhecer as regras de quem apaga.
    manutencao: dict = field(default_factory=dict)
    #: Primeiro mês que o painel oferece no seletor. `None` = sem piso.
    competencia_inicial: str | None = None
    bruto: dict = field(default_factory=dict, repr=False)

    # -- travas de escrita ------------------------------------------------- #

    @staticmethod
    def _dentro_de(alvo: Path, pasta: Path) -> bool:
        """True se `alvo` está em `pasta` (ou é ela mesma), já resolvidos."""
        try:
            alvo.resolve().relative_to(pasta.resolve())
        except (ValueError, OSError):
            return False
        return True

    def e_producao(self, alvo: Path) -> bool:
        """
        True para o que jamais pode ser apagado, sob nenhuma configuração.

        É a pasta das autorizações, a planilha de controle e a raiz do
        OneDrive sincronizado. Esta pergunta é feita ANTES da de área livre:
        se algum dia `caminhos.trabalho` for apontado por engano para dentro
        de "Autorizações de pagamento", a proibição continua valendo — não
        existe combinação de settings.yaml que libere exclusão aqui.
        """
        if alvo.resolve() == self.caminhos.planilha_contas.resolve():
            return True
        return any(self._dentro_de(alvo, raiz) for raiz in self.raizes_de_producao())

    def pastas_do_painel(self) -> tuple[Path, ...]:
        """
        As pastas que são bancada do painel — as únicas onde ele pode apagar.

        `dados/` do projeto (banco, cofre, inventário) mais as três
        configuráveis: trabalho, entrada e backups. Desde 02/09/2026 elas
        podem morar no OneDrive ("Painel Temp"), e é por isso que a área livre
        deixou de ser "dados/" e passou a ser esta lista.
        """
        return (
            RAIZ_PROJETO / "dados",
            self.caminhos.trabalho,
            self.caminhos.entrada,
            self.caminhos.backups,
        )

    def e_area_local(self, destino: Path) -> bool:
        """
        True se o caminho está numa pasta do painel — área livre para escrita.

        Produção nunca é área livre, mesmo que apontada como bancada.
        """
        if self.e_producao(destino):
            return False
        return any(self._dentro_de(destino, pasta) for pasta in self.pastas_do_painel())

    def pode_gravar(self, destino: Path) -> tuple[bool, str]:
        """
        Diz se é permitido escrever em `destino` agora.

        Área local (dados/) é sempre liberada. Fora dela, só com
        `seguranca.simulacao = false`.
        """
        if self.e_area_local(destino):
            return True, ""
        if self.seguranca.simulacao:
            return False, (
                "modo simulação ativo — escrita bloqueada fora de dados/. "
                "Desative em config/settings.yaml (seguranca.simulacao: false) "
                "ou confirme a publicação pelo painel."
            )
        return True, ""

    def exigir_permissao(self, destino: Path) -> None:
        permitido, motivo = self.pode_gravar(destino)
        if not permitido:
            raise PermissionError(f"{destino}: {motivo}")

    # -- proibição absoluta de exclusão ------------------------------------ #

    def exigir_permissao_para_apagar(self, alvo: Path) -> None:
        """
        Recusa apagar qualquer coisa fora das pastas do painel. Sempre.

        Diferente de `exigir_permissao`, esta trava NÃO tem chave: desligar
        `seguranca.simulacao` libera a GRAVAÇÃO no OneDrive, nunca a exclusão.

        A regra sobre a produção é do responsável pelo processo, e é
        categórica: nada em "Autorizações de pagamento" pode ser apagado —
        nem arquivo de conta encerrada, nem pasta de fornecedor
        descontinuado, nem documento substituído. Arquivo que seria
        sobrescrito vira backup antes.

        O que mudou em 02/09/2026: a bancada do painel (trabalho, entrada,
        backups) saiu de dentro do projeto e foi para o OneDrive, para não
        ocupar disco. A área livre passou a ser "as pastas do painel", em vez
        de "dados/". A produção não entrou nessa conta — e `e_producao` é
        consultada primeiro, justamente para que nenhuma configuração possa
        colocá-la lá.
        """
        if self.e_producao(alvo):
            raise PermissionError(
                f"EXCLUSÃO BLOQUEADA (produção): {alvo}\n"
                "Nada em 'Autorizações de pagamento' nem a planilha de "
                "controle podem ser apagados por esta automação, em nenhuma "
                "configuração. Se o arquivo precisa mesmo sair, faça isso à "
                "mão, conferindo antes."
            )
        if self.e_area_local(alvo):
            return
        raise PermissionError(
            f"EXCLUSÃO BLOQUEADA: {alvo}\n"
            "Esta automação só apaga dentro das próprias pastas de trabalho "
            f"({', '.join(str(p) for p in self.pastas_do_painel())}). "
            "Se o arquivo precisa mesmo sair, faça isso à mão."
        )

    def pode_apagar(self, alvo: Path) -> bool:
        return self.e_area_local(alvo)

    # -- atalhos ----------------------------------------------------------- #

    @property
    def simulando(self) -> bool:
        return self.seguranca.simulacao

    def pasta_trabalho(self, conta_id: str, competencia) -> Path:
        p = self.caminhos.trabalho / str(competencia) / conta_id
        p.mkdir(parents=True, exist_ok=True)
        return p

    def raiz_de(self, conta: Conta) -> Path:
        """
        A pasta do OneDrive onde esta conta arquiva.

        Quase todas usam "Autorizações de pagamento". `raiz_onedrive` no
        registro aponta outra — absoluta, ou relativa à raiz do OneDrive.
        """
        if not conta.raiz_onedrive:
            return self.caminhos.autorizacoes
        bruto = Path(os.path.expandvars(conta.raiz_onedrive.strip()))
        return bruto if bruto.is_absolute() else (self.caminhos.onedrive / bruto)

    def pasta_do_fornecedor(self, conta: Conta) -> Path:
        """
        A pasta que guarda todos os meses desta conta.

        É o que a etapa 2 confere antes de prometer criar a pasta do mês, e o
        que o publicador usa como teto ao criar pastas. Com `raiz_onedrive`,
        a raiz JÁ é a pasta do fornecedor — procurar uma subpasta com o nome
        dele lá dentro daria "a pasta do fornecedor não existe" para uma
        conta perfeitamente configurada.
        """
        if conta.raiz_onedrive:
            return self.raiz_de(conta)
        return self.caminhos.autorizacoes / conta.pasta

    def destino_onedrive(self, conta: Conta, competencia) -> Path:
        return self.raiz_de(conta) / conta.caminho_relativo(competencia)

    def raizes_de_producao(self) -> tuple[Path, ...]:
        """
        Todas as pastas de arquivo do processo — as que nunca podem ser
        apagadas. É "Autorizações de pagamento" mais qualquer `raiz_onedrive`
        declarada no registro; sem isto, a pasta de quem mora fora da árvore
        padrão ficaria sem a proteção explícita que as outras têm.
        """
        raizes = {self.caminhos.autorizacoes.resolve()}
        try:
            for c in contas():
                if c.raiz_onedrive:
                    raizes.add(self.raiz_de(c).resolve())
        except Exception:  # noqa: BLE001 — registro ilegível não derruba a trava
            pass
        return tuple(raizes)


def carregar_ambiente(caminho: Path | None = None) -> Ambiente:
    dados = _ler_yaml(caminho or (DIR_CONFIG / "settings.yaml"))

    bloco_caminhos = dados.get("caminhos") or {}
    faltando = [k for k in ("onedrive", "autorizacoes", "planilha_contas") if not bloco_caminhos.get(k)]
    if faltando:
        raise ErroConfiguracao(f"caminhos obrigatórios ausentes em settings.yaml: {faltando}")

    onedrive = Path(os.path.expandvars(str(bloco_caminhos["onedrive"]).strip()))
    caminhos = Caminhos(
        onedrive=onedrive,
        autorizacoes=onedrive / str(bloco_caminhos["autorizacoes"]).strip(),
        planilha_contas=onedrive / str(bloco_caminhos["planilha_contas"]).strip(),
        trabalho=_resolver(bloco_caminhos.get("trabalho", "dados/trabalho")),
        entrada=_resolver(bloco_caminhos.get("entrada", "dados/entrada")),
        backups=_resolver(bloco_caminhos.get("backups", "dados/backups")),
    )
    caminhos.preparar_locais()

    seg = dados.get("seguranca") or {}
    seguranca = Seguranca(
        simulacao=bool(seg.get("simulacao", True)),
        confirmar_antes_de_gravar=bool(seg.get("confirmar_antes_de_gravar", True)),
        backup_antes_de_sobrescrever=bool(seg.get("backup_antes_de_sobrescrever", True)),
    )

    return Ambiente(
        caminhos=caminhos,
        email=dados.get("email") or {},
        pdf=dados.get("pdf") or {},
        checklist=dados.get("checklist") or {},
        outlook=dados.get("outlook") or {},
        seguranca=seguranca,
        manutencao=dados.get("manutencao") or {},
        competencia_inicial=(str(dados["competencia_inicial"]).strip()
                             if dados.get("competencia_inicial") else None),
        bruto=dados,
    )


# --------------------------------------------------------------------------- #
# Registro de contas
# --------------------------------------------------------------------------- #


def _conta_de_dict(bruto: dict) -> Conta:
    if not bruto.get("id"):
        raise ErroConfiguracao(f"conta sem 'id': {bruto!r}")
    if not bruto.get("pasta"):
        raise ErroConfiguracao(f"conta {bruto['id']!r} sem 'pasta'")

    email_bruto = bruto.get("email") or {}
    coleta_bruto = bruto.get("coleta") or {}

    return Conta(
        id=str(bruto["id"]).strip(),
        pasta=str(bruto["pasta"]).strip(),
        subunidade=(bruto.get("subunidade") or None),
        apelido=(bruto.get("apelido") or None),
        ativo=bool(bruto.get("ativo", True)),
        formato_mes=str(bruto.get("formato_mes") or "MM-AAAA").strip(),
        deslocamento_pasta=int(bruto.get("deslocamento_pasta") or 0),
        raiz_onedrive=(bruto.get("raiz_onedrive") or None),
        padrao_nome_arquivo=bruto.get("padrao_nome_arquivo") or None,
        padrao_nome_pasta=bruto.get("padrao_nome_pasta") or None,
        aba_autorizacao=str(bruto.get("aba_autorizacao") or "AUTORIZAÇÃO"),
        modelo_base=bruto.get("modelo_base") or None,
        vencimento_dia=bruto.get("vencimento_dia"),
        valor_estimado=(
            float(bruto["valor_estimado"])
            if bruto.get("valor_estimado") not in (None, "")
            else None
        ),
        autorizacao=DadosAutorizacao.de_dict(bruto.get("autorizacao")),
        email=ConfigEmail(
            eh_operadora=bool(email_bruto.get("eh_operadora") or False),
            cidade=email_bruto.get("cidade") or None,
            nome_empresa=email_bruto.get("nome_empresa") or None,
            grupo=email_bruto.get("grupo") or None,
        ),
        coleta=ConfigColeta(
            remetentes=list(coleta_bruto.get("remetentes") or []),
            assunto_contem=list(coleta_bruto.get("assunto_contem") or []),
            identificadores=list(coleta_bruto.get("identificadores") or []),
        ),
        chaves_planilha=list(
            (bruto.get("planilha_contas") or {}).get("chaves")
            or bruto.get("chaves_planilha")
            or []
        ),
        observacao=str(bruto.get("observacao") or ""),
    )


def carregar_contas(caminho: Path | None = None) -> list[Conta]:
    alvo = caminho or (DIR_CONFIG / "fornecedores.yaml")
    if not alvo.is_file():
        rascunho = DIR_CONFIG / "fornecedores.gerado.yaml"
        if rascunho.is_file():
            raise ErroConfiguracao(
                f"{alvo.name} ainda não existe. Revise {rascunho.name} "
                f"e renomeie para {alvo.name} quando estiver conferido."
            )
        raise ErroConfiguracao(f"registro de contas não encontrado: {alvo}")

    dados = _ler_yaml(alvo)
    contas = [_conta_de_dict(c) for c in (dados.get("contas") or [])]

    duplicados = {c.id for c in contas if [x.id for x in contas].count(c.id) > 1}
    if duplicados:
        raise ErroConfiguracao(f"ids de conta duplicados: {sorted(duplicados)}")
    return contas


@lru_cache(maxsize=1)
def ambiente() -> Ambiente:
    return carregar_ambiente()


@lru_cache(maxsize=1)
def contas() -> tuple[Conta, ...]:
    return tuple(carregar_contas())


def conta_por_id(conta_id: str) -> Conta:
    for c in contas():
        if c.id == conta_id:
            return c
    raise KeyError(f"conta desconhecida: {conta_id!r}")


def contas_do_grupo(grupo: str) -> list[Conta]:
    """
    Contas que saem num e-mail só, com um anexo por fatura.

    Há fornecedor com quatro faturas cujo financeiro recebe uma mensagem única
    com os quatro PDFs. Ordena pela sub-unidade para o corpo do e-mail listar
    sempre na mesma sequência.
    """
    do_grupo = [c for c in contas() if c.ativo and c.email.grupo == grupo]
    return sorted(do_grupo, key=lambda c: (c.subunidade or "", c.id))


def grupos_de_email() -> dict[str, list[Conta]]:
    """Todos os grupos configurados, indexados pelo nome do grupo."""
    agrupadas: dict[str, list[Conta]] = {}
    for c in contas():
        if c.ativo and c.email.grupo:
            agrupadas.setdefault(c.email.grupo, []).append(c)
    for grupo in agrupadas:
        agrupadas[grupo].sort(key=lambda c: (c.subunidade or "", c.id))
    return agrupadas


# --------------------------------------------------------------------------- #
# Pagantes e beneficiários
# --------------------------------------------------------------------------- #

CAMINHO_PAGANTES = DIR_CONFIG / "pagantes.yaml"
CAMINHO_BENEFICIARIOS = DIR_CONFIG / "beneficiarios.yaml"


def _tudo_texto(registro: dict) -> dict:
    """Números que na verdade são códigos voltam a ser texto ao carregar."""
    return {k: (None if v is None else str(v)) for k, v in (registro or {}).items()}


@lru_cache(maxsize=1)
def pagantes() -> tuple[Pagante, ...]:
    """Empresas que pagam, na ordem do arquivo."""
    dados = _ler_yaml(CAMINHO_PAGANTES) if CAMINHO_PAGANTES.is_file() else {}
    return tuple(
        Pagante.de_dict(_tudo_texto(d))
        for d in (dados.get("pagantes") or [])
        if d.get("nome")
    )


@lru_cache(maxsize=1)
def beneficiarios() -> tuple[Beneficiario, ...]:
    """Quem recebe, em ordem alfabética — a lista é longa e vira um select."""
    dados = _ler_yaml(CAMINHO_BENEFICIARIOS) if CAMINHO_BENEFICIARIOS.is_file() else {}
    lista = [
        Beneficiario.de_dict(_tudo_texto(d))
        for d in (dados.get("beneficiarios") or [])
        if d.get("nome")
    ]
    return tuple(sorted(lista, key=lambda b: b.nome.casefold()))


def pagante_por_id(identificador: str) -> Pagante | None:
    return next((p for p in pagantes() if p.id == identificador), None)


def beneficiario_por_id(identificador: str) -> Beneficiario | None:
    return next((b for b in beneficiarios() if b.id == identificador), None)


def _chave_de_nome(nome: str | None) -> str:
    """
    Forma canônica para comparar razão social.

    A mesma empresa aparece escrita de jeitos diferentes nas autorizações:
    `INTELIGÊNCIA` e `INTELIGENCIA`, `A & B` e `A&B`, `... LTDA.` com e sem o
    ponto final. Comparar cru deixava o painel
    avisando "não está no cadastro" para empresa que está, sim.

    Ignora acento, maiúsculas, pontuação e espaço repetido. Não ignora
    `LTDA` vs `S.A.`: isso é outra pessoa jurídica, e passar batido seria
    escrever o CNPJ errado na autorização.
    """
    limpo = normalizar(nome or "")
    return re.sub(r"[^a-z0-9]+", " ", limpo).strip()


def pagante_por_nome(nome: str | None) -> Pagante | None:
    """Casa pelo nome atual OU pelo nome antigo que ficou na tabela da planilha."""
    if not nome:
        return None
    alvo = _chave_de_nome(nome)
    if not alvo:
        return None
    for p in pagantes():
        if alvo in {_chave_de_nome(p.nome), _chave_de_nome(p.nome_na_planilha)}:
            return p
    return None


def beneficiario_por_nome(nome: str | None) -> Beneficiario | None:
    if not nome:
        return None
    alvo = _chave_de_nome(nome)
    if not alvo:
        return None
    return next((b for b in beneficiarios() if _chave_de_nome(b.nome) == alvo), None)


def _gravar_lista(caminho: Path, chave: str, itens: list[dict]) -> None:
    """
    Reescreve um cadastro preservando o cabeçalho explicativo do arquivo.

    Os comentários no topo dizem *por que* cada campo existe — perder isso na
    primeira edição pelo painel deixaria o próximo leitor sem contexto.
    """
    cabecalho = ""
    if caminho.is_file():
        linhas = caminho.read_text(encoding="utf-8").splitlines(keepends=True)
        for linha in linhas:
            if linha.startswith("#") or not linha.strip():
                cabecalho += linha
            else:
                break

    # Tudo vira texto antes de gravar. Agência e conta são identificadores,
    # não números: `0050` sem aspas é lido pelo YAML como octal e vira 40, e
    # `0500` vira 320 — a autorização sairia com a agência errada.
    limpos = [
        {k: (None if v in (None, "") else str(v)) for k, v in item.items()}
        for item in itens
    ]
    corpo = yaml.safe_dump(
        {chave: limpos},
        allow_unicode=True,
        sort_keys=False,
        default_flow_style=False,
        # Sem isto o PyYAML quebra em 80 colunas, e razão social longa vira
        # duas linhas. O YAML entende, mas quem abre o arquivo tropeça.
        width=4096,
        indent=2,
    )
    caminho.parent.mkdir(parents=True, exist_ok=True)
    caminho.write_text(cabecalho + corpo, encoding="utf-8")


def salvar_pagantes(lista: list[Pagante]) -> None:
    _gravar_lista(CAMINHO_PAGANTES, "pagantes", [asdict(p) for p in lista])
    pagantes.cache_clear()


def salvar_beneficiarios(lista: list[Beneficiario]) -> None:
    _gravar_lista(CAMINHO_BENEFICIARIOS, "beneficiarios", [asdict(b) for b in lista])
    beneficiarios.cache_clear()


# --------------------------------------------------------------------------- #
# Edição do settings.yaml pelo painel
# --------------------------------------------------------------------------- #

CAMINHO_SETTINGS = DIR_CONFIG / "settings.yaml"


def _yaml_com_comentarios():
    """
    Leitor/escritor que devolve o arquivo com os comentários no lugar.

    O `settings.yaml` tem 32 comentários explicando cada campo — é metade do
    valor do arquivo. `yaml.safe_dump` apagaria todos na primeira gravação
    pelo painel, então aqui usamos ruamel, que faz ida e volta preservando
    comentários, aspas e blocos de texto.
    """
    from ruamel.yaml import YAML

    y = YAML()
    y.preserve_quotes = True
    y.width = 4096
    y.indent(mapping=2, sequence=4, offset=2)
    return y


def _atribuir(dados: dict, caminho: str, valor) -> tuple[object, object] | None:
    """Escreve `a.b.c` e devolve `(antes, depois)`, ou `None` se nada mudou."""
    partes = caminho.split(".")
    alvo = dados
    for parte in partes[:-1]:
        alvo = alvo[parte]
    antes = alvo.get(partes[-1])
    if antes == valor:
        return None
    alvo[partes[-1]] = valor
    return antes, valor


def salvar_settings(mudancas: dict[str, object]) -> list[str]:
    """
    Grava alterações no `settings.yaml`, preservando o resto do arquivo.

    Faz cópia de segurança antes: `dados/backups/settings-AAAAMMDD-HHMMSS.yaml`.
    Um caminho errado aqui quebra tudo, e voltar tem que ser trivial.

    Args:
        mudancas: `{"email.assunto": "...", "caminhos.onedrive": "..."}`.

    Returns:
        Descrição do que mudou, uma linha por campo. Lista vazia = nada mudou.
    """
    import io
    from datetime import datetime

    y = _yaml_com_comentarios()
    texto = CAMINHO_SETTINGS.read_text(encoding="utf-8")
    dados = y.load(texto)

    aplicadas: list[str] = []
    for caminho, valor in mudancas.items():
        resultado = _atribuir(dados, caminho, valor)
        if resultado is None:
            continue
        antes, depois = resultado
        aplicadas.append(f"{caminho}: {_resumir(antes)} → {_resumir(depois)}")

    if not aplicadas:
        return []

    backups = _resolver("dados/backups")
    backups.mkdir(parents=True, exist_ok=True)
    carimbo = datetime.now().strftime("%Y%m%d-%H%M%S")
    copia = backups / f"settings-{carimbo}.yaml"
    copia.write_text(texto, encoding="utf-8")

    buffer = io.StringIO()
    y.dump(dados, buffer)
    CAMINHO_SETTINGS.write_text(buffer.getvalue(), encoding="utf-8")

    recarregar()
    return aplicadas


def _resumir(valor) -> str:
    """Valor legível para o registro da mudança, sem despejar o texto inteiro."""
    if valor is None or valor == "":
        return "(vazio)"
    if isinstance(valor, list):
        return f"{len(valor)} item(ns)"
    texto = " ".join(str(valor).split())
    return texto if len(texto) <= 70 else texto[:67] + "…"


def recarregar() -> None:
    """Limpa o cache — usar depois de editar os YAML."""
    ambiente.cache_clear()
    contas.cache_clear()
    pagantes.cache_clear()
    beneficiarios.cache_clear()
