"""
Quem enxerga o quê dentro do painel.

O login diz *se* a pessoa entra; isto diz *até onde* ela vai. São duas coisas
diferentes e por isso ficam em arquivos diferentes: `usuarios.yaml` guarda
pessoas, `permissoes.yaml` guarda grupos.

Decisões que valem explicação:

* **A tela some, não só trava.** Oferecer um link que devolve "sem permissão"
  é pior que não oferecer: a pessoa clica, erra, e fica achando que o painel
  quebrou. O menu é montado a partir do que o grupo enxerga.
* **A trava é no servidor, não no menu.** Esconder o link não protege nada — o
  endereço continua digitável. O middleware confere o caminho de cada
  requisição contra as telas do grupo; o menu é só a consequência visível.
* **Administrador vê tudo, sempre.** Não é um grupo editável: se desse para
  tirar telas dele, um clique errado trancaria todo mundo para fora do próprio
  cadastro de usuários, e não sobraria por onde desfazer.
* **Tela nova nasce fechada.** Um `id` que nenhum grupo lista não aparece para
  ninguém além do administrador — que é quem decide liberá-la.
"""

from __future__ import annotations

import logging
import re
import unicodedata
from dataclasses import dataclass, field

import yaml

from automacao.nucleo.config import DIR_CONFIG

log = logging.getLogger("automacao.acesso.permissoes")

CAMINHO_PERMISSOES = DIR_CONFIG / "permissoes.yaml"

#: O grupo que não se edita nem se remove.
GRUPO_ADMIN = "administrador"

#: Em que grupo alguém cai quando o cadastro não diz outro — arquivo antigo,
#: campo em branco. É o mais apertado que ainda deixa trabalhar.
GRUPO_PADRAO = "operador"


@dataclass(frozen=True)
class Tela:
    """
    Uma área do painel, do ponto de vista de quem libera acesso.

    `caminhos` são os prefixos de URL que pertencem a ela — inclusive os de
    POST, senão liberaríamos a leitura e deixaríamos a gravação aberta.
    """

    id: str
    rotulo: str
    apoio: str
    caminhos: tuple[str, ...]
    #: Telas de administração: não entram na lista de nenhum grupo.
    so_admin: bool = False


#: O catálogo. Acrescentar tela aqui é o que a faz aparecer na tela de
#: permissões — e, enquanto nenhum grupo a marcar, ela fica só para o
#: administrador.
TELAS: tuple[Tela, ...] = (
    Tela("contas", "Contas do mês", "a lista do mês e cada conta por dentro",
         ("/", "/conta", "/recarregar")),
    Tela("entrada", "Entrada manual", "subir fatura que não chegou por e-mail",
         ("/entrada",)),
    Tela("auditoria", "Auditoria", "o histórico do que a automação fez",
         ("/auditoria",)),
    Tela("faturas", "Configuração › Faturas ativas",
         "quais contas valem neste mês",
         ("/configuracao", "/configuracao/ativas", "/configuracao/copiar",
          "/configuracao/nova")),
    Tela("pagantes", "Configuração › Pagantes", "quem paga, com CNPJ",
         ("/configuracao/pagantes",)),
    Tela("beneficiarios", "Configuração › Beneficiários",
         "quem recebe e dados bancários", ("/configuracao/beneficiarios",)),
    Tela("parametros", "Configuração › Parâmetros do sistema",
         "caminhos, e-mail e travas", ("/configuracao/parametros",)),
    Tela("email", "Configuração › E-mail",
         "sua assinatura e o seu texto padrão", ("/configuracao/email",)),
    Tela("manutencao", "Configuração › Manutenção",
         "limpar a bancada de arquivos já publicados",
         ("/configuracao/manutencao",)),
    Tela("diagnostico", "Diagnóstico", "o que está no ar e o que falta",
         ("/diagnostico",)),
    Tela("usuarios", "Configuração › Usuários", "quem entra no painel",
         ("/configuracao/usuarios",), so_admin=True),
    Tela("permissoes", "Configuração › Permissões", "o que cada grupo enxerga",
         ("/configuracao/permissoes",), so_admin=True),
)

POR_ID = {t.id: t for t in TELAS}

#: Só estas entram na tela de permissões — as de administração não se negociam.
TELAS_LIBERAVEIS = tuple(t for t in TELAS if not t.so_admin)

#: Aberto a qualquer pessoa já logada: sair, trocar a própria senha, o sinal de
#: vida e os arquivos estáticos. Barrar isto trancaria alguém sem grupo dentro
#: do painel, sem nem conseguir sair.
SEMPRE_LIBERADO = ("/entrar", "/sair", "/trocar-senha", "/static", "/vivo",
                   "/favicon.ico")


@dataclass
class Grupo:
    """Um conjunto de telas com nome. `chave` é o que fica no cadastro."""

    chave: str
    nome: str
    telas: set[str] = field(default_factory=set)

    @property
    def e_admin(self) -> bool:
        return self.chave == GRUPO_ADMIN

    def ve(self, tela_id: str) -> bool:
        return self.e_admin or tela_id in self.telas

    @property
    def telas_ordenadas(self) -> list[Tela]:
        """Na ordem do catálogo, não na ordem em que foram marcadas."""
        if self.e_admin:
            return list(TELAS)
        return [t for t in TELAS if t.id in self.telas]


def normalizar(nome: str) -> str:
    """
    Nome de grupo vira chave: sem acento, sem espaço, minúsculo.

    O cadastro guarda a chave, não o nome digitado — assim corrigir "Operador"
    para "Operadores" na tela de permissões não desliga ninguém do grupo.
    """
    cru = unicodedata.normalize("NFKD", (nome or "").strip())
    cru = "".join(c for c in cru if not unicodedata.combining(c))
    cru = re.sub(r"[^A-Za-z0-9]+", "-", cru).strip("-").lower()
    return cru


#: Com que o painel começa quando ainda não existe `permissoes.yaml`. O
#: operador toca o mês inteiro; o de consulta só olha. Nenhum dos dois entra em
#: Parâmetros, que mexe em caminho de rede e em trava de envio.
GRUPOS_PADRAO: tuple[tuple[str, str, tuple[str, ...]], ...] = (
    (GRUPO_ADMIN, "Administrador", ()),
    ("operador", "Operador",
     ("contas", "entrada", "auditoria", "faturas", "pagantes",
      "beneficiarios", "email", "diagnostico")),
    ("consulta", "Consulta", ("contas", "auditoria")),
)


def _ordenar(grupos: list[Grupo]) -> list[Grupo]:
    """
    Administrador na frente, o resto em ordem de nome.

    Vale também para os padrões: sem isto, a lista mudava de ordem na primeira
    gravação e o seletor de permissão parecia ter se embaralhado sozinho.
    """
    return sorted(grupos, key=lambda g: (not g.e_admin, g.nome.lower()))


def _padrao() -> list[Grupo]:
    return _ordenar(
        [Grupo(chave, nome, set(telas)) for chave, nome, telas in GRUPOS_PADRAO]
    )


def carregar() -> list[Grupo]:
    """
    Os grupos do arquivo, com o administrador sempre na frente.

    Arquivo ausente devolve os padrões em vez de lista vazia: sem nenhum grupo
    não haveria o que escolher ao cadastrar alguém.
    """
    if not CAMINHO_PERMISSOES.is_file():
        return _padrao()
    try:
        dados = yaml.safe_load(CAMINHO_PERMISSOES.read_text(encoding="utf-8")) or {}
    except yaml.YAMLError as erro:
        log.error("%s ilegível (%s) — valendo os grupos padrão",
                  CAMINHO_PERMISSOES.name, erro)
        return _padrao()

    grupos: list[Grupo] = []
    for bruto in dados.get("grupos") or []:
        chave = normalizar(bruto.get("chave") or bruto.get("nome") or "")
        if not chave:
            continue
        # Id fora do catálogo é descartado em silêncio: tela renomeada ou
        # removida não pode derrubar a leitura do arquivo inteiro.
        telas = {t for t in (bruto.get("telas") or []) if t in POR_ID}
        grupos.append(Grupo(chave, str(bruto.get("nome") or chave).strip(), telas))

    if not any(g.e_admin for g in grupos):
        grupos.append(Grupo(GRUPO_ADMIN, "Administrador", set()))
    return _ordenar(grupos)


def gravar(grupos: list[Grupo]) -> None:
    """Reescreve o arquivo inteiro, com o aviso no topo."""
    CAMINHO_PERMISSOES.parent.mkdir(parents=True, exist_ok=True)
    cabecalho = (
        "# O que cada grupo de usuário enxerga no painel.\n"
        "#\n"
        "# 'telas' sao os ids do catalogo em automacao/permissoes.py. O grupo\n"
        "# 'administrador' ve tudo por construcao — a lista dele e ignorada.\n"
        "# Quem esta em cada grupo fica em usuarios.yaml, campo 'grupo'.\n"
    )
    corpo = yaml.safe_dump(
        {"grupos": [
            {"chave": g.chave, "nome": g.nome,
             "telas": [t.id for t in TELAS if t.id in g.telas]}
            for g in grupos
        ]},
        allow_unicode=True, sort_keys=False, width=4096,
    )
    CAMINHO_PERMISSOES.write_text(cabecalho + corpo, encoding="utf-8")


def por_chave(chave: str) -> Grupo | None:
    alvo = normalizar(chave)
    return next((g for g in carregar() if g.chave == alvo), None)


def grupo_de(usuario) -> Grupo:
    """
    O grupo de alguém.

    Grupo apagado depois do cadastro não vira acesso total: devolve um grupo
    vazio, e a pessoa cai na tela de "sem permissão" até um administrador
    arrumar. O contrário — abrir tudo — seria a falha calada de sempre.
    """
    if usuario is None:
        return Grupo("", "", set())
    chave = normalizar(getattr(usuario, "grupo", "") or "")
    achado = por_chave(chave)
    if achado is not None:
        return achado
    return Grupo(chave, "(grupo removido)", set())


def ve(usuario, tela_id: str) -> bool:
    """Se a pessoa enxerga uma tela do catálogo."""
    return grupo_de(usuario).ve(tela_id)


def telas_de(usuario) -> list[Tela]:
    return grupo_de(usuario).telas_ordenadas


def tela_do_caminho(caminho: str) -> Tela | None:
    """
    A qual tela pertence uma URL. `None` = caminho fora do catálogo.

    Casa pelo prefixo mais longo: '/configuracao/pagantes' também começa com
    '/configuracao', e a ordem ingênua liberaria Pagantes para quem só tem
    Faturas ativas.
    """
    melhor: tuple[int, Tela] | None = None
    for tela in TELAS:
        for prefixo in tela.caminhos:
            casa = (caminho == prefixo if prefixo == "/"
                    else caminho == prefixo or caminho.startswith(prefixo + "/"))
            if casa and (melhor is None or len(prefixo) > melhor[0]):
                melhor = (len(prefixo), tela)
    return melhor[1] if melhor else None


def pode_acessar(usuario, caminho: str) -> bool:
    """
    Se a requisição passa.

    Caminho fora do catálogo fica liberado a quem já entrou: negar por omissão
    trancaria o painel a cada endereço novo, e o catálogo é justamente o que se
    revisa quando uma tela nasce.
    """
    if caminho.startswith(SEMPRE_LIBERADO):
        return True
    tela = tela_do_caminho(caminho)
    if tela is None:
        return True
    if tela.so_admin:
        return grupo_de(usuario).e_admin
    return ve(usuario, tela.id)


def primeiro_caminho(usuario) -> str:
    """
    Para onde mandar alguém que caiu numa tela fechada. Vazio = não há nenhuma.
    """
    telas = telas_de(usuario)
    if not telas:
        return ""
    prefixo = telas[0].caminhos[0]
    # '/conta' é detalhe de conta, que sozinho não abre nada: a porta é a lista.
    return "/" if prefixo == "/conta" else prefixo


def em_uso(chave: str) -> list[str]:
    """Quem está no grupo. Serve para não apagar um grupo com gente dentro."""
    from automacao.acesso import usuarios as u

    alvo = normalizar(chave)
    return [x.email for x in u.carregar() if normalizar(x.grupo) == alvo]
