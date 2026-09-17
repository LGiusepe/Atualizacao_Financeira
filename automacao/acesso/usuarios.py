"""
Quem pode abrir o painel.

O painel escuta só em 127.0.0.1, então já não é alcançável de outra máquina.
O login existe para outra coisa: a máquina é compartilhada e a tela mostra
valor de fatura, CNPJ e dado bancário de fornecedor. Sem senha, qualquer um
que sente na cadeira vê tudo — e pior, pode despachar autorização.

Decisões que valem explicação:

* **A senha nunca é guardada.** Fica só o resumo `scrypt` dela, com sal
  próprio por usuário. Nem eu, nem quem abrir o `usuarios.yaml`, nem quem
  copiar o arquivo consegue recuperar a senha original.
* **`scrypt` da biblioteca padrão**, não bcrypt nem argon2. Os parâmetros
  (n=2^14, r=8, p=1) são os recomendados para uso interativo e levam ~100 ms
  aqui — lento o bastante para tentativa e erro não valer a pena, rápido o
  bastante para ninguém reclamar do login. E é stdlib: uma dependência a
  menos para instalar numa máquina nova.
* **Comparação em tempo constante** (`compare_digest`). Comparar com `==`
  vaza, pelo tempo de resposta, quantos bytes do resumo bateram.
* **A primeira senha obriga troca.** Ela é previsível por construção (foi
  combinada por escrito), então serve para entrar uma vez e mais nada.
"""

from __future__ import annotations

import base64
import hashlib
import hmac
import logging
import os
import re
import secrets
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

import yaml

from automacao.nucleo.config import DIR_CONFIG, RAIZ_PROJETO
from automacao.acesso.permissoes import GRUPO_ADMIN, GRUPO_PADRAO

log = logging.getLogger("automacao.acesso.usuarios")

CAMINHO_USUARIOS = DIR_CONFIG / "usuarios.yaml"
#: Segredo que assina o cookie de sessão. Fica fora do git, como o cofre.
CAMINHO_SEGREDO = RAIZ_PROJETO / "dados" / "sessao.chave"

#: Parâmetros do scrypt. Mudar aqui invalida os resumos já gravados — quem
#: mexer precisa obrigar todo mundo a trocar a senha.
_N, _R, _P = 2**14, 8, 1
_TAMANHO_SAL = 16
_TAMANHO_RESUMO = 32

#: Faixa de tamanho de uma senha nova. Não é política de banco: o mínimo é o
#: suficiente para não aceitar "123" num painel que despacha pagamento, e o
#: máximo é o teto combinado com quem opera.
TAMANHO_MINIMO_SENHA = 8
TAMANHO_MAXIMO_SENHA = 16


class ErroUsuario(RuntimeError):
    """Falha ao ler, gravar ou autenticar um usuário."""


@dataclass
class Usuario:
    """Uma pessoa com acesso ao painel."""

    email: str
    nome: str = ""
    #: Cargo e área de quem assina — as linhas 2 e 3 da assinatura do e-mail.
    #: Ficam aqui, e não no settings.yaml, porque são de pessoa, não do painel:
    #: duas pessoas usando a mesma máquina assinam cada uma com o seu.
    funcao: str = ""
    setor: str = ""
    #: Molde pessoal do corpo do e-mail. Vazio = vale o `email.corpo` do
    #: settings.yaml, que é o de todo mundo.
    corpo_email: str = ""
    #: `scrypt$<sal em hex>$<resumo em hex>`. Nunca a senha.
    senha: str = ""
    #: Enquanto True, qualquer tela redireciona para a troca de senha.
    precisa_trocar_senha: bool = True
    #: Chave do grupo de permissão (ver automacao/permissoes.py). Quem manda em
    #: quais telas a pessoa enxerga; 'administrador' vê tudo.
    grupo: str = GRUPO_PADRAO
    ativo: bool = True
    criado_em: str = ""
    ultimo_acesso: str = ""

    @property
    def rotulo(self) -> str:
        return self.nome or self.email

    @property
    def administrador(self) -> bool:
        """
        Derivado do grupo, não guardado ao lado dele.

        Enquanto eram dois campos independentes, dava para gravar
        `administrador: true` com `grupo: consulta` e ninguém saber qual dos
        dois valia. Agora só existe uma fonte.
        """
        return (self.grupo or "").strip().lower() == GRUPO_ADMIN

    def confere(self, senha: str) -> bool:
        return _confere(senha, self.senha)

    @classmethod
    def de_dict(cls, dados: dict) -> Usuario:
        dados = dict(dados or {})
        # Cadastro anterior aos grupos guardava só `administrador: true/false`.
        # Sem esta conversão, todo mundo cairia no grupo padrão de uma vez — os
        # administradores inclusive, que perderiam o próprio cadastro.
        if not str(dados.get("grupo") or "").strip():
            dados["grupo"] = GRUPO_ADMIN if dados.get("administrador") else GRUPO_PADRAO
        campos = set(cls.__dataclass_fields__)
        return cls(**{k: v for k, v in dados.items() if k in campos})

    def para_dict(self) -> dict:
        return {
            "email": self.email,
            "nome": self.nome,
            "funcao": self.funcao,
            "setor": self.setor,
            "corpo_email": self.corpo_email,
            "senha": self.senha,
            "precisa_trocar_senha": bool(self.precisa_trocar_senha),
            "grupo": self.grupo,
            "ativo": bool(self.ativo),
            "criado_em": self.criado_em,
            "ultimo_acesso": self.ultimo_acesso,
        }


# --------------------------------------------------------------------------- #
# Assinatura do e-mail
# --------------------------------------------------------------------------- #


def assinatura_de(usuario: Usuario | None, empresa: str = "") -> str:
    """
    O bloco de três linhas que fecha o e-mail da autorização.

        Fulana de Tal
        Analista de Sistemas
        TI | Nome da Empresa

    Campo vazio **some a linha**, não deixa linha em branco: buraco no meio de
    uma assinatura parece mensagem cortada, e um `" | "` solto (setor vazio,
    empresa preenchida) parece placeholder que ninguém substituiu. Por isso a
    terceira linha é montada juntando só o que existe.

    Args:
        usuario: quem está logado. `None` devolve string vazia — o corpo sai
            sem assinatura, que é melhor que sair com uma pela metade.
        empresa: nome da empresa, de `email.assinatura_empresa` no
            settings.yaml. É global: a empresa é a mesma para todo mundo.

    Returns:
        As linhas já unidas por `\\n`, sem quebra no fim. Vazio quando não há
        nada para assinar.
    """
    if usuario is None:
        return ""

    linhas: list[str] = []
    # `rotulo` cai no e-mail quando o nome não foi preenchido: assinar com o
    # endereço é feio, mas some menos que não assinar.
    nome = str(getattr(usuario, "rotulo", "") or "").strip()
    if nome:
        linhas.append(nome)

    funcao = str(getattr(usuario, "funcao", "") or "").strip()
    if funcao:
        linhas.append(funcao)

    setor = str(getattr(usuario, "setor", "") or "").strip()
    terceira = " | ".join(p for p in (setor, (empresa or "").strip()) if p)
    if terceira:
        linhas.append(terceira)

    return "\n".join(linhas)


def corpo_de(usuario: Usuario | None) -> str:
    """
    O molde pessoal do corpo, ou string vazia quando a pessoa não tem um.

    Existe para o painel não precisar saber que o campo se chama
    `corpo_email` nem lidar com usuário ausente.
    """
    return str(getattr(usuario, "corpo_email", "") or "")


# --------------------------------------------------------------------------- #
# Senha
# --------------------------------------------------------------------------- #


def resumir(senha: str) -> str:
    """Transforma a senha no que vai para o arquivo — e só isso volta de lá."""
    sal = secrets.token_bytes(_TAMANHO_SAL)
    bruto = hashlib.scrypt(
        senha.encode("utf-8"), salt=sal, n=_N, r=_R, p=_P, dklen=_TAMANHO_RESUMO
    )
    return f"scrypt${sal.hex()}${bruto.hex()}"


def _confere(senha: str, guardado: str) -> bool:
    """
    Compara sem revelar nada pelo caminho.

    Formato desconhecido devolve False em vez de levantar: um arquivo
    corrompido não pode virar porta aberta nem tela de erro.
    """
    try:
        algoritmo, sal_hex, esperado_hex = (guardado or "").split("$")
        if algoritmo != "scrypt":
            return False
        calculado = hashlib.scrypt(
            senha.encode("utf-8"),
            salt=bytes.fromhex(sal_hex),
            n=_N, r=_R, p=_P, dklen=_TAMANHO_RESUMO,
        )
    except (ValueError, AttributeError):
        return False
    return hmac.compare_digest(calculado, bytes.fromhex(esperado_hex))


#: Molde usado quando `config/dados-locais.yaml` não traz nenhum. Neutro de
#: propósito — o nome da empresa não pode voltar para dentro do código.
_SENHA_INICIAL_NEUTRA = "Trocar{ano}@"


def _modelo_da_senha_inicial() -> str:
    """
    O molde da senha inicial, lido de `config/dados-locais.yaml`.

    Já foi um literal escrito aqui dentro. Com o repositório público, esse
    literal entregava de graça o nome da empresa e a senha de primeiro acesso
    de todo usuário novo — e o `conferir_publicacao.py` não acusava, porque o
    padrão dele exigia um espaço que a grafia usada não tinha.

    Arquivo ausente ou ilegível cai no molde neutro em vez de levantar: um
    clone recém-feito precisa conseguir criar o primeiro usuário.
    """
    try:
        bruto = yaml.safe_load(
            (DIR_CONFIG / "dados-locais.yaml").read_text(encoding="utf-8")
        ) or {}
        modelo = str(bruto.get("senha_inicial") or "").strip()
    except (OSError, AttributeError, yaml.YAMLError):
        modelo = ""
    return modelo or _SENHA_INICIAL_NEUTRA


def senha_inicial(quando: date | None = None) -> str:
    """
    A senha combinada para o primeiro acesso.

    É previsível de propósito — foi acertada por escrito e serve para entrar
    uma vez. Quem usa é obrigado a trocar antes de ver qualquer tela. O molde
    vem de `config/dados-locais.yaml`, que não vai para o repositório.
    """
    ano = (quando or date.today()).year
    return _modelo_da_senha_inicial().replace("{ano}", str(ano))


def criticar_senha(senha: str, email: str = "") -> list[str]:
    """
    O que há de errado com uma senha nova. Lista vazia = pode usar.

    Regras curtas e explicáveis. Recusar a senha inicial é o ponto principal:
    sem isso, "trocar a senha" viraria digitar a mesma coisa de novo.
    """
    problemas: list[str] = []
    if len(senha) < TAMANHO_MINIMO_SENHA:
        problemas.append(f"precisa ter pelo menos {TAMANHO_MINIMO_SENHA} caracteres")
    if len(senha) > TAMANHO_MAXIMO_SENHA:
        problemas.append(f"não pode passar de {TAMANHO_MAXIMO_SENHA} caracteres")
    if not re.search(r"[A-Za-zÀ-ÿ]", senha):
        problemas.append("precisa ter ao menos uma letra")
    if not re.search(r"\d", senha):
        problemas.append("precisa ter ao menos um número")
    if senha.strip() != senha:
        problemas.append("não pode começar nem terminar com espaço")
    for ano in (date.today().year, date.today().year - 1, date.today().year + 1):
        if senha == senha_inicial(date(ano, 1, 1)):
            problemas.append("é a senha inicial — escolha outra, só sua")
            break
    if email and senha.lower() == email.split("@")[0].lower():
        problemas.append("não pode ser o seu próprio usuário")
    return problemas


# --------------------------------------------------------------------------- #
# Arquivo
# --------------------------------------------------------------------------- #


def _agora() -> str:
    return datetime.now().isoformat(timespec="seconds")


def normalizar_email(email: str) -> str:
    return (email or "").strip().lower()


def carregar() -> list[Usuario]:
    """Todos os usuários do arquivo. Arquivo ausente devolve lista vazia."""
    if not CAMINHO_USUARIOS.is_file():
        return []
    try:
        dados = yaml.safe_load(CAMINHO_USUARIOS.read_text(encoding="utf-8")) or {}
    except yaml.YAMLError as erro:
        raise ErroUsuario(f"{CAMINHO_USUARIOS.name} ilegível: {erro}") from erro
    return [Usuario.de_dict(u) for u in (dados.get("usuarios") or [])]


def gravar(usuarios: list[Usuario]) -> None:
    """Reescreve o arquivo inteiro, com o aviso no topo."""
    CAMINHO_USUARIOS.parent.mkdir(parents=True, exist_ok=True)
    cabecalho = (
        "# Quem pode abrir o painel.\n"
        "#\n"
        "# O campo 'senha' NÃO é a senha: é o resumo scrypt dela, com sal\n"
        "# próprio. Não dá para voltar atrás a partir daqui — esqueceu a\n"
        "# senha, um administrador redefine pela tela de Usuários.\n"
        "#\n"
        "# Este arquivo fica FORA do git.\n"
    )
    corpo = yaml.safe_dump(
        {"usuarios": [u.para_dict() for u in usuarios]},
        allow_unicode=True, sort_keys=False, width=4096,
    )
    CAMINHO_USUARIOS.write_text(cabecalho + corpo, encoding="utf-8")
    try:
        os.chmod(CAMINHO_USUARIOS, 0o600)
    except OSError:  # pragma: no cover — sistema de arquivos sem permissão POSIX
        pass


def por_email(email: str) -> Usuario | None:
    alvo = normalizar_email(email)
    return next((u for u in carregar() if normalizar_email(u.email) == alvo), None)


def salvar(usuario: Usuario) -> None:
    """Insere ou substitui, casando pelo e-mail."""
    alvo = normalizar_email(usuario.email)
    atuais = [u for u in carregar() if normalizar_email(u.email) != alvo]
    gravar(sorted(atuais + [usuario], key=lambda u: u.email.lower()))


def criar(
    email: str,
    *,
    nome: str = "",
    senha: str | None = None,
    grupo: str = GRUPO_PADRAO,
) -> Usuario:
    """
    Cadastra alguém com senha provisória, obrigado a trocar no primeiro acesso.

    Senha não informada usa a inicial do ano — a mesma combinada para o
    primeiro usuário. É previsível, e é por isso que a troca é obrigatória.
    """
    from automacao.acesso import permissoes

    email = normalizar_email(email)
    if not email or "@" not in email:
        raise ErroUsuario(f"e-mail inválido: {email!r}")
    if por_email(email):
        raise ErroUsuario(f"já existe usuário com o e-mail {email}")

    chave = permissoes.normalizar(grupo) or GRUPO_PADRAO
    if permissoes.por_chave(chave) is None:
        raise ErroUsuario(f"grupo de permissão inexistente: {grupo}")

    usuario = Usuario(
        email=email,
        nome=nome.strip(),
        senha=resumir(senha or senha_inicial()),
        precisa_trocar_senha=True,
        grupo=chave,
        criado_em=_agora(),
    )
    salvar(usuario)
    log.info("usuário %s cadastrado (grupo=%s)", email, chave)
    return usuario


def trocar_grupo(email: str, grupo: str) -> Usuario:
    """
    Muda o grupo de permissão de alguém.

    Recusa tirar o último administrador ativo: sem ninguém no grupo, o cadastro
    de usuários e o de permissões ficariam sem dono e não haveria por onde
    desfazer — teria que se editar o YAML na mão.
    """
    from automacao.acesso import permissoes

    alvo = por_email(email)
    if alvo is None:
        raise ErroUsuario(f"usuário não encontrado: {email}")

    chave = permissoes.normalizar(grupo)
    if permissoes.por_chave(chave) is None:
        raise ErroUsuario(f"grupo de permissão inexistente: {grupo}")
    if chave == alvo.grupo:
        return alvo

    if alvo.administrador and chave != GRUPO_ADMIN:
        outros = [
            u for u in carregar()
            if u.administrador and u.ativo
            and normalizar_email(u.email) != normalizar_email(email)
        ]
        if not outros:
            raise ErroUsuario(
                "este é o último administrador ativo — promova outra pessoa "
                "antes de rebaixá-lo, senão ninguém mais mexe em usuários e "
                "permissões."
            )

    alvo.grupo = chave
    salvar(alvo)
    log.warning("usuário %s passou para o grupo %s", alvo.email, chave)
    return alvo


def trocar_senha(email: str, nova: str) -> None:
    """Grava a senha nova e tira a obrigação de trocar."""
    usuario = por_email(email)
    if usuario is None:
        raise ErroUsuario(f"usuário não encontrado: {email}")
    problemas = criticar_senha(nova, email)
    if problemas:
        raise ErroUsuario("; ".join(problemas))
    usuario.senha = resumir(nova)
    usuario.precisa_trocar_senha = False
    salvar(usuario)
    log.info("usuário %s trocou a senha", email)


def redefinir_senha(email: str) -> str:
    """
    Devolve o usuário à senha inicial e volta a obrigar a troca.

    É o que um administrador faz quando alguém esquece a senha: como o resumo
    não tem volta, não existe "ver a senha atual" — só substituir.
    """
    usuario = por_email(email)
    if usuario is None:
        raise ErroUsuario(f"usuário não encontrado: {email}")
    provisoria = senha_inicial()
    usuario.senha = resumir(provisoria)
    usuario.precisa_trocar_senha = True
    salvar(usuario)
    log.warning("senha de %s redefinida para a provisória", email)
    return provisoria


def remover(email: str) -> None:
    alvo = normalizar_email(email)
    restantes = [u for u in carregar() if normalizar_email(u.email) != alvo]
    if len(restantes) == len(carregar()):
        raise ErroUsuario(f"usuário não encontrado: {email}")
    if not any(u.administrador and u.ativo for u in restantes):
        raise ErroUsuario(
            "este é o último administrador ativo — cadastre outro antes de "
            "removê-lo, senão ninguém mais entra no cadastro de usuários."
        )
    gravar(restantes)
    log.warning("usuário %s removido", email)


def autenticar(email: str, senha: str) -> Usuario | None:
    """
    Devolve o usuário quando e-mail e senha batem; `None` em qualquer falha.

    Uma mensagem só para os dois casos (e-mail errado, senha errada) é
    deliberado: dizer "este e-mail não existe" entrega quem tem acesso.
    """
    usuario = por_email(email)
    if usuario is None or not usuario.ativo:
        # Gasta o mesmo tempo de um scrypt mesmo sem usuário, para o relógio
        # não denunciar quais e-mails existem.
        _confere(senha, resumir("desperdicio"))
        log.warning("tentativa de acesso recusada para %r", normalizar_email(email))
        return None
    if not usuario.confere(senha):
        log.warning("senha incorreta para %s", usuario.email)
        return None
    usuario.ultimo_acesso = _agora()
    salvar(usuario)
    return usuario


# --------------------------------------------------------------------------- #
# Segredo do cookie
# --------------------------------------------------------------------------- #


def segredo_da_sessao() -> bytes:
    """
    Chave que assina o cookie. Criada na primeira vez e guardada em dados/.

    Apagar o arquivo desloga todo mundo — é o botão de pânico se alguém
    suspeitar que um cookie vazou.
    """
    if CAMINHO_SEGREDO.is_file():
        dados = CAMINHO_SEGREDO.read_bytes().strip()
        if len(dados) >= 32:
            return dados
    CAMINHO_SEGREDO.parent.mkdir(parents=True, exist_ok=True)
    novo = secrets.token_bytes(48)
    CAMINHO_SEGREDO.write_bytes(novo)
    try:
        os.chmod(CAMINHO_SEGREDO, 0o600)
    except OSError:  # pragma: no cover
        pass
    log.info("chave de sessão criada em %s", CAMINHO_SEGREDO)
    # Relê em vez de devolver `novo`: a leitura acima passa `.strip()`, e em
    # ~4,5% dos sorteios os 48 bytes começam ou terminam com um byte de espaço
    # (0x20, \t, \n, \r, \v, \f). Devolvendo o cru, a chave desta chamada não
    # seria a das seguintes — o cookie assinado logo depois de criar o arquivo
    # não conferia e o login devolvia a pessoa para a tela de entrar, com a
    # senha certa. Um em vinte primeiras subidas do painel.
    return CAMINHO_SEGREDO.read_bytes().strip()


def assinar(email: str) -> str:
    """
    Cookie: `<e-mail em base64url>.<assinatura>`.

    O e-mail vai codificado, e não em claro, por um motivo prático: `@` não é
    caractere simples de cookie, então o servidor envolve o valor em aspas ao
    enviar — e elas voltam junto na requisição seguinte, quebrando a
    conferência da assinatura. Custou uma sessão que não colava.

    Sem prazo próprio além do `max_age` do cookie: o painel é de mesa, e
    expirar no meio de um mês de faturas só atrapalharia.
    """
    cru = base64.urlsafe_b64encode(email.encode("utf-8")).decode("ascii").rstrip("=")
    marca = hmac.new(segredo_da_sessao(), cru.encode("ascii"), hashlib.sha256)
    return f"{cru}.{marca.hexdigest()}"


def de_cookie(bruto: str | None) -> Usuario | None:
    """O usuário de um cookie válido, ou `None`. Assinatura errada = None."""
    if not bruto:
        return None
    # Cliente antigo (ou servidor) pode devolver o valor entre aspas.
    bruto = bruto.strip().strip('"')
    if "." not in bruto:
        return None

    cru, _, assinatura = bruto.rpartition(".")
    esperado = hmac.new(
        segredo_da_sessao(), cru.encode("ascii", "ignore"), hashlib.sha256
    ).hexdigest()
    if not hmac.compare_digest(assinatura, esperado):
        log.warning("cookie de sessão com assinatura inválida")
        return None

    try:
        email = base64.urlsafe_b64decode(cru + "=" * (-len(cru) % 4)).decode("utf-8")
    except (ValueError, UnicodeDecodeError):
        return None
    usuario = por_email(email)
    return usuario if usuario and usuario.ativo else None


def garantir_primeiro_usuario(email: str, nome: str = "") -> Usuario | None:
    """
    Cadastra o primeiro administrador, se o arquivo ainda estiver vazio.

    Roda na subida do painel. Sem isso, a primeira execução mostraria uma
    tela de login sem ninguém para entrar.
    """
    if carregar():
        return None
    if not normalizar_email(email) or "@" not in email:
        log.error(
            "não há usuário cadastrado e 'acesso.primeiro_administrador' não "
            "está preenchido em settings.yaml — ninguém consegue entrar. "
            "Preencha com o e-mail de quem administra o painel."
        )
        return None
    usuario = criar(email, nome=nome, grupo=GRUPO_ADMIN)
    log.warning(
        "primeiro acesso: usuário %s criado com a senha inicial do ano. "
        "A troca é obrigatória no primeiro login.",
        email,
    )
    return usuario
