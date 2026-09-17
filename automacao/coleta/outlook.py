"""
Coleta de faturas direto do Outlook clássico (COM/MAPI).

Quatro cuidados que este módulo leva a sério:

* **Somente leitura.** Nada de marcar como lido, mover ou apagar mensagem.
  A única coisa que sai daqui é uma cópia do anexo na pasta de trabalho local.
* **Nunca varre a caixa inteira.** Sempre `Items.Restrict("[ReceivedTime] …")`
  com a janela de `outlook.janela_dias`, e `Items.Sort("[ReceivedTime]", True)`.
* **Conexão resiliente.** Tentamos, nesta ordem, `GetActiveObject` →
  `Dispatch` → `DispatchEx`. `GetActiveObject` só funciona quando o Outlook já
  terminou de subir e se registrou na ROT visível para este processo — na
  máquina de produção ele falha sempre com "Operação não disponível"
  (MK_E_UNAVAILABLE) e quem resolve é o `Dispatch`. `DispatchEx` fica de
  último recurso.
* **`Namespace.Logon()` é obrigatório.** Só pegar o namespace não basta: sem o
  `Logon` explícito, `CurrentUser`, `GetDefaultFolder()` e `Stores` devolvem
  "Falha na operação" / "Falha na chamada de procedimento remoto" e
  `Namespace.Folders` **pendura**. Com `Logon("", "", False, False)` — sem
  diálogo e sem sessão nova, ou seja, encaixando na sessão que já existe —
  tudo passa a responder.
* **Nada pendura o painel.** Este módulo roda dentro de um servidor web e um
  Outlook meio configurado trava o COM para sempre, esperando um assistente
  que ninguém vai responder. Por isso toda conversa com o MAPI acontece numa
  thread própria, com prazo, e antes disso conferimos no registro se existe
  perfil.
"""

from __future__ import annotations

import hashlib
import logging
import re
import shutil
import threading
from contextlib import contextmanager
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Callable, Iterator, TypeVar

from automacao.coleta.classificador import classificar
from automacao.nucleo.config import Ambiente, ambiente
from automacao.nucleo.modelos import (
    Competencia,
    Conta,
    Documento,
    Etapa,
    OrigemDocumento,
    ResultadoEtapa,
    Situacao,
    normalizar,
)

log = logging.getLogger("automacao.coleta.outlook")

__all__ = [
    "ErroOutlook",
    "buscar",
    "buscar_candidatos",
    "listar_pastas",
    "nome_livre",
    "copiar_sem_duplicar",
    "outlook_disponivel",
    "perfis_mapi",
    "esquecer_travamento",
]

T = TypeVar("T")

# Constantes do modelo de objetos do Outlook.
PASTA_CAIXA_DE_ENTRADA = 6      # olFolderInbox
ANEXO_POR_VALOR = 1             # olByValue
ITEM_DE_EMAIL = "IPM.Note"

# Teto de segurança: mesmo com Restrict, não processamos mais que isto por pasta.
MAX_MENSAGENS_POR_PASTA = 400
# Profundidade máxima ao listar as pastas para o painel.
PROFUNDIDADE_MAXIMA = 4
# Prazos (segundos). Um "ping" de diagnóstico tem que responder rápido; uma
# varredura de caixa postal com várias pastas no Exchange demora mesmo.
TEMPO_LIMITE_DIAGNOSTICO = 30.0
TEMPO_LIMITE_VARREDURA = 300.0
# Depois de pendurar uma vez, ficamos este tempo recusando na hora, para não
# acumular threads presas no COM a cada clique do painel.
CARENCIA_APOS_TRAVAMENTO = 120.0

# Onde o Office guarda os perfis MAPI. Varia por versão instalada.
CHAVES_PERFIS = (
    r"Software\Microsoft\Office\16.0\Outlook\Profiles",
    r"Software\Microsoft\Office\15.0\Outlook\Profiles",
    r"Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles",
)

MENSAGEM_SEM_PERFIL = (
    "O Outlook está instalado e responde, mas NÃO existe nenhum perfil de e-mail "
    "(perfil MAPI) configurado nesta conta do Windows. Enquanto isso não for "
    "resolvido, qualquer leitura de caixa postal trava esperando o assistente "
    "'Bem-vindo ao Microsoft Outlook'. Abra o Outlook clássico na mão, conclua o "
    "assistente adicionando a conta de e-mail, deixe-o sincronizar e tente de novo."
)

MENSAGEM_SEM_OUTLOOK = (
    "Não foi possível falar com o Outlook. Confira se o Outlook clássico (desktop) "
    "está aberto na MESMA sessão de usuário do Windows em que o painel roda — se um "
    "dos dois foi iniciado como administrador e o outro não, o Windows bloqueia a "
    "conversa entre eles e o erro aparece como 'Falha na execução do servidor'. "
    "Abra o Outlook normalmente (sem 'Executar como administrador'), espere ele "
    "terminar de carregar e tente de novo."
)

MENSAGEM_TRAVOU = (
    "O Outlook aceitou a conexão mas não terminou '{operacao}' em {prazo:.0f}s. "
    "As duas causas comuns são: (1) um diálogo esperando resposta na tela do "
    "Outlook — assistente de perfil, pedido de senha, aviso de segurança; ou "
    "(2) caixa postal grande demais para a janela pedida. Confira o Outlook e, "
    "se não houver diálogo aberto, reduza 'outlook.janela_dias' ou aponte "
    "'outlook.pastas' para a pasta certa em config/settings.yaml."
)


class ErroOutlook(RuntimeError):
    """Não foi possível conversar com o Outlook."""


# Lembrete de que a última tentativa pendurou (evita empilhar threads presas).
_travamento: dict[str, Any] = {"quando": None, "detalhe": ""}
# Qual estratégia funcionou da última vez — o painel mostra isso no diagnóstico.
_ultima_estrategia: dict[str, str] = {"nome": ""}


def esquecer_travamento() -> None:
    """Zera a carência depois de um travamento — o painel chama ao 'tentar de novo'."""
    _travamento["quando"] = None
    _travamento["detalhe"] = ""


# --------------------------------------------------------------------------- #
# Pré-voo: existe perfil MAPI?
# --------------------------------------------------------------------------- #


def perfis_mapi() -> list[str]:
    """
    Nomes dos perfis de e-mail configurados nesta conta do Windows.

    Consulta só o registro (leitura), sem acordar o Outlook. Lista vazia
    significa que o Outlook nunca foi configurado — e que chamar
    `GetNamespace("MAPI")` vai travar.
    """
    try:
        import winreg
    except ImportError:  # pragma: no cover - só existe no Windows
        return []

    encontrados: list[str] = []
    for chave in CHAVES_PERFIS:
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, chave) as raiz:
                indice = 0
                while True:
                    try:
                        nome = winreg.EnumKey(raiz, indice)
                    except OSError:
                        break
                    if nome not in encontrados:
                        encontrados.append(nome)
                    indice += 1
        except OSError:
            continue
    return encontrados


# --------------------------------------------------------------------------- #
# Conexão
# --------------------------------------------------------------------------- #


@contextmanager
def _com_inicializado() -> Iterator[None]:
    """
    Inicializa o COM na thread atual e devolve tudo no fim.

    Se a thread já estava inicializada (o servidor pode ter feito isso),
    respeitamos e não desinicializamos no final.
    """
    import pythoncom

    inicializamos = False
    try:
        pythoncom.CoInitialize()
        inicializamos = True
    except Exception as exc:  # já inicializado noutro modo de apartamento
        log.debug("CoInitialize dispensado nesta thread: %s", exc)
    try:
        yield
    finally:
        if inicializamos:
            try:
                pythoncom.CoUninitialize()
            except Exception as exc:  # pragma: no cover
                log.debug("CoUninitialize falhou: %s", exc)


def _descrever_erro(exc: BaseException) -> str:
    """Mensagem curta e legível de um erro COM."""
    texto = str(exc).strip()
    codigo = getattr(exc, "hresult", None)
    if codigo and str(codigo) not in texto:
        texto = f"{texto} [hresult {codigo}]"
    return f"{type(exc).__name__}: {texto}" if texto else type(exc).__name__


def _abrir_aplicacao() -> tuple[Any, str, list[str]]:
    """
    Abre o objeto `Outlook.Application`, tentando as três estratégias.

    Só o objeto da aplicação — pegar o namespace MAPI é o passo seguinte, e é
    ele que pode travar.

    Returns:
        `(aplicacao, estrategia_que_funcionou, relato_das_tentativas)`.

    Raises:
        ErroOutlook: quando as três falham.
    """
    import win32com.client

    estrategias = (
        ("GetActiveObject", lambda: win32com.client.GetActiveObject("Outlook.Application")),
        ("Dispatch", lambda: win32com.client.Dispatch("Outlook.Application")),
        ("DispatchEx", lambda: win32com.client.DispatchEx("Outlook.Application")),
    )

    tentativas: list[str] = []
    for nome, abrir in estrategias:
        try:
            aplicacao = abrir()
            versao = str(aplicacao.Version)
        except Exception as exc:
            detalhe = _descrever_erro(exc)
            tentativas.append(f"{nome} → {detalhe}")
            log.info("estratégia %s falhou: %s", nome, detalhe)
            continue
        tentativas.append(f"{nome} → OK (Outlook {versao})")
        _ultima_estrategia["nome"] = nome
        log.info("conectado ao Outlook %s via %s", versao, nome)
        return aplicacao, nome, tentativas

    raise ErroOutlook(f"{MENSAGEM_SEM_OUTLOOK} Tentativas: {' | '.join(tentativas)}")


def _abrir_namespace(aplicacao: Any) -> Any:
    """
    Devolve o namespace MAPI já logado.

    O `Logon` não cria perfil nem sessão nova: com `ShowDialog=False` e
    `NewSession=False` ele apenas encaixa na sessão que o Outlook já mantém
    aberta. Sem essa chamada, `GetDefaultFolder()` responde "Falha na operação"
    e `Namespace.Folders` pendura — foi exatamente o que acontecia na máquina
    de produção.
    """
    namespace = aplicacao.GetNamespace("MAPI")
    try:
        namespace.Logon("", "", False, False)
    except Exception as exc:
        # Perfis já logados às vezes recusam um segundo Logon; se o namespace
        # estiver utilizável, seguimos assim mesmo.
        log.debug("Logon recusado (talvez já estivesse logado): %s", _descrever_erro(exc))
    return namespace


def _tempo_limite(chave: str, padrao: float) -> float:
    """Prazo configurável em `outlook.<chave>` do settings.yaml."""
    try:
        bruto = ambiente().outlook.get(chave)
        return float(bruto) if bruto else padrao
    except Exception:
        return padrao


def _em_carencia() -> str:
    """Explicação se ainda estamos no período de carência pós-travamento."""
    quando = _travamento["quando"]
    if quando is None:
        return ""
    faltam = CARENCIA_APOS_TRAVAMENTO - (datetime.now() - quando).total_seconds()
    if faltam <= 0:
        esquecer_travamento()
        return ""
    return (
        f"{_travamento['detalhe']} Uma tentativa anterior ficou presa e a thread "
        f"ainda está lá; esperando {faltam:.0f}s antes de tentar de novo."
    )


def _executar_no_outlook(
    tarefa: Callable[[Any, str], T],
    *,
    tempo_limite: float | None = None,
    operacao: str = "conversar com o Outlook",
) -> T:
    """
    Roda `tarefa(namespace, estrategia)` numa thread própria, com prazo.

    A thread inicializa o próprio COM e devolve **apenas dados Python** — nenhum
    objeto COM atravessa a fronteira da thread, que seria um erro de apartamento.
    Se estourar o prazo, a thread fica lá (não dá para matar thread presa no
    COM), mas ela é daemon e o painel volta na hora com um erro em português.

    Args:
        tarefa: função que recebe o namespace MAPI logado e a estratégia usada.
        tempo_limite: prazo em segundos.
        operacao: nome da operação, para a mensagem de erro fazer sentido.

    Raises:
        ErroOutlook: sem Outlook, sem perfil MAPI ou prazo estourado.
    """
    prazo = tempo_limite or TEMPO_LIMITE_DIAGNOSTICO

    pendencia = _em_carencia()
    if pendencia:
        raise ErroOutlook(pendencia)

    if not perfis_mapi():
        raise ErroOutlook(MENSAGEM_SEM_PERFIL)

    caixa: dict[str, Any] = {}

    def alvo() -> None:
        with _com_inicializado():
            try:
                aplicacao, estrategia, _ = _abrir_aplicacao()
                namespace = _abrir_namespace(aplicacao)
                caixa["valor"] = tarefa(namespace, estrategia)
            except BaseException as exc:  # noqa: BLE001 - repassado para a thread chamadora
                caixa["erro"] = exc

    thread = threading.Thread(target=alvo, name="automacao-outlook-mapi", daemon=True)
    thread.start()
    thread.join(prazo)

    if thread.is_alive():
        _travamento["quando"] = datetime.now()
        _travamento["detalhe"] = MENSAGEM_TRAVOU.format(prazo=prazo, operacao=operacao)
        log.error("'%s' estourou o prazo de %.0fs", operacao, prazo)
        raise ErroOutlook(_travamento["detalhe"])

    if "erro" in caixa:
        erro = caixa["erro"]
        if isinstance(erro, ErroOutlook):
            raise erro
        raise ErroOutlook(f"{MENSAGEM_SEM_OUTLOOK} Detalhe técnico: {_descrever_erro(erro)}")

    return caixa["valor"]


def outlook_disponivel() -> tuple[bool, str]:
    """
    Diagnóstico para o painel — nunca levanta exceção, nunca pendura.

    Returns:
        `(True, "conectado via <estratégia> — perfil <nome>")` ou
        `(False, explicação em português do que o usuário precisa fazer)`.
    """
    def tarefa(namespace: Any, estrategia: str) -> str:
        try:
            usuario = str(namespace.CurrentUser.Name)
        except Exception:
            usuario = "(usuário não informado)"
        caixa = str(namespace.GetDefaultFolder(PASTA_CAIXA_DE_ENTRADA).Name)
        return f"conectado via {estrategia} — usuário {usuario}, caixa '{caixa}'"

    try:
        return True, _executar_no_outlook(
            tarefa,
            tempo_limite=_tempo_limite("tempo_limite_s", TEMPO_LIMITE_DIAGNOSTICO),
            operacao="abrir a caixa postal",
        )
    except ErroOutlook as exc:
        return False, str(exc)
    except Exception as exc:  # pragma: no cover - defensivo
        return False, f"{MENSAGEM_SEM_OUTLOOK} Detalhe técnico: {_descrever_erro(exc)}"


# --------------------------------------------------------------------------- #
# Navegação de pastas
# --------------------------------------------------------------------------- #


def _subpastas(pasta: Any, profundidade: int = 0) -> Iterator[tuple[str, Any]]:
    """Percorre a pasta e suas filhas, devolvendo `(caminho_legivel, pasta)`."""
    try:
        nome = str(pasta.Name)
    except Exception:
        return
    yield nome, pasta
    if profundidade >= PROFUNDIDADE_MAXIMA:
        return
    try:
        filhas = list(pasta.Folders)
    except Exception as exc:
        log.debug("não consegui listar subpastas de %s: %s", nome, _descrever_erro(exc))
        return
    for filha in filhas:
        for caminho, encontrada in _subpastas(filha, profundidade + 1):
            yield f"{nome}\\{caminho}", encontrada


def listar_pastas(*, tempo_limite: float | None = None) -> list[str]:
    """
    Caminhos de todas as pastas visíveis no perfil, para o painel montar um combo.

    Ex.: `['Caixa de Entrada', 'Caixa de Entrada\\\\Faturas', 'Itens Enviados', …]`

    Args:
        tempo_limite: prazo em segundos. O padrão é folgado porque um perfil
            com várias caixas compartilhadas e arquivos morto demora — cada
            store é uma ida ao servidor.

    Raises:
        ErroOutlook: se o Outlook não responder dentro do prazo.
    """
    def tarefa(namespace: Any, _estrategia: str) -> list[str]:
        caminhos: list[str] = []
        for raiz in namespace.Folders:
            for caminho, _pasta in _subpastas(raiz):
                caminhos.append(caminho)
        return caminhos

    return _executar_no_outlook(
        tarefa,
        tempo_limite=tempo_limite or TEMPO_LIMITE_VARREDURA,
        operacao="listar as pastas",
    )


def _resolver_pastas(namespace: Any, nomes_configurados: list[str]) -> list[tuple[str, Any]]:
    """
    Pastas que serão varridas.

    Lista vazia na configuração = Caixa de Entrada e suas subpastas.
    """
    if not nomes_configurados:
        entrada = namespace.GetDefaultFolder(PASTA_CAIXA_DE_ENTRADA)
        return list(_subpastas(entrada))

    desejadas = {normalizar(n) for n in nomes_configurados}
    encontradas: list[tuple[str, Any]] = []
    for raiz in namespace.Folders:
        for caminho, pasta in _subpastas(raiz):
            ultimo = caminho.rsplit("\\", 1)[-1]
            if normalizar(caminho) in desejadas or normalizar(ultimo) in desejadas:
                encontradas.append((caminho, pasta))
    if not encontradas:
        log.warning(
            "nenhuma pasta do Outlook bateu com %s — caindo para a Caixa de Entrada",
            nomes_configurados,
        )
        entrada = namespace.GetDefaultFolder(PASTA_CAIXA_DE_ENTRADA)
        return list(_subpastas(entrada))
    return encontradas


# --------------------------------------------------------------------------- #
# Janela de datas e Restrict
# --------------------------------------------------------------------------- #


def _intervalo(competencia: Competencia, janela_dias: int) -> tuple[datetime, datetime]:
    """
    Período de recebimento que interessa para uma competência.

    A fatura de julho costuma chegar entre o fim de junho e o meio de agosto,
    então abrimos `janela_dias` antes do dia 1º e `janela_dias` depois do
    último dia do mês.
    """
    primeiro = datetime(competencia.ano, competencia.mes, 1)
    proximo = competencia.proxima()
    ultimo = datetime(proximo.ano, proximo.mes, 1) - timedelta(seconds=1)
    return primeiro - timedelta(days=janela_dias), ultimo + timedelta(days=janela_dias)


def _formatos_de_filtro(inicio: datetime, fim: datetime) -> list[str]:
    """
    Filtros `Restrict` candidatos.

    O Outlook interpreta a data conforme as configurações regionais da máquina;
    mandamos o formato americano primeiro (o mais aceito) e o brasileiro como
    reserva.
    """
    moldes = ("%m/%d/%Y %I:%M %p", "%d/%m/%Y %H:%M")
    return [
        f"[ReceivedTime] >= '{inicio.strftime(m)}' AND [ReceivedTime] <= '{fim.strftime(m)}'"
        for m in moldes
    ]


def _itens_da_janela(pasta: Any, inicio: datetime, fim: datetime) -> list[Any]:
    """
    Mensagens da pasta dentro da janela, mais novas primeiro.

    Nunca itera a coleção inteira: primeiro ordena, depois `Restrict`.
    """
    try:
        itens = pasta.Items
        itens.Sort("[ReceivedTime]", True)
    except Exception as exc:
        log.debug("pasta sem itens utilizáveis: %s", _descrever_erro(exc))
        return []

    ultimo_erro = ""
    for filtro in _formatos_de_filtro(inicio, fim):
        try:
            restrito = itens.Restrict(filtro)
            colhidos: list[Any] = []
            for item in restrito:
                colhidos.append(item)
                if len(colhidos) >= MAX_MENSAGENS_POR_PASTA:
                    log.warning(
                        "pasta atingiu o teto de %d mensagens — reduza outlook.janela_dias",
                        MAX_MENSAGENS_POR_PASTA,
                    )
                    break
            return colhidos
        except Exception as exc:
            ultimo_erro = _descrever_erro(exc)
            log.debug("filtro %r recusado: %s", filtro, ultimo_erro)
    log.warning("nenhum formato de data foi aceito pelo Restrict: %s", ultimo_erro)
    return []


# --------------------------------------------------------------------------- #
# Leitura das mensagens (tudo somente leitura)
# --------------------------------------------------------------------------- #


def _texto(item: Any, propriedade: str, limite: int | None = None) -> str:
    try:
        valor = getattr(item, propriedade, "") or ""
    except Exception:
        return ""
    valor = str(valor)
    return valor[:limite] if limite else valor


def _recebido_em(item: Any) -> datetime | None:
    try:
        bruto = item.ReceivedTime
        return datetime(bruto.year, bruto.month, bruto.day, bruto.hour, bruto.minute, bruto.second)
    except Exception:
        return None


def _remetente_barato(item: Any) -> str:
    """
    Quem enviou, sem ida ao servidor.

    Serve para o primeiro filtro de uma varredura: resolver o SMTP real de
    cada mensagem custa uma consulta ao Exchange, e numa caixa cheia isso
    sozinho estoura qualquer prazo.
    """
    endereco = _texto(item, "SenderEmailAddress")
    nome = _texto(item, "SenderName")
    return f"{endereco} {nome}".strip()


def _remetente(item: Any) -> str:
    """
    Endereço de quem enviou.

    Contas do Exchange devolvem um caminho X500 em `SenderEmailAddress`; nesse
    caso tentamos o endereço SMTP de verdade antes de desistir. Custa uma ida
    ao servidor — use `_remetente_barato()` quando for só para filtrar.
    """
    endereco = _texto(item, "SenderEmailAddress")
    if endereco and not endereco.upper().startswith("/O="):
        return endereco
    try:
        remetente = item.Sender
        if remetente is not None:
            usuario = remetente.GetExchangeUser()
            if usuario is not None and usuario.PrimarySmtpAddress:
                return str(usuario.PrimarySmtpAddress)
    except Exception:
        pass
    return endereco or _texto(item, "SenderName")


def _nomes_de_anexos(item: Any) -> list[str]:
    try:
        return [str(a.FileName) for a in item.Attachments]
    except Exception:
        return []


def _e_email(item: Any) -> bool:
    try:
        return str(item.MessageClass or "").startswith(ITEM_DE_EMAIL)
    except Exception:
        return False


# --------------------------------------------------------------------------- #
# Casamento conta × mensagem
# --------------------------------------------------------------------------- #


def _contem_algum(alvo: str, termos: list[str]) -> list[str]:
    """Termos da lista que aparecem em `alvo` (`alvo` já normalizado)."""
    return [t for t in termos if t and normalizar(t) in alvo]


def _so_digitos(texto: str) -> str:
    return re.sub(r"\D", "", texto)


def _identificador_bate(identificador: str, alvo: str, alvo_digitos: str) -> bool:
    """
    Um identificador casa por texto ou só pelos dígitos.

    Número de linha aparece escrito de mil jeitos — "(00) 0000-0000",
    "0000000000", "00 0000 0000". Comparar só os dígitos resolve, e é o que
    separa entre si as várias contas de uma mesma operadora.
    """
    if normalizar(identificador) in alvo:
        return True
    digitos = _so_digitos(identificador)
    return bool(digitos) and len(digitos) >= 5 and digitos in alvo_digitos


def _avaliar_mensagem(item: Any, conta: Conta) -> tuple[bool, list[str]]:
    """
    Decide se a mensagem é desta conta.

    Regra: precisa passar em todos os critérios que a conta declarou. Critério
    não declarado não reprova ninguém.

    Returns:
        `(casou, motivos)` — `motivos` vai para o painel explicar a escolha.
    """
    motivos: list[str] = []

    # Da propriedade mais barata para a mais cara: remetente e assunto são
    # locais; o corpo da mensagem pode custar uma ida ao Exchange, então só é
    # lido se a conta realmente usa identificadores.
    if conta.coleta.remetentes:
        remetente_barato = _remetente_barato(item)
        casados = _contem_algum(normalizar(remetente_barato), conta.coleta.remetentes)
        if not casados:
            return False, []
        motivos.append(f"remetente {remetente_barato.strip()} bate com {casados[0]!r}")

    if conta.coleta.assunto_contem:
        casados = _contem_algum(normalizar(_texto(item, "Subject", 500)), conta.coleta.assunto_contem)
        if not casados:
            return False, []
        motivos.append(f"assunto contém {casados[0]!r}")

    if conta.coleta.identificadores:
        caldo = normalizar(
            f"{_texto(item, 'Subject', 500)} {_texto(item, 'Body', 8000)} "
            f"{' '.join(_nomes_de_anexos(item))}"
        )
        caldo_digitos = _so_digitos(caldo)
        casados = [
            i for i in conta.coleta.identificadores
            if _identificador_bate(i, caldo, caldo_digitos)
        ]
        if not casados:
            return False, []
        motivos.append(f"identificador {casados[0]!r} encontrado na mensagem")

    return True, motivos


# --------------------------------------------------------------------------- #
# Anexos
# --------------------------------------------------------------------------- #


def nome_livre(pasta: Path, nome: str) -> Path:
    """
    Caminho que ainda não existe: `fatura.pdf`, `fatura_2.pdf`, `fatura_3.pdf`…

    Nunca sobrescreve: metade dos fornecedores manda um anexo chamado
    "boleto.pdf" e os dois importam. Usado também pela coleta manual — a regra
    de não perder arquivo vale para as duas origens.
    """
    limpo = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", nome).strip() or "anexo"
    destino = pasta / limpo
    if not destino.exists():
        return destino
    base, extensao = destino.stem, destino.suffix
    contador = 2
    while True:
        candidato = pasta / f"{base}_{contador}{extensao}"
        if not candidato.exists():
            return candidato
        contador += 1


def _impressao_digital(caminho: Path) -> str:
    """Identidade do arquivo pelo conteúdo."""
    h = hashlib.sha256()
    with caminho.open("rb") as fh:
        for bloco in iter(lambda: fh.read(1 << 18), b""):
            h.update(bloco)
    return h.hexdigest()


def copiar_sem_duplicar(origem: Path, pasta: Path) -> tuple[Path, bool]:
    """
    Copia para `pasta`, mas reaproveita o arquivo se ele já estiver lá.

    Devolve `(caminho, copiou)`.

    A mesma fatura chega por dois caminhos — anexo do Outlook e PDF que você
    baixou e largou em `dados/entrada/`. Usando só `nome_livre`, a segunda
    cópia virava `boleto_2.pdf` e o PDF final saía com a página repetida.
    Aqui, arquivo byte a byte idêntico não é copiado de novo; nome igual com
    conteúdo diferente continua ganhando sufixo, para não perder nada.
    """
    pasta.mkdir(parents=True, exist_ok=True)
    try:
        digital = _impressao_digital(origem)
    except OSError:
        return shutil.copy2(origem, nome_livre(pasta, origem.name)), True

    tamanho = origem.stat().st_size
    for existente in pasta.iterdir():
        if not existente.is_file() or existente.stat().st_size != tamanho:
            continue
        try:
            if _impressao_digital(existente) == digital:
                log.debug("%s já está em %s — reaproveitando", origem.name, pasta)
                return existente, False
        except OSError:
            continue

    alvo = nome_livre(pasta, origem.name)
    shutil.copy2(origem, alvo)
    return alvo, True


def _salvar_anexos(item: Any, destino: Path, extensoes: set[str]) -> list[Path]:
    """Grava na pasta de trabalho os anexos com extensão aceita. Não altera o e-mail."""
    salvos: list[Path] = []
    try:
        anexos = list(item.Attachments)
    except Exception as exc:
        log.debug("mensagem sem anexos acessíveis: %s", _descrever_erro(exc))
        return salvos

    for anexo in anexos:
        try:
            nome = str(anexo.FileName)
        except Exception:
            continue
        if Path(nome).suffix.lower() not in extensoes:
            continue
        try:
            if int(getattr(anexo, "Type", ANEXO_POR_VALOR)) != ANEXO_POR_VALOR:
                continue  # imagem embutida na assinatura, não é anexo de verdade
        except Exception:
            pass
        caminho = nome_livre(destino, nome)
        try:
            anexo.SaveAsFile(str(caminho))
        except Exception as exc:
            log.warning("não consegui salvar o anexo %s: %s", nome, _descrever_erro(exc))
            continue
        salvos.append(caminho)
        log.info("anexo salvo: %s", caminho.name)
    return salvos


# --------------------------------------------------------------------------- #
# API pública
# --------------------------------------------------------------------------- #


def buscar(
    conta: Conta,
    competencia: Competencia,
    *,
    ambiente_: Ambiente | None = None,
) -> ResultadoEtapa:
    """
    Procura no Outlook as faturas desta conta e baixa os anexos.

    Args:
        conta: conta a processar — usa `conta.coleta` para casar as mensagens.
        competencia: mês de referência.
        ambiente_: ambiente já carregado (útil em teste); padrão é `ambiente()`.

    Returns:
        `ResultadoEtapa` de `Etapa.COLETA`. Em `detalhes["documentos"]` vêm os
        `Documento` já classificados; em `detalhes["mensagens"]`, o resumo dos
        e-mails que casaram, para o painel mostrar de onde veio cada arquivo.
        Nunca levanta exceção: problema de Outlook vira `Situacao.ERRO`.
    """
    amb = ambiente_ or ambiente()
    etapa = Etapa.COLETA

    if not (conta.coleta.remetentes or conta.coleta.assunto_contem or conta.coleta.identificadores):
        return ResultadoEtapa.pulado(
            etapa,
            f"a conta {conta.rotulo!r} ainda não tem regras de coleta configuradas "
            f"(bloco 'coleta:' vazio em fornecedores.yaml). Sem remetente, assunto ou "
            f"identificador não dá para saber qual e-mail é dela — use a busca livre "
            f"do painel ou a entrada manual.",
            detalhes={"documentos": [], "mensagens": []},
        )

    janela_dias = int(amb.outlook.get("janela_dias", 45) or 45)
    extensoes = {str(e).lower() for e in (amb.outlook.get("extensoes_anexo") or [".pdf", ".xml"])}
    pastas_configuradas = list(amb.outlook.get("pastas") or [])
    inicio, fim = _intervalo(competencia, janela_dias)
    destino = amb.pasta_trabalho(conta.id, competencia)

    def tarefa(namespace: Any, estrategia: str) -> dict:
        documentos: list[Documento] = []
        mensagens: list[dict] = []
        artefatos: list[Path] = []
        pastas = _resolver_pastas(namespace, pastas_configuradas)

        for caminho_pasta, pasta in pastas:
            for item in _itens_da_janela(pasta, inicio, fim):
                if not _e_email(item):
                    continue
                casou, motivos = _avaliar_mensagem(item, conta)
                if not casou:
                    continue

                assunto = _texto(item, "Subject", 300)
                remetente = _remetente(item)
                recebido = _recebido_em(item)
                salvos = _salvar_anexos(item, destino, extensoes)

                for arquivo in salvos:
                    documento = classificar(arquivo, contexto=assunto)
                    documento.origem = OrigemDocumento.OUTLOOK
                    documento.remetente = remetente
                    documento.assunto_email = assunto
                    documentos.append(documento)
                    artefatos.append(arquivo)

                mensagens.append({
                    "pasta": caminho_pasta,
                    "assunto": assunto,
                    "remetente": remetente,
                    "recebido_em": recebido.isoformat() if recebido else None,
                    "anexos_salvos": [p.name for p in salvos],
                    "por_que_casou": motivos,
                })

        return {
            "documentos": documentos,
            "mensagens": mensagens,
            "artefatos": artefatos,
            "estrategia_conexao": estrategia,
            "pastas_varridas": [c for c, _ in pastas],
        }

    try:
        colheita = _executar_no_outlook(
            tarefa,
            tempo_limite=_tempo_limite("tempo_limite_varredura_s", TEMPO_LIMITE_VARREDURA),
            operacao=f"procurar os e-mails de {conta.rotulo}",
        )
    except ErroOutlook as exc:
        return ResultadoEtapa.erro(etapa, str(exc), detalhes={"documentos": [], "mensagens": []})

    documentos: list[Documento] = colheita["documentos"]
    mensagens: list[dict] = colheita["mensagens"]
    artefatos: list[Path] = colheita["artefatos"]
    detalhes = {
        "documentos": documentos,
        "mensagens": mensagens,
        "estrategia_conexao": colheita["estrategia_conexao"],
        "pastas_varridas": colheita["pastas_varridas"],
        "janela": {"de": inicio.isoformat(), "ate": fim.isoformat()},
    }

    if not mensagens:
        return ResultadoEtapa(
            etapa,
            Situacao.ATENCAO,
            f"nenhum e-mail de {conta.rotulo} entre {inicio:%d/%m/%Y} e {fim:%d/%m/%Y}. "
            f"Tente a busca livre do painel ou jogue o PDF em dados/entrada/.",
            detalhes=detalhes,
        )
    if not documentos:
        return ResultadoEtapa(
            etapa,
            Situacao.ATENCAO,
            f"{len(mensagens)} e-mail(s) casaram com {conta.rotulo}, mas nenhum trazia "
            f"anexo {'/'.join(sorted(extensoes))}.",
            detalhes=detalhes,
        )

    duvidosos = [d for d in documentos if d.confianca < 0.65]
    mensagem = (
        f"{len(documentos)} arquivo(s) coletado(s) de {len(mensagens)} e-mail(s) em {destino}"
    )
    if duvidosos:
        return ResultadoEtapa.atencao(
            etapa,
            f"{mensagem}. {len(duvidosos)} precisa(m) de confirmação do tipo: "
            + ", ".join(d.nome for d in duvidosos),
            detalhes=detalhes,
            artefatos=artefatos,
        )
    return ResultadoEtapa.sucesso(etapa, mensagem, detalhes=detalhes, artefatos=artefatos)


def buscar_candidatos(competencia: Competencia, *, termos: list[str]) -> list[dict]:
    """
    Busca livre, para o painel oferecer "achei estes e-mails, é algum destes?".

    Não baixa anexo nenhum: só lista o que existe. É a saída quando `buscar()`
    não encontrou nada e o humano precisa apontar a mensagem certa.

    Args:
        competencia: mês de referência (define a janela de datas).
        termos: pedaços de texto procurados no remetente, no assunto e nos nomes
            dos anexos. Qualquer termo que casar já traz a mensagem. Lista vazia
            devolve todos os e-mails com anexo na janela.

    Returns:
        Lista ordenada da mensagem mais nova para a mais antiga, com `pasta`,
        `assunto`, `remetente`, `recebido_em`, `anexos`, `entry_id` e
        `termos_encontrados`.

    Raises:
        ErroOutlook: se o Outlook não responder dentro do prazo.
    """
    amb = ambiente()
    janela_dias = int(amb.outlook.get("janela_dias", 45) or 45)
    pastas_configuradas = list(amb.outlook.get("pastas") or [])
    inicio, fim = _intervalo(competencia, janela_dias)
    procurados = [normalizar(t) for t in termos if t and t.strip()]

    def tarefa(namespace: Any, _estrategia: str) -> list[dict]:
        achados: list[dict] = []
        for caminho_pasta, pasta in _resolver_pastas(namespace, pastas_configuradas):
            for item in _itens_da_janela(pasta, inicio, fim):
                if not _e_email(item):
                    continue
                assunto = _texto(item, "Subject", 300)
                anexos = _nomes_de_anexos(item)
                # Filtra com o remetente barato; só resolve o SMTP de verdade
                # para as mensagens que sobreviverem ao filtro.
                alvo = normalizar(f"{assunto} {_remetente_barato(item)} {' '.join(anexos)}")
                alvo_digitos = _so_digitos(alvo)

                if procurados:
                    casados = [
                        t for t in procurados
                        if t in alvo or _identificador_bate(t, alvo, alvo_digitos)
                    ]
                    if not casados:
                        continue
                else:
                    casados = []
                    if not anexos:
                        continue  # sem termo, só interessa quem tem anexo

                recebido = _recebido_em(item)
                achados.append({
                    "pasta": caminho_pasta,
                    "assunto": assunto,
                    "remetente": _remetente(item),
                    "recebido_em": recebido.isoformat() if recebido else None,
                    "anexos": anexos,
                    "entry_id": _texto(item, "EntryID"),
                    "termos_encontrados": casados,
                })
        return achados

    achados = _executar_no_outlook(
        tarefa,
        tempo_limite=_tempo_limite("tempo_limite_varredura_s", TEMPO_LIMITE_VARREDURA),
        operacao="varrer as pastas em busca de candidatos",
    )
    achados.sort(key=lambda m: m["recebido_em"] or "", reverse=True)
    log.info("busca livre devolveu %d mensagem(ns)", len(achados))
    return achados


# --------------------------------------------------------------------------- #
# Teste manual
# --------------------------------------------------------------------------- #

if __name__ == "__main__":
    import sys

    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")

    print(f"Perfis MAPI configurados: {perfis_mapi() or '(nenhum)'}")

    disponivel, recado = outlook_disponivel()
    print(f"\nOutlook disponível? {disponivel}\n  {recado}")
    if not disponivel:
        raise SystemExit(1)

    print("\nPastas visíveis:")
    for caminho in listar_pastas()[:40]:
        print(f"  - {caminho}")

    if len(sys.argv) > 1:
        comp = Competencia.de_texto(sys.argv[1])
        alvos = sys.argv[2:]
        print(f"\nBusca livre em {comp} com termos {alvos or '(qualquer e-mail com anexo)'}:")
        for candidato in buscar_candidatos(comp, termos=alvos)[:20]:
            print(f"  {candidato['recebido_em']}  {candidato['remetente']}")
            print(f"      {candidato['assunto']}")
            print(f"      anexos: {candidato['anexos']}")
