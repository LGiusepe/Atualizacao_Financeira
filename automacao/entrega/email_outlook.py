"""
E-mail de autorização, pelo Outlook clássico (COM).

Regras deste módulo:

* **O envio direto NÃO funciona nesta máquina, e o motivo importa.** O COM
  só conversa com o Outlook **clássico**. Aqui quem fica aberto é o Outlook
  **novo** (`olk`), que tem armazenamento próprio; o clássico é iniciado sem
  janela pela automação. Nesse arranjo o `Send()` apenas move a mensagem para
  a *Caixa de Saída* do perfil clássico, e ninguém esvazia essa fila: o
  processo headless morre antes de transmitir e o Outlook novo nem enxerga
  aquela pasta. Testado em 02/09/2026 — cinco mensagens ficaram paradas e o
  painel dizia "enviado". Por isso `email.enviar_automaticamente` está em
  `false` e a etapa volta a só rascunhar. Despacho de verdade exigiria
  Microsoft Graph, que é outro projeto.
* **Se alguém religar o flag, a mentira não passa.** `_montar_mensagem`
  confere se a mensagem saiu da Caixa de Saída e, se não saiu, a etapa termina
  em ERRO dizendo exatamente isso — em vez de anunciar um envio que não
  aconteceu.
* **Simulação sempre vence.** Com `seguranca.simulacao: true` a mensagem é
  rascunhada mesmo que o envio esteja ligado. Modo de teste não fala com
  gente de fora.
* **Enviar é irreversível.** Não existe "desenviar". Por isso o despacho é
  guardado no banco (`email_enviado`/`enviado_em`) e a etapa se recusa a
  repetir — ver `etapa_email` no orquestrador. Rascunho duplicado é chateação;
  e-mail duplicado para o financeiro é retrabalho de outra pessoa.
* **Toda chamada COM é posicional.** O pywin32 em *late binding* descarta
  argumentos nomeados em silêncio — `Copy(After=x)` já virou `Copy()` neste
  projeto e apagou uma aba. Por isso é `Attachments.Add(caminho)` e nunca
  `Attachments.Add(Source=caminho)`.
* **`previa()` não toca no Outlook.** O painel monta a prévia a cada
  carregamento de página; acordar o COM ali travaria a interface. Prévia é
  texto puro; COM só em `criar_rascunho*()`.

Dois formatos de mensagem:

* **Individual** — `previa()` / `criar_rascunho()`: uma conta, um PDF anexo.
* **Agrupado** — `previa_agrupada()` / `criar_rascunho_agrupado()`: várias
  contas do mesmo fornecedor num único e-mail, **um anexo por fatura**. Foi a
  decisão do usuário para o fornecedor que tem quatro faturas por mês, uma
  por serviço contratado: um e-mail só, quatro PDFs separados — nada de
  mesclar tudo num arquivo, nada de quatro e-mails.
"""

from __future__ import annotations

import logging
import re
import shutil
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterator, Sequence

from automacao.nucleo.config import Ambiente, ambiente
from automacao.nucleo.modelos import Competencia, Conta, Etapa, ResultadoEtapa, Situacao

# Texto longo e já revisado: reaproveitado em vez de reescrito.
from automacao.coleta.outlook import MENSAGEM_SEM_OUTLOOK

log = logging.getLogger("automacao.entrega.email_outlook")

__all__ = [
    "ErroEmail",
    "previa",
    "previa_agrupada",
    "criar_rascunho",
    "criar_rascunho_agrupado",
    "empresa_da_assinatura",
    "montar_corpo",
    "formatar_valor",
    "formatar_data",
    "LIMITE_ANEXOS_MB",
]

# Constantes do modelo de objetos do Outlook.
ITEM_DE_EMAIL = 0        # olMailItem
PASTA_CAIXA_DE_SAIDA = 4  # olFolderOutbox

# Teto prático de anexos por mensagem no Exchange corporativo. Acima disso o
# servidor recusa na hora do envio — e o usuário só descobre depois de clicar.
LIMITE_ANEXOS_MB = 20.0
LIMITE_ANEXOS_BYTES = int(LIMITE_ANEXOS_MB * 1024 * 1024)

# Usados quando o settings.yaml não traz o molde correspondente.
ASSUNTO_PADRAO = "Fatura {empresa}"
TRECHO_CIDADE_PADRAO = " — unidade {cidade}"

CORPO_PADRAO = """\
Boa tarde.

Segue em anexo a autorização de pagamento referente à fatura da {empresa}{trecho_cidade}, competência {competencia_extenso}.

Vencimento: {vencimento}
Valor: {valor}

Qualquer dúvida, estou à disposição.

Atenciosamente,
"""

CORPO_AGRUPADO_PADRAO = """\
Boa tarde.

Seguem em anexo as autorizações de pagamento referentes às faturas da {empresa}, competência {competencia_extenso}. Cada fatura vai em um anexo separado:

{lista_faturas}

Valor total: {valor_total}

Qualquer dúvida, estou à disposição.

Atenciosamente,
"""

# Como não temos o dado, é isso que aparece no lugar do valor/vencimento.
A_CONFIRMAR = "(a confirmar)"


class ErroEmail(RuntimeError):
    """Não foi possível preparar ou rascunhar a mensagem."""


# --------------------------------------------------------------------------- #
# Formatação
# --------------------------------------------------------------------------- #


def formatar_valor(valor: float | int | str | None) -> str:
    """`1234.56` vira `'R$ 1.234,56'`. Sem valor, devolve `'(a confirmar)'`."""
    if valor is None or valor == "":
        return A_CONFIRMAR
    try:
        numero = float(valor)
    except (TypeError, ValueError):
        return str(valor)
    # Formata no padrão americano e troca os separadores de uma vez só.
    americano = f"{numero:,.2f}"
    return "R$ " + americano.replace(",", "\x00").replace(".", ",").replace("\x00", ".")


def formatar_data(quando: date | datetime | str | None) -> str:
    """Qualquer data vira `'dd/mm/aaaa'`. Sem data, devolve `'(a confirmar)'`."""
    if quando in (None, ""):
        return A_CONFIRMAR
    if isinstance(quando, datetime):
        return quando.strftime("%d/%m/%Y")
    if isinstance(quando, date):
        return quando.strftime("%d/%m/%Y")
    texto = str(quando).strip()
    for molde in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(texto, molde).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return texto


class _Valores(dict):
    """Dicionário de placeholders que não explode com chave desconhecida."""

    def __missing__(self, chave: str) -> str:
        log.warning("placeholder {%s} não existe neste molde de e-mail", chave)
        return "{" + chave + "}"


def _preencher(molde: str, **valores: Any) -> str:
    """`str.format` tolerante: molde com placeholder errado não derruba o painel."""
    try:
        return molde.format_map(_Valores(valores))
    except (IndexError, ValueError) as exc:
        log.warning("molde de e-mail malformado (%s) — usando o texto cru", exc)
        return molde


def montar_corpo(molde: str, assinatura: str = "", **valores: Any) -> str:
    """
    Preenche o molde e garante a assinatura no fim — uma vez só.

    Três situações, e as três precisam sair certas:

    * o molde traz `{assinatura}` — `_preencher` já colocou no lugar escolhido
      por quem escreveu o texto, e não há o que acrescentar;
    * o molde não traz (é o caso de todo `settings.yaml` escrito antes da tela
      de Configuração › E-mail) — o bloco entra no fim, separado por uma linha
      em branco;
    * o molde já termina com a assinatura **digitada à mão** — foi assim que o
      settings.yaml de produção nasceu, com as três linhas batidas no texto.
      Acrescentar de novo sairia duplicado no e-mail do fornecedor.

    O `rstrip` no fim não é cosmético: os moldes em uso terminam em
    "Atenciosamente," seguido de linhas em branco, e uma assinatura três
    linhas abaixo do resto parece anexo de outra mensagem. Sem assinatura
    nenhuma, essas linhas também não podem sobrar.
    """
    texto = _preencher(molde, assinatura=assinatura, **valores)
    if "{assinatura}" in molde:
        return texto

    limpo = texto.rstrip()
    bloco = (assinatura or "").strip()
    if not bloco or limpo.endswith(bloco):
        return limpo + "\n"
    return f"{limpo}\n\n{bloco}\n"


# --------------------------------------------------------------------------- #
# Destinatários e moldes, lidos de `ambiente().email`
# --------------------------------------------------------------------------- #


def _pessoas(bruto: Any) -> list[tuple[str, str]]:
    """
    Normaliza o bloco `email.para` / `email.copia` para `[(nome, endereço)]`.

    Aceita as três formas que já apareceram na configuração: lista de
    dicionários `{nome, endereco}`, lista de strings e string única.
    """
    if not bruto:
        return []
    if isinstance(bruto, str):
        bruto = [bruto]
    pessoas: list[tuple[str, str]] = []
    for item in bruto:
        if isinstance(item, dict):
            endereco = str(item.get("endereco") or item.get("email") or "").strip()
            nome = str(item.get("nome") or "").strip()
        else:
            endereco, nome = str(item).strip(), ""
        if endereco:
            pessoas.append((nome, endereco))
    return pessoas


def _para_humano(pessoas: Sequence[tuple[str, str]]) -> str:
    """`'Fulana de Tal <fulana@exemplo.com.br>'` — é o que o painel mostra."""
    return "; ".join(f"{n} <{e}>" if n else e for n, e in pessoas) or "(ninguém configurado)"


def _para_outlook(pessoas: Sequence[tuple[str, str]]) -> str:
    """Só os endereços, como o campo `To`/`CC` do Outlook espera."""
    return "; ".join(e for _n, e in pessoas)


def _molde(amb: Ambiente, chave: str, padrao: str) -> str:
    bruto = amb.email.get(chave)
    return str(bruto) if bruto else padrao


def _molde_do_corpo(amb: Ambiente, chave: str, padrao: str, pessoal: str) -> str:
    """
    Qual texto vale: o da pessoa, o do settings.yaml ou a constante.

    O molde pessoal (Configuração › E-mail) vence o global porque foi escrito
    depois e por quem vai assinar. Em branco não é escolha — é campo que
    ninguém preencheu —, então cai no global.
    """
    if (pessoal or "").strip():
        return pessoal
    return _molde(amb, chave, padrao)


def duplicaria_assinatura(molde: str, assinatura: str) -> bool:
    """
    Se este molde faria o nome de quem assina sair **duas vezes**.

    Existe porque o `settings.yaml` em uso nasceu assim: as três linhas da
    assinatura digitadas dentro do `email.corpo`, antes de haver tela para
    isto. Com função e setor ainda em branco no cadastro, a assinatura montada
    fica *parecida mas não igual* à digitada — a guarda de `montar_corpo` não
    reconhece, o bloco é acrescentado abaixo, e o fornecedor recebe o nome nas
    duas pontas. Só se vê no rascunho, que já é tarde.

    Não conserta nada sozinha: a tela avisa e quem escreveu o texto decide. Os
    três casos em que devolve `False` são os que já saem certos — molde com
    `{assinatura}`, molde que termina exatamente no bloco (aí `montar_corpo`
    não acrescenta) e molde sem assinatura nenhuma.

    O nome é procurado só nas últimas linhas: citado no meio, é assunto da
    mensagem, não assinatura.
    """
    bloco = (assinatura or "").strip()
    if not molde or not bloco or "{assinatura}" in molde:
        return False
    limpo = molde.rstrip()
    if limpo.endswith(bloco):
        return False
    nome = bloco.splitlines()[0].strip()
    return bool(nome) and nome in [l.strip() for l in limpo.splitlines()[-4:]]


def empresa_da_assinatura(amb: Ambiente) -> str:
    """
    Nome da empresa que fecha a assinatura, de `email.assinatura_empresa`.

    Global de propósito: a empresa é a mesma para todo mundo que usa o painel,
    e deixar cada um digitar o nome produziria três grafias diferentes saindo
    para o mesmo fornecedor.
    """
    return str(amb.email.get("assinatura_empresa") or "").strip()


def _trecho_cidade(conta: Conta, amb: Ambiente) -> str:
    """
    Complemento do assunto/corpo com a cidade atendida.

    Só existe para operadora (`conta.email.eh_operadora`), porque só nela a
    mesma empresa manda várias faturas e a cidade é o que diferencia uma da
    outra. Para as demais, string vazia — o molde tem que continuar legível.
    """
    if not conta.email.eh_operadora:
        return ""
    cidade = (conta.email.cidade or "").strip()
    if not cidade:
        log.warning(
            "conta %r está marcada como operadora mas não tem 'email.cidade' — "
            "o e-mail vai sair sem a unidade no texto",
            conta.id,
        )
        return ""
    return _preencher(_molde(amb, "trecho_cidade_operadora", TRECHO_CIDADE_PADRAO), cidade=cidade)


def deve_enviar(amb: Ambiente) -> bool:
    """
    Se esta mensagem sai agora ou fica em Rascunhos.

    Duas condições, e a simulação tem a última palavra: ninguém liga o modo de
    teste esperando que ele converse com o financeiro.
    """
    pedido = bool(amb.email.get("enviar_automaticamente"))
    if pedido and amb.simulando:
        log.warning(
            "envio automático está ligado, mas a simulação também — a mensagem "
            "vai ficar em Rascunhos. Desligue seguranca.simulacao para despachar."
        )
        return False
    return pedido


# --------------------------------------------------------------------------- #
# Prévia individual
# --------------------------------------------------------------------------- #


def previa(
    conta: Conta,
    competencia: Competencia,
    *,
    valor: float | None,
    vencimento: date | None,
    ambiente_: Ambiente | None = None,
    para: str | None = None,
    copia: str | None = None,
    assunto: str | None = None,
    corpo: str | None = None,
    assinatura: str = "",
    corpo_pessoal: str = "",
) -> dict:
    """
    Monta o texto do e-mail sem encostar no Outlook.

    O painel chama isto a cada carregamento de página da conta, então tem que
    ser barato e nunca falhar: é só leitura de configuração e formatação.

    Args:
        conta: conta cuja fatura será enviada.
        competencia: mês de referência.
        valor: valor da fatura; `None` vira "(a confirmar)" no texto.
        vencimento: data de vencimento; `None` vira "(a confirmar)".
        ambiente_: ambiente já carregado (útil em teste).
        para: destinatários digitados no painel, no formato
            `'Nome <a@b>; c@d'`. Substituem o bloco `email.para` só nesta
            mensagem — a configuração não é alterada.
        copia: idem, para o campo de cópia. String vazia tira todo mundo.
        assunto: assunto digitado, no lugar do molde.
        corpo: texto digitado, no lugar do molde. Sai exatamente como veio: já
            foi mostrado na prévia com a assinatura dentro, e acrescentá-la de
            novo passaria por cima de quem a tirou de propósito.
        assinatura: bloco de três linhas de quem está logado
            (`usuarios.assinatura_de`). Vazio = corpo sem assinatura.
        corpo_pessoal: molde do usuário (Configuração › E-mail). Vence o
            `email.corpo` do settings.yaml; vazio cai nele.

    Returns:
        `{para, copia, assunto, corpo}` prontos para exibição, mais
        `para_enderecos` / `copia_enderecos` (o que vai para o Outlook) e
        `editado` dizendo quais campos vieram do painel.
    """
    amb = ambiente_ or ambiente()

    # `None` = usa a configuração. String (mesmo vazia) = o que você digitou.
    destinatarios = _pessoas(amb.email.get("para")) if para is None else _pessoas_de_texto(para)
    copiados = _pessoas(amb.email.get("copia")) if copia is None else _pessoas_de_texto(copia)

    empresa = conta.nome_empresa
    valores = {
        "empresa": empresa,
        "competencia_extenso": competencia.extenso,
        "competencia": str(competencia),
        "vencimento": formatar_data(vencimento),
        "valor": formatar_valor(valor),
        "trecho_cidade": _trecho_cidade(conta, amb),
    }

    editado = [
        campo
        for campo, dado in (
            ("para", para), ("copia", copia), ("assunto", assunto), ("corpo", corpo),
        )
        if dado is not None
    ]

    return {
        "para": _para_humano(destinatarios),
        "copia": _para_humano(copiados),
        "assunto": assunto if assunto is not None
                   else _preencher(_molde(amb, "assunto", ASSUNTO_PADRAO), **valores),
        "corpo": corpo if corpo is not None
                 else montar_corpo(
                     _molde_do_corpo(amb, "corpo", CORPO_PADRAO, corpo_pessoal),
                     assinatura,
                     **valores,
                 ),
        "para_enderecos": _para_outlook(destinatarios),
        "copia_enderecos": _para_outlook(copiados),
        "enderecos_suspeitos": _suspeitos(destinatarios + copiados),
        "editado": editado,
        "assinatura": assinatura,
    }


def _pessoas_de_texto(bruto: str) -> list[tuple[str, str]]:
    """
    `'Fulana <fulana@x.br>; ti@x.br'` -> `[('Fulana','fulana@x.br'), ...]`

    Aceita `;` e `,` como separador, e o formato `Nome <endereço>` que o
    próprio painel exibe — assim dá para copiar, colar e editar.
    """
    pessoas: list[tuple[str, str]] = []
    for pedaco in re.split(r"[;,]", bruto or ""):
        pedaco = pedaco.strip()
        if not pedaco:
            continue
        achado = re.match(r"^(.*?)<([^>]+)>$", pedaco)
        if achado:
            nome, endereco = achado.group(1).strip(), achado.group(2).strip()
        else:
            nome, endereco = "", pedaco
        if endereco:
            pessoas.append((nome, endereco))
    return pessoas


def _suspeitos(pessoas: Sequence[tuple[str, str]]) -> list[str]:
    """Endereços que não parecem endereços — o painel avisa antes de rascunhar."""
    return [e for _n, e in pessoas if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", e)]


# --------------------------------------------------------------------------- #
# Prévia agrupada (o caso do fornecedor com quatro faturas)
# --------------------------------------------------------------------------- #


@dataclass
class _Fatura:
    """Uma linha do e-mail agrupado, já normalizada."""

    conta: Conta
    pdf: Path | None
    valor: float | None
    vencimento: date | None

    @property
    def rotulo(self) -> str:
        """Como a fatura é chamada no corpo: a unidade/serviço, quando existe."""
        return self.conta.subunidade or self.conta.nome_empresa


def _normalizar_grupo(contas_e_dados: Sequence[dict]) -> list[_Fatura]:
    """Valida a lista de dicionários e devolve `_Fatura` na mesma ordem."""
    faturas: list[_Fatura] = []
    for posicao, item in enumerate(contas_e_dados, start=1):
        if not isinstance(item, dict):
            raise ErroEmail(
                f"item {posicao} do grupo não é um dicionário "
                f"{{conta, pdf, valor, vencimento}}: {item!r}"
            )
        conta = item.get("conta")
        if not isinstance(conta, Conta):
            raise ErroEmail(f"item {posicao} do grupo está sem a chave 'conta'")
        pdf = item.get("pdf")
        faturas.append(
            _Fatura(
                conta=conta,
                pdf=Path(pdf) if pdf else None,
                valor=item.get("valor"),
                vencimento=item.get("vencimento"),
            )
        )
    if not faturas:
        raise ErroEmail("o grupo está vazio — nada a rascunhar")
    return faturas


def _empresa_do_grupo(faturas: Sequence[_Fatura]) -> str:
    """
    Nome do fornecedor comum, para o assunto ("Fatura <fornecedor>").

    Se o grupo misturar fornecedores — o que não deveria acontecer — juntamos
    os nomes em vez de escolher um, para o erro ficar visível no assunto.
    """
    nomes: list[str] = []
    for fatura in faturas:
        nome = fatura.conta.nome_empresa
        if nome not in nomes:
            nomes.append(nome)
    return " / ".join(nomes)


def _lista_faturas(faturas: Sequence[_Fatura]) -> str:
    """Uma linha por fatura: unidade/serviço, vencimento e valor."""
    return "\n".join(
        f"- {f.rotulo}: vencimento {formatar_data(f.vencimento)}, "
        f"valor {formatar_valor(f.valor)}"
        for f in faturas
    )


def _vencimentos_distintos(faturas: Sequence[_Fatura]) -> list[date]:
    """Vencimentos diferentes encontrados no grupo, em ordem."""
    distintos: list[date] = []
    for fatura in faturas:
        if fatura.vencimento and fatura.vencimento not in distintos:
            distintos.append(fatura.vencimento)
    return sorted(distintos)


def previa_agrupada(
    contas_e_dados: Sequence[dict],
    competencia: Competencia,
    *,
    ambiente_: Ambiente | None = None,
    assinatura: str = "",
    corpo_pessoal: str = "",
) -> dict:
    """
    Prévia do e-mail único que cobre várias faturas do mesmo fornecedor.

    Não encosta no Outlook — mesma promessa de `previa()`.

    Args:
        contas_e_dados: lista de `{conta, pdf, valor, vencimento}`, uma entrada
            por fatura. A ordem da lista é a ordem do corpo e dos anexos.
        competencia: mês de referência.
        ambiente_: ambiente já carregado (útil em teste).
        assinatura: bloco de três linhas de quem está logado.
        corpo_pessoal: molde agrupado do usuário; vazio cai no global.

    Returns:
        `{para, copia, assunto, corpo}` mais `valor_total`, `faturas`
        (rótulo/vencimento/valor/anexo de cada uma), `anexos` e `avisos`.
        Vencimentos diferentes entram em `avisos` — quem transforma isso em
        `Situacao.ATENCAO` é `criar_rascunho_agrupado()`.

    Raises:
        ErroEmail: lista vazia ou item sem a chave `conta`.
    """
    amb = ambiente_ or ambiente()
    faturas = _normalizar_grupo(contas_e_dados)
    destinatarios = _pessoas(amb.email.get("para"))
    copiados = _pessoas(amb.email.get("copia"))

    total = sum(float(f.valor) for f in faturas if f.valor is not None)
    sem_valor = [f.rotulo for f in faturas if f.valor is None]

    valores = {
        "empresa": _empresa_do_grupo(faturas),
        "competencia_extenso": competencia.extenso,
        "competencia": str(competencia),
        "lista_faturas": _lista_faturas(faturas),
        "valor_total": formatar_valor(total) if not sem_valor else f"{formatar_valor(total)} (parcial)",
    }

    avisos: list[str] = []
    distintos = _vencimentos_distintos(faturas)
    if len(distintos) > 1:
        avisos.append(
            "as faturas do grupo têm vencimentos diferentes ("
            + ", ".join(formatar_data(d) for d in distintos)
            + ") — cada um está listado no corpo, confira antes de mandar."
        )
    if sem_valor:
        avisos.append(
            "sem valor informado para: " + ", ".join(sem_valor)
            + " — o total do corpo está incompleto."
        )
    if len({f.conta.nome_empresa for f in faturas}) > 1:
        avisos.append(
            "o grupo mistura fornecedores diferentes — o e-mail agrupado foi "
            "pensado para faturas da mesma empresa."
        )

    return {
        "para": _para_humano(destinatarios),
        "copia": _para_humano(copiados),
        "assunto": _preencher(_molde(amb, "assunto", ASSUNTO_PADRAO), **valores),
        "corpo": montar_corpo(
            _molde_do_corpo(amb, "corpo_agrupado", CORPO_AGRUPADO_PADRAO, corpo_pessoal),
            assinatura,
            **valores,
        ),
        "assinatura": assinatura,
        "para_enderecos": _para_outlook(destinatarios),
        "copia_enderecos": _para_outlook(copiados),
        "valor_total": total,
        "faturas": [
            {
                "conta_id": f.conta.id,
                "rotulo": f.rotulo,
                "vencimento": f.vencimento.isoformat() if f.vencimento else None,
                "valor": f.valor,
                "anexo": f.pdf.name if f.pdf else None,
            }
            for f in faturas
        ],
        "anexos": [str(f.pdf) for f in faturas if f.pdf],
        "avisos": avisos,
    }


# --------------------------------------------------------------------------- #
# Anexos
# --------------------------------------------------------------------------- #


@dataclass
class _Anexos:
    """Resultado da conferência dos arquivos que vão junto na mensagem."""

    caminhos: list[Path] = field(default_factory=list)
    problemas: list[str] = field(default_factory=list)
    temporarios: list[Path] = field(default_factory=list)

    @property
    def bytes_totais(self) -> int:
        return sum(c.stat().st_size for c in self.caminhos if c.is_file())

    @property
    def mb_totais(self) -> float:
        return self.bytes_totais / (1024 * 1024)

    @property
    def excede_limite(self) -> bool:
        return self.bytes_totais > LIMITE_ANEXOS_BYTES


def _conferir_anexo(pdf: Path | None, rotulo: str, destino: _Anexos) -> None:
    """Valida um arquivo e o acrescenta à lista, ou registra o problema."""
    if pdf is None:
        destino.problemas.append(f"{rotulo}: nenhum PDF informado")
        return
    if not pdf.is_file():
        destino.problemas.append(f"{rotulo}: o PDF não existe ({pdf})")
        return
    if pdf.stat().st_size <= 0:
        destino.problemas.append(f"{rotulo}: o PDF está vazio (0 byte) — {pdf.name}")
        return
    destino.caminhos.append(pdf)


def _desambiguar(faturas: Sequence[_Fatura], anexos: _Anexos, amb: Ambiente,
                 competencia: Competencia) -> None:
    """
    Garante que cada anexo tenha nome próprio.

    O Outlook nomeia o anexo pelo arquivo de origem; dois PDFs chamados
    "fatura.pdf" chegariam indistinguíveis. Quando isso acontece, fazemos uma
    cópia renomeada — sempre dentro de `dados/`, nunca no OneDrive.
    """
    vistos: dict[str, int] = {}
    for indice, caminho in enumerate(list(anexos.caminhos)):
        chave = caminho.name.lower()
        vistos[chave] = vistos.get(chave, 0) + 1
        if vistos[chave] == 1:
            continue
        fatura = faturas[indice] if indice < len(faturas) else None
        prefixo = fatura.conta.id if fatura else f"anexo{indice + 1}"
        rascunhos = amb.caminhos.trabalho / str(competencia) / "_anexos_agrupados"
        rascunhos.mkdir(parents=True, exist_ok=True)
        copia = rascunhos / f"{prefixo} - {caminho.name}"
        try:
            shutil.copy2(caminho, copia)
        except OSError as exc:
            anexos.problemas.append(f"{caminho.name}: não consegui renomear a cópia ({exc})")
            continue
        anexos.caminhos[indice] = copia
        anexos.temporarios.append(copia)
        log.info("anexo renomeado para não colidir: %s", copia.name)


# --------------------------------------------------------------------------- #
# Conexão com o Outlook
# --------------------------------------------------------------------------- #


def _descrever_erro(exc: BaseException) -> str:
    texto = str(exc).strip()
    return f"{type(exc).__name__}: {texto}" if texto else type(exc).__name__


@contextmanager
def _com_inicializado() -> Iterator[None]:
    """Inicializa o COM nesta thread e devolve tudo no fim."""
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
            except Exception as exc:  # pragma: no cover - defensivo
                log.debug("CoUninitialize falhou: %s", exc)


def _abrir_outlook() -> Any:
    """
    Objeto `Outlook.Application`, tentando as três estratégias na mesma ordem
    que `coleta.outlook` usa.

    `GetActiveObject` só funciona quando o Outlook já se registrou na ROT
    visível para este processo; na máquina de produção quem resolve é o
    `Dispatch`, e o `DispatchEx` fica de último recurso. Aqui só precisamos do
    objeto da aplicação — nada de namespace MAPI, que é o que costuma pendurar.

    Raises:
        ErroEmail: quando as três falham.
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
            tentativas.append(f"{nome} → {_descrever_erro(exc)}")
            continue
        log.info("conectado ao Outlook %s via %s", versao, nome)
        return aplicacao

    raise ErroEmail(f"{MENSAGEM_SEM_OUTLOOK} Tentativas: {' | '.join(tentativas)}")


def _montar_mensagem(
    *,
    para: str,
    copia: str,
    assunto: str,
    corpo: str,
    anexos: Sequence[Path],
    abrir: bool,
    enviar: bool,
) -> dict:
    """
    Monta a mensagem no Outlook e, conforme `enviar`, despacha ou guarda.

    Todas as chamadas COM são posicionais, de propósito: em *late binding* o
    pywin32 descarta argumento nomeado sem avisar.

    O `EntryID` é lido ANTES do `Send()`: depois de despachada a mensagem sai
    de Rascunhos e o objeto pode não responder mais.

    Returns:
        `{"entry_id": ..., "exibida": bool, "enviada": bool}`.
    """
    with _com_inicializado():
        aplicacao = _abrir_outlook()
        mensagem = aplicacao.CreateItem(ITEM_DE_EMAIL)      # posicional
        mensagem.To = para
        mensagem.CC = copia
        mensagem.Subject = assunto
        mensagem.Body = corpo

        for caminho in anexos:
            # POSICIONAL. Jamais `Attachments.Add(Source=...)`.
            mensagem.Attachments.Add(str(caminho))
            log.info("anexo incluído na mensagem: %s", Path(caminho).name)

        mensagem.Save()

        try:
            entry_id = str(mensagem.EntryID)
        except Exception:
            entry_id = ""

        exibida = False
        if abrir and not enviar:
            # Abrir a janela de algo que já foi despachado não faz sentido.
            try:
                mensagem.Display(False)                     # posicional, não-modal
                exibida = True
            except Exception as exc:  # janela é conforto, não requisito
                log.warning("não consegui abrir a janela do rascunho: %s", _descrever_erro(exc))

        if enviar:
            mensagem.Send()
            if _ainda_na_caixa_de_saida(aplicacao, assunto):
                raise ErroEmail(
                    "o Outlook aceitou a mensagem mas ela ficou parada na Caixa "
                    "de Saída — não foi transmitida. Acontece quando só o "
                    "Outlook novo está aberto: o Send() do COM entrega ao perfil "
                    "clássico, que não fica rodando para esvaziar a fila. Abra a "
                    "Caixa de Saída do Outlook clássico e decida o que fazer com "
                    "ela, e volte email.enviar_automaticamente para false."
                )
            log.info("mensagem despachada para %s — assunto %r", para, assunto)

    return {"entry_id": entry_id, "exibida": exibida, "enviada": bool(enviar)}


def _ainda_na_caixa_de_saida(aplicacao: Any, assunto: str) -> bool:
    """
    Se a mensagem continua na Caixa de Saída depois do `Send()`.

    `Send()` só move o item para a fila; a transmissão é outra etapa, e nesta
    máquina ela não acontece (ver o topo do arquivo). Sem esta conferência a
    automação anunciava "e-mail enviado" para algo que nunca saiu — foi assim
    que cinco autorizações ficaram paradas sem ninguém notar.

    Na dúvida devolve `False`: acusar um envio que deu certo seria pior que
    não acusar. Quem confere de verdade é a Caixa de Saída, à vista.
    """
    try:
        fila = aplicacao.GetNamespace("MAPI").GetDefaultFolder(PASTA_CAIXA_DE_SAIDA)
        itens = fila.Items
        for i in range(1, min(int(itens.Count), 25) + 1):
            if str(getattr(itens.Item(i), "Subject", "")) == assunto:
                log.error("mensagem %r ficou na Caixa de Saída — não transmitida", assunto)
                return True
    except Exception as exc:  # noqa: BLE001 — conferência é best-effort
        log.warning("não consegui conferir a Caixa de Saída: %s", _descrever_erro(exc))
    return False


# --------------------------------------------------------------------------- #
# Rascunho individual
# --------------------------------------------------------------------------- #


def criar_rascunho(
    conta: Conta,
    competencia: Competencia,
    *,
    pdf: Path,
    valor: float | None,
    vencimento: date | None,
    ambiente_: Ambiente | None = None,
    abrir: bool = True,
    para: str | None = None,
    copia: str | None = None,
    assunto: str | None = None,
    corpo: str | None = None,
    assinatura: str = "",
    corpo_pessoal: str = "",
) -> ResultadoEtapa:
    """
    Deixa o e-mail da conta pronto em *Rascunhos*, com o PDF final anexado.

    Não despacha nada: salva e, se `abrir`, mostra a janela para conferência.

    Args:
        conta: conta cuja fatura será enviada.
        competencia: mês de referência.
        pdf: PDF único do mês (autorização + demonstrativo + boleto + nota).
        valor: valor da fatura.
        vencimento: data de vencimento.
        ambiente_: ambiente já carregado (útil em teste).
        abrir: abre a janela do rascunho no Outlook.
        assinatura: bloco de três linhas de quem está logado.
        corpo_pessoal: molde do usuário; vazio cai no global.

    Returns:
        `ResultadoEtapa` de `Etapa.EMAIL`. `Situacao.ATENCAO` quando o anexo
        passa de 20 MB — o servidor recusa mensagens acima disso.
        Nunca levanta exceção: problema de Outlook vira `Situacao.ERRO`.
    """
    amb = ambiente_ or ambiente()
    etapa = Etapa.EMAIL
    enviar = deve_enviar(amb)

    texto = previa(
        conta,
        competencia,
        valor=valor,
        vencimento=vencimento,
        ambiente_=amb,
        para=para,
        copia=copia,
        assunto=assunto,
        corpo=corpo,
        assinatura=assinatura,
        corpo_pessoal=corpo_pessoal,
    )

    anexos = _Anexos()
    _conferir_anexo(Path(pdf) if pdf else None, conta.rotulo, anexos)
    if anexos.problemas:
        return ResultadoEtapa.erro(
            etapa,
            "não dá para rascunhar sem o anexo: " + "; ".join(anexos.problemas),
            detalhes={**texto, "problemas": anexos.problemas},
        )

    if not texto["para_enderecos"]:
        return ResultadoEtapa.erro(
            etapa,
            "o campo Para está vazio — informe pelo menos um destinatário."
            if para is not None
            else "não há destinatário em config/settings.yaml (bloco email.para).",
            detalhes=texto,
        )

    if texto["enderecos_suspeitos"]:
        return ResultadoEtapa.erro(
            etapa,
            "endereço com formato inválido: "
            + ", ".join(texto["enderecos_suspeitos"])
            + ". Corrija antes de rascunhar.",
            detalhes=texto,
        )

    detalhes = {
        **texto,
        "anexo": anexos.caminhos[0].name,
        "anexo_caminho": str(anexos.caminhos[0]),
        "anexos_mb": round(anexos.mb_totais, 2),
    }

    # O anexo grande é conferido ANTES de despachar: passar do teto do
    # servidor com um rascunho custa um clique; com um envio, custa uma
    # mensagem recusada que ninguém vê voltar.
    if enviar and anexos.excede_limite:
        return ResultadoEtapa.erro(
            etapa,
            f"o anexo tem {anexos.mb_totais:.1f} MB e o servidor recusa acima de "
            f"{LIMITE_ANEXOS_MB:.0f} MB — o e-mail NÃO foi enviado. Compacte o PDF "
            f"ou desligue email.enviar_automaticamente para mandar à mão.",
            detalhes=detalhes,
        )

    try:
        criado = _montar_mensagem(
            para=texto["para_enderecos"],
            copia=texto["copia_enderecos"],
            assunto=texto["assunto"],
            corpo=texto["corpo"],
            anexos=anexos.caminhos,
            abrir=abrir,
            enviar=enviar,
        )
    except ErroEmail as exc:
        return ResultadoEtapa.erro(etapa, str(exc), detalhes=detalhes)
    except Exception as exc:  # noqa: BLE001 - o painel não pode quebrar por causa do COM
        log.exception("falha ao %s o e-mail de %s",
                      "enviar" if enviar else "rascunhar", conta.id)
        return ResultadoEtapa.erro(
            etapa,
            f"{MENSAGEM_SEM_OUTLOOK} Detalhe técnico: {_descrever_erro(exc)}",
            detalhes=detalhes,
        )

    detalhes.update(criado)
    if criado["enviada"]:
        mensagem = (
            f"e-mail enviado para {texto['para']} — assunto {texto['assunto']!r}, "
            f"anexo {detalhes['anexo']} ({detalhes['anexos_mb']:.2f} MB). "
            f"Saiu do Outlook; está em Itens Enviados."
        )
    else:
        mensagem = (
            f"rascunho salvo em Rascunhos para {texto['para']} — "
            f"assunto {texto['assunto']!r}, anexo {detalhes['anexo']} "
            f"({detalhes['anexos_mb']:.2f} MB). Nada foi despachado: confira e clique você."
        )

    if anexos.excede_limite:
        return ResultadoEtapa.atencao(
            etapa,
            f"{mensagem} ATENÇÃO: o anexo tem {anexos.mb_totais:.1f} MB e o servidor "
            f"costuma recusar acima de {LIMITE_ANEXOS_MB:.0f} MB — considere compactar "
            f"o PDF antes de mandar.",
            detalhes=detalhes,
            artefatos=list(anexos.caminhos),
        )
    return ResultadoEtapa.sucesso(
        etapa, mensagem, detalhes=detalhes, artefatos=list(anexos.caminhos)
    )


# --------------------------------------------------------------------------- #
# Rascunho agrupado (o caso do fornecedor com quatro faturas)
# --------------------------------------------------------------------------- #


def criar_rascunho_agrupado(
    contas_e_dados: Sequence[dict],
    competencia: Competencia,
    *,
    ambiente_: Ambiente | None = None,
    abrir: bool = True,
    assinatura: str = "",
    corpo_pessoal: str = "",
) -> ResultadoEtapa:
    """
    Um e-mail só para várias faturas do mesmo fornecedor, um anexo por fatura.

    É o formato escolhido pelo usuário para o fornecedor com quatro faturas
    por mês, uma por serviço contratado, em uma única mensagem, com os quatro
    PDFs anexados separadamente. Nada é mesclado
    e nada é despachado.

    Args:
        contas_e_dados: lista de `{conta, pdf, valor, vencimento}`, uma entrada
            por fatura. A ordem vira a ordem do corpo e dos anexos.
        competencia: mês de referência.
        ambiente_: ambiente já carregado (útil em teste).
        abrir: abre a janela do rascunho no Outlook.

    Returns:
        `ResultadoEtapa` de `Etapa.EMAIL`. Devolve `Situacao.ATENCAO` quando os
        vencimentos do grupo divergem, quando falta valor em alguma fatura ou
        quando os anexos somados passam de 20 MB.
    """
    amb = ambiente_ or ambiente()
    etapa = Etapa.EMAIL
    enviar = deve_enviar(amb)

    try:
        faturas = _normalizar_grupo(contas_e_dados)
        texto = previa_agrupada(
            contas_e_dados,
            competencia,
            ambiente_=amb,
            assinatura=assinatura,
            corpo_pessoal=corpo_pessoal,
        )
    except ErroEmail as exc:
        return ResultadoEtapa.erro(etapa, str(exc))

    anexos = _Anexos()
    for fatura in faturas:
        _conferir_anexo(fatura.pdf, fatura.rotulo, anexos)
    if anexos.problemas:
        return ResultadoEtapa.erro(
            etapa,
            "não dá para rascunhar o e-mail agrupado: " + "; ".join(anexos.problemas),
            detalhes={**texto, "problemas": anexos.problemas},
        )
    _desambiguar(faturas, anexos, amb, competencia)

    if not texto["para_enderecos"]:
        return ResultadoEtapa.erro(
            etapa,
            "não há destinatário em config/settings.yaml (bloco email.para).",
            detalhes=texto,
        )

    detalhes = {
        **texto,
        "anexo": ", ".join(c.name for c in anexos.caminhos),
        "anexos_caminhos": [str(c) for c in anexos.caminhos],
        "anexos_mb": round(anexos.mb_totais, 2),
        "contas": [f.conta.id for f in faturas],
    }

    if enviar and anexos.excede_limite:
        return ResultadoEtapa.erro(
            etapa,
            f"os anexos somam {anexos.mb_totais:.1f} MB e o servidor recusa acima de "
            f"{LIMITE_ANEXOS_MB:.0f} MB — o e-mail NÃO foi enviado. Divida em duas "
            f"mensagens ou desligue email.enviar_automaticamente.",
            detalhes=detalhes,
        )

    try:
        criado = _montar_mensagem(
            para=texto["para_enderecos"],
            copia=texto["copia_enderecos"],
            assunto=texto["assunto"],
            corpo=texto["corpo"],
            anexos=anexos.caminhos,
            abrir=abrir,
            enviar=enviar,
        )
    except ErroEmail as exc:
        return ResultadoEtapa.erro(etapa, str(exc), detalhes=detalhes)
    except Exception as exc:  # noqa: BLE001
        log.exception("falha no e-mail agrupado de %s", detalhes["contas"])
        return ResultadoEtapa.erro(
            etapa,
            f"{MENSAGEM_SEM_OUTLOOK} Detalhe técnico: {_descrever_erro(exc)}",
            detalhes=detalhes,
        )

    detalhes.update(criado)
    if criado["enviada"]:
        mensagem = (
            f"e-mail único enviado com {len(anexos.caminhos)} anexo(s) "
            f"({anexos.mb_totais:.2f} MB) cobrindo {len(faturas)} fatura(s) — "
            f"assunto {texto['assunto']!r}. Está em Itens Enviados."
        )
    else:
        mensagem = (
            f"rascunho único salvo com {len(anexos.caminhos)} anexo(s) "
            f"({anexos.mb_totais:.2f} MB) cobrindo {len(faturas)} fatura(s) — "
            f"assunto {texto['assunto']!r}. Nada foi despachado."
        )

    avisos = list(texto["avisos"])
    if anexos.excede_limite:
        avisos.append(
            f"os anexos somam {anexos.mb_totais:.1f} MB e o servidor costuma recusar "
            f"acima de {LIMITE_ANEXOS_MB:.0f} MB — considere dividir em duas mensagens."
        )

    if avisos:
        return ResultadoEtapa.atencao(
            etapa,
            f"{mensagem} Confira: " + " ".join(avisos),
            detalhes={**detalhes, "avisos": avisos},
            artefatos=list(anexos.caminhos),
        )
    return ResultadoEtapa.sucesso(
        etapa, mensagem, detalhes=detalhes, artefatos=list(anexos.caminhos)
    )


# --------------------------------------------------------------------------- #
# Teste manual
# --------------------------------------------------------------------------- #

if __name__ == "__main__":
    import sys
    from datetime import date as _date

    from automacao.nucleo.config import DIR_CONFIG, carregar_ambiente, carregar_contas

    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")

    # Este teste NUNCA cria rascunho: só imprime o texto, para conferir a redação.
    amb = carregar_ambiente()
    registro = DIR_CONFIG / "fornecedores.enriquecido.yaml"
    if not registro.is_file():
        registro = DIR_CONFIG / "fornecedores.yaml"
    todas = {c.id: c for c in carregar_contas(registro)}

    comp = Competencia.de_texto(sys.argv[1]) if len(sys.argv) > 1 else Competencia.atual()

    alvo = todas.get("claro-guarulhos")
    if alvo:
        print("=" * 70)
        print("PRÉVIA INDIVIDUAL — claro-guarulhos")
        print("=" * 70)
        p = previa(alvo, comp, valor=184.90, vencimento=comp.dia_vencimento(20), ambiente_=amb)
        for chave in ("para", "copia", "assunto"):
            print(f"{chave.upper():9}: {p[chave]}")
        print("-" * 70)
        print(p["corpo"])

    grupo = [
        {
            "conta": todas[cid],
            "pdf": None,
            "valor": vl,
            "vencimento": _date(comp.ano, comp.mes, 20),
        }
        for cid, vl in (
            ("algar-ddg-executivo", 2398.99),
            ("algar-internet-link", 1955.90),
            ("algar-internet-link-multimidia", 1521.04),
            ("algar-voz-total", 434.87),
        )
        if cid in todas
    ]
    if grupo:
        print("=" * 70)
        print("PRÉVIA AGRUPADA — as 4 contas do grupo")
        print("=" * 70)
        g = previa_agrupada(grupo, comp, ambiente_=amb)
        for chave in ("para", "copia", "assunto"):
            print(f"{chave.upper():9}: {g[chave]}")
        print("-" * 70)
        print(g["corpo"])
        if g["avisos"]:
            print("AVISOS:", g["avisos"])
