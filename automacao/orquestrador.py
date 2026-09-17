"""
Orquestrador do pipeline.

Encadeia as etapas de uma conta em uma competência:

    coleta → pasta → autorizacao → pdf_autorizacao → pdf_final
           → publicacao → email → checklist

Duas travas de projeto, ambas exigência direta do usuário:

* As etapas que escrevem fora de `dados/` (PUBLICACAO, EMAIL, CHECKLIST)
  só rodam com `confirmado=True`. Sem isso, devolvem o plano detalhado e param.
* O e-mail nunca é enviado — no máximo vira rascunho no Outlook.

Os módulos pesados são importados sob demanda: assim o painel abre mesmo se
o Excel ou o Outlook não estiverem disponíveis na máquina.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from datetime import date
from functools import cached_property
from pathlib import Path

from automacao.nucleo import estado
from automacao.nucleo.config import Ambiente, ambiente, conta_por_id
from automacao.nucleo.modelos import (
    ORDEM_ETAPAS_PADRAO,
    Competencia,
    Conta,
    Documento,
    Etapa,
    OrigemDocumento,
    Processamento,
    ResultadoEtapa,
    Situacao,
    TipoDocumento,
)

log = logging.getLogger("automacao.orquestrador")

# Etapas que só executam depois de você confirmar na tela.
#
# O e-mail saiu desta lista em 01/09/2026, a pedido do usuário: agora um
# clique já despacha. Publicação e checklist continuam pedindo confirmação —
# são as que escrevem no OneDrive e na planilha compartilhada, onde um engano
# atinge arquivo de outra pessoa. O e-mail ganhou proteção de outro tipo, que
# a confirmação não dava: `etapa_email` se recusa a mandar duas vezes.
ETAPAS_SENSIVEIS = {Etapa.PUBLICACAO, Etapa.CHECKLIST}

ORDEM_PADRAO = ORDEM_ETAPAS_PADRAO


@dataclass
class Contexto:
    """O que circula entre as etapas de uma execução."""

    conta: Conta
    competencia: Competencia
    ambiente: Ambiente
    processamento: Processamento
    # Confirmações concedidas pelo usuário nesta execução.
    confirmadas: set[Etapa] = field(default_factory=set)
    # Valores informados/corrigidos no painel sobrepõem o que foi extraído.
    valor_manual: float | None = None
    vencimento_manual: date | None = None
    numero_documento_manual: str | None = None
    descricao_manual: str | None = None
    pagante_manual: str | None = None
    beneficiario_manual: str | None = None
    meio_pgto_manual: str | None = None
    forma_pgto_manual: str | None = None
    abrir_email: bool = True
    # Destinatários/assunto/corpo editados no painel. `None` = usa o
    # settings.yaml; string (mesmo vazia) = o que você digitou, só nesta
    # mensagem — a configuração fica intacta.
    email_para: str | None = None
    email_copia: str | None = None
    email_assunto: str | None = None
    email_corpo: str | None = None
    # Quem assina e com que molde. Vêm do cadastro de quem está logado
    # (Configuração › E-mail), não do settings.yaml: a assinatura é da pessoa,
    # e numa máquina compartilhada duas pessoas assinam diferente.
    # Vazio é ausência de dado, não escolha — cai no global.
    email_assinatura: str = ""
    email_corpo_pessoal: str = ""

    @property
    def pasta_trabalho(self) -> Path:
        return self.ambiente.pasta_trabalho(self.conta.id, self.competencia)

    def autorizado(self, etapa: Etapa) -> bool:
        return etapa not in ETAPAS_SENSIVEIS or etapa in self.confirmadas

    @cached_property
    def valor_do_mes_anterior(self) -> tuple[float, Competencia] | None:
        """
        Último valor pago desta conta, e de que mês ele veio.

        Consulta preguiçosa e memorizada de propósito: na maioria das execuções
        o boleto trouxe o valor e ninguém precisa disto; quando precisa, é uma
        linha só, e `valor` é lido várias vezes durante a etapa.
        """
        return estado.valor_anterior(self.conta.id, self.competencia)

    @property
    def _valor_conhecido(self) -> float | None:
        """O valor de verdade: digitado por você, já gravado, ou lido do PDF."""
        if self.valor_manual is not None:
            return self.valor_manual
        if self.processamento.valor is not None:
            return self.processamento.valor
        achado = self.valor_do_documento
        return achado[0] if achado is not None else None

    @property
    def valor(self) -> float | None:
        """
        Valor da fatura, do mais confiável para o menos.

        Esgotado o que é fato, entra palpite — para a autorização não sair
        zerada —, e a etapa avisa de onde ele veio. Entre os dois palpites, o
        último valor pago vem antes do número do registro porque acompanha a
        realidade sozinho: reajuste, linha a mais, licença a menos. O registro
        só muda quando alguém lembra de revisar as contas uma a uma, e por isso
        fica para a conta nova, que ainda não tem mês nenhum atrás dela.
        """
        conhecido = self._valor_conhecido
        if conhecido is not None:
            return conhecido
        anterior = self.valor_do_mes_anterior
        if anterior is not None:
            return anterior[0]
        return self.conta.valor_estimado

    @property
    def valor_do_documento(self) -> tuple[float, Documento] | None:
        """
        Valor extraído dos documentos, com o boleto na frente.

        É o boleto que manda: o demonstrativo traz subtotais por linha e a nota
        fiscal pode cobrir só parte do que se paga. Pegar "o primeiro documento
        da lista" dava o valor errado quando o demonstrativo vinha antes.
        """
        return _de_documentos(self.processamento, lambda d: d.valor)

    @property
    def valor_veio_do_estimado(self) -> bool:
        """O valor mostrado é palpite — venha ele de onde vier."""
        return self._valor_conhecido is None and self.valor is not None

    @property
    def origem_do_palpite(self) -> str:
        """
        De onde o palpite saiu, em uma frase.

        O aviso precisa dizer isso: "confira o valor" sem dizer contra o quê
        conferir é o tipo de alerta que se aprende a ignorar.
        """
        anterior = self.valor_do_mes_anterior
        if anterior is not None:
            return f"valor pago em {anterior[1].extenso}"
        return "valor aproximado do registro"

    @property
    def vencimento(self) -> date | None:
        if self.vencimento_manual is not None:
            return self.vencimento_manual
        if self.processamento.vencimento is not None:
            return self.processamento.vencimento
        achado = self.vencimento_do_documento
        if achado is not None:
            return achado[0]
        if self.conta.vencimento_dia:
            return self.competencia.dia_vencimento(self.conta.vencimento_dia)
        return None

    @property
    def vencimento_do_documento(self) -> tuple[date, Documento] | None:
        """Vencimento extraído dos documentos — boleto primeiro, pelo mesmo motivo."""
        return _de_documentos(self.processamento, lambda d: d.vencimento)

    @property
    def numero_documento(self) -> str | None:
        if self.numero_documento_manual:
            return self.numero_documento_manual
        if self.processamento.numero_documento:
            return self.processamento.numero_documento
        achado = self.numero_do_documento
        return achado[0] if achado else None

    @property
    def numero_do_documento(self) -> tuple[str, Documento] | None:
        """Nº do documento — aqui a nota fiscal vem antes do boleto."""
        return _de_documentos(
            self.processamento,
            lambda d: d.numero_documento or None,
            ordem=(TipoDocumento.NOTA_FISCAL, TipoDocumento.BOLETO),
        )

    @property
    def descricao(self) -> str | None:
        """
        Texto das OBSERVAÇÕES, do mais recente para o mais antigo.

        Sem nada seu, cai na sugestão — o molde do registro com o mês acertado
        para esta competência (`autorizacao.sugerir_descricao`).
        """
        if self.descricao_manual is not None:
            return self.descricao_manual
        if self.processamento.descricao is not None:
            return self.processamento.descricao
        from automacao.documentos import autorizacao

        texto, _como = autorizacao.sugerir_descricao(self.conta, self.competencia)
        return texto

    @property
    def pagante(self) -> str | None:
        return self.pagante_manual or self.processamento.pagante or self.conta.autorizacao.pagante

    @property
    def beneficiario(self) -> str | None:
        return (
            self.beneficiario_manual
            or self.processamento.beneficiario
            or self.conta.autorizacao.beneficiario
        )

    @property
    def meio_pgto(self) -> str | None:
        return self.meio_pgto_manual or self.processamento.meio_pgto or self.conta.autorizacao.meio_pgto

    @property
    def forma_pgto(self) -> str | None:
        return self.forma_pgto_manual or self.processamento.forma_pgto or self.conta.autorizacao.forma_pgto

    def nome_base(self) -> str:
        return self.conta.nome_base_arquivo(self.competencia)


def _primeiro(valores):
    for v in valores:
        if v is not None:
            return v
    return None


# Boleto na frente: é o documento que diz quanto e quando pagar.
ORDEM_DE_CONFIANCA = (
    TipoDocumento.BOLETO,
    TipoDocumento.NOTA_FISCAL,
    TipoDocumento.DEMONSTRATIVO,
)


def _de_documentos(processamento, campo, ordem=ORDEM_DE_CONFIANCA):
    """
    Primeiro documento que tem o campo preenchido, na ordem de confiança.

    Devolve `(valor, documento)` — o painel mostra de onde veio, para você
    conferir sem abrir o PDF.
    """
    vistos: list[Documento] = []
    for tipo in ordem:
        vistos.extend(processamento.documentos_por_tipo(tipo))
    # Os que sobraram (tipo desconhecido, outros) entram no fim. Comparação por
    # identidade: dois documentos iguais campo a campo não são o mesmo arquivo.
    ja = {id(d) for d in vistos}
    vistos.extend(d for d in processamento.documentos if id(d) not in ja)

    for doc in vistos:
        achado = campo(doc)
        if achado is not None:
            return achado, doc
    return None


def _reais(valor: float) -> str:
    """
    1234.5 -> 'R$ 1.234,50'

    Formata só o número. Aplicar a troca de separadores na frase inteira
    transformava os pontos finais em vírgulas.
    """
    return "R$ " + f"{valor:,.2f}".replace(",", "·").replace(".", ",").replace("·", ".")


# --------------------------------------------------------------------------- #
# Etapas
# --------------------------------------------------------------------------- #


def etapa_coleta(ctx: Contexto) -> ResultadoEtapa:
    """Junta os documentos: Outlook + o que o operador escolheu + o que já veio."""
    from automacao.coleta import classificador

    encontrados: list[Documento] = []
    notas: list[str] = []

    # 1) Outlook (só se a conta tem regras de busca configuradas)
    if ctx.conta.coleta.remetentes or ctx.conta.coleta.assunto_contem:
        try:
            from automacao.coleta import outlook

            r = outlook.buscar(
                ctx.conta, ctx.competencia, ambiente_=ctx.ambiente
            )
            encontrados.extend(r.detalhes.get("documentos") or [])
            notas.append(f"Outlook: {r.mensagem}")
        except Exception as exc:  # Outlook indisponível não pode derrubar o fluxo
            log.warning("coleta no Outlook falhou: %s", exc)
            notas.append(f"Outlook indisponível ({exc})")
    else:
        notas.append("Outlook: sem regras de busca para esta conta")

    # 2) Não existe mais um segundo lugar de onde puxar documento.
    #
    # Existia: `entrada/<conta_id>/`, varrida inteira a cada coleta. Como a
    # pasta era a MESMA para todos os meses, o que ninguém recolheu de um mês
    # entrou no seguinte — em 14/09/2026 uma autorização saiu com o valor, o
    # vencimento e a nota fiscal do mês anterior. A pasta declarava a conta,
    # nunca a competência, e deduzir competência pela data lida do PDF seria
    # palpite sobre palpite.
    #
    # Hoje o operador manda os arquivos da fatura na própria etapa 1 e eles são
    # gravados direto em `trabalho/<competencia>/<conta_id>/` — que é o passo 3
    # aqui embaixo. Um mês não enxerga o outro porque não há pasta em comum.
    notas.append("Envio pela tela: grava direto na pasta desta competência")

    # 3) A pasta desta competência: o que o operador enviou, e o que uma
    #    execução interrompida deixou pelo caminho.
    ja_vistos = {d.caminho.resolve() for d in encontrados}
    nome_base = ctx.nome_base()
    for arq in sorted(ctx.pasta_trabalho.iterdir()):
        if not arq.is_file() or arq.resolve() in ja_vistos:
            continue
        if arq.suffix.lower() not in (".pdf", ".xml"):
            continue
        # Nada que a própria automação gerou entra como documento de entrada:
        # nem o PDF final, nem o PDF intermediário da autorização. Sem isso, a
        # autorização do mês passado voltava para dentro do PDF novo.
        if arq.stem.startswith(nome_base):
            continue
        doc = classificador.classificar(arq, contexto=ctx.conta.rotulo)
        doc.origem = OrigemDocumento.JA_NA_PASTA
        encontrados.append(doc)

    ctx.processamento.documentos = _deduplicar(encontrados)
    estado.salvar_documentos(ctx.conta.id, ctx.competencia, ctx.processamento.documentos)

    if not ctx.processamento.documentos:
        return ResultadoEtapa.erro(
            Etapa.COLETA,
            "nenhum documento nesta competência. Suba o boleto, a nota "
            "fiscal e o demonstrativo desta fatura em “Enviar e coletar”.",
            detalhes={"notas": notas},
        )

    duvidosos = [d for d in ctx.processamento.documentos if d.confianca < 0.65]
    tem_boleto = bool(ctx.processamento.documentos_por_tipo(TipoDocumento.BOLETO))

    if duvidosos or not tem_boleto:
        pendencias = []
        if duvidosos:
            pendencias.append(
                f"{len(duvidosos)} documento(s) com classificação incerta"
            )
        if not tem_boleto:
            pendencias.append("nenhum boleto identificado")
        return ResultadoEtapa.atencao(
            Etapa.COLETA,
            f"{len(ctx.processamento.documentos)} documento(s) — " + "; ".join(pendencias),
            detalhes={
                "notas": notas,
                "documentos": ctx.processamento.documentos,
                "duvidosos": [d.nome for d in duvidosos],
            },
        )

    return ResultadoEtapa.sucesso(
        Etapa.COLETA,
        f"{len(ctx.processamento.documentos)} documento(s) coletado(s)",
        detalhes={"notas": notas, "documentos": ctx.processamento.documentos},
    )


def _impressao(caminho: Path) -> str:
    """Identidade do arquivo pelo conteúdo, não pelo nome."""
    import hashlib

    try:
        h = hashlib.sha256()
        with caminho.open("rb") as fh:
            for bloco in iter(lambda: fh.read(1 << 18), b""):
                h.update(bloco)
        return h.hexdigest()
    except OSError:
        return f"caminho::{caminho}"


def _deduplicar(documentos: list[Documento]) -> list[Documento]:
    """
    Mesmo arquivo vindo por dois caminhos: fica o de maior confiança.

    A comparação é pelo CONTEÚDO. O mesmo boleto chega como
    'Boleto.pdf' pelo Outlook e como 'Boleto_2.pdf' pela pasta de entrada —
    comparar por nome deixava os dois passarem e o PDF final saía com a
    página repetida.
    """
    por_conteudo: dict[str, Documento] = {}
    for d in documentos:
        chave = _impressao(d.caminho)
        atual = por_conteudo.get(chave)
        if atual is None or d.confianca > atual.confianca:
            por_conteudo[chave] = d
    return list(por_conteudo.values())


def etapa_pasta(ctx: Contexto) -> ResultadoEtapa:
    """Confere o destino no OneDrive. NÃO cria nada — quem cria é a publicação."""
    destino = ctx.ambiente.destino_onedrive(ctx.conta, ctx.competencia)
    raiz_forn = ctx.ambiente.pasta_do_fornecedor(ctx.conta)

    if not raiz_forn.is_dir():
        return ResultadoEtapa.erro(
            Etapa.PASTA,
            f"a pasta do fornecedor não existe no OneDrive: {raiz_forn}",
            detalhes={"pasta_fornecedor": str(raiz_forn)},
        )

    if destino.is_dir():
        existentes = sorted(p.name for p in destino.iterdir() if p.is_file())
        return ResultadoEtapa.atencao(
            Etapa.PASTA,
            f"a pasta do mês já existe com {len(existentes)} arquivo(s)"
            if existentes
            else "a pasta do mês já existe (vazia)",
            detalhes={"destino": str(destino), "arquivos_existentes": existentes},
        )

    # Caminho curto na mensagem, relativo à raiz que ESTA conta usa. Fixar a
    # raiz padrão aqui estourava com quem arquiva em outra árvore.
    raiz = ctx.ambiente.raiz_de(ctx.conta)
    try:
        curto = destino.relative_to(raiz)
    except ValueError:
        curto = destino
    return ResultadoEtapa.sucesso(
        Etapa.PASTA,
        f"pasta a criar: {curto} (em {raiz.name})",
        detalhes={"destino": str(destino), "sera_criada": True},
    )


def etapa_autorizacao(ctx: Contexto) -> ResultadoEtapa:
    from automacao.documentos import autorizacao

    if not ctx.conta.modelo_base:
        return ResultadoEtapa.erro(
            Etapa.AUTORIZACAO,
            "a conta não tem 'modelo_base' no registro — sem ele não dá para "
            "saber qual planilha usar de base.",
        )

    destino = ctx.pasta_trabalho / f"{ctx.nome_base()}.xlsx"
    resultado = autorizacao.preencher(
        ctx.conta,
        ctx.competencia,
        valor=ctx.valor,
        vencimento=ctx.vencimento,
        numero_documento=ctx.numero_documento,
        descricao=ctx.descricao,
        pagante=ctx.pagante,
        beneficiario=ctx.beneficiario,
        meio_pgto=ctx.meio_pgto,
        forma_pgto=ctx.forma_pgto,
        destino=destino,
        ambiente_=ctx.ambiente,
    )
    # Aviso de procedência do valor. O alerta de variação (">30% do
    # habitual") foi retirado a pedido: gerava ruído em contas cujo valor
    # oscila de um mês para outro por natureza — telefonia por consumo,
    # licença por número de usuários.
    if ctx.valor_veio_do_estimado:
        resultado.detalhes.setdefault("conferir", []).append(
            f"o valor R$ {ctx.valor:.2f} é o {ctx.origem_do_palpite}, não veio "
            "do boleto. Confirme antes de publicar."
        )
        if resultado.situacao is Situacao.OK:
            resultado.situacao = Situacao.ATENCAO

    if resultado.ok and resultado.artefatos:
        ctx.processamento.xlsx = resultado.artefatos[0]
        estado.salvar_dados(
            ctx.conta.id,
            ctx.competencia,
            xlsx=ctx.processamento.xlsx,
            valor=ctx.valor,
            vencimento=ctx.vencimento,
            numero_documento=ctx.numero_documento,
            descricao=ctx.descricao,
            pagante=ctx.pagante,
            beneficiario=ctx.beneficiario,
            meio_pgto=ctx.meio_pgto,
            forma_pgto=ctx.forma_pgto,
        )
        _invalidar_pdfs(ctx)
    return resultado


def _invalidar_pdfs(ctx: Contexto) -> None:
    """
    Reabre a exportação e a montagem do PDF depois de a planilha mudar.

    Sem isto, corrigir um campo na etapa 3 deixava as etapas 4 e 5 marcadas
    como concluídas apontando para PDFs gerados ANTES da correção — e era
    esse PDF velho que ia para o OneDrive e para o financeiro. Aconteceu numa
    conta real: o beneficiário foi trocado na planilha e o PDF continuou
    trazendo o antigo.

    Reabrir estas duas tranca as seguintes (a trilha não deixa pular etapa),
    então publicação, e-mail e checklist esperam o PDF novo. O registro de
    despacho do e-mail NÃO é tocado: a trava contra segunda via continua de pé.
    """
    for etapa in (Etapa.PDF_AUTORIZACAO, Etapa.PDF_FINAL):
        if etapa in ctx.processamento.resultados:
            del ctx.processamento.resultados[etapa]
            estado.reabrir(ctx.conta.id, ctx.competencia, etapa)
            log.info(
                "[%s/%s] %s reaberta: a planilha mudou",
                ctx.conta.id, ctx.competencia, etapa.value,
            )


def etapa_pdf_autorizacao(ctx: Contexto) -> ResultadoEtapa:
    from automacao.documentos import exportar_pdf

    xlsx = ctx.processamento.xlsx
    if not xlsx or not xlsx.is_file():
        return ResultadoEtapa.erro(
            Etapa.PDF_AUTORIZACAO, "a planilha de autorização ainda não foi gerada."
        )

    # PDF intermediário: fica com sufixo para não colidir com o PDF final,
    # que leva o nome limpo.
    destino = ctx.pasta_trabalho / f"{ctx.nome_base()}__autorizacao.pdf"
    return exportar_pdf.exportar(xlsx, ctx.conta.aba_autorizacao, destino)


def etapa_pdf_final(ctx: Contexto) -> ResultadoEtapa:
    from automacao.documentos import montar_pdf

    pdf_autorizacao = ctx.processamento.resultados.get(Etapa.PDF_AUTORIZACAO)
    if not pdf_autorizacao or not pdf_autorizacao.artefatos:
        return ResultadoEtapa.erro(
            Etapa.PDF_FINAL, "o PDF da autorização ainda não foi exportado."
        )

    doc_autorizacao = Documento(
        caminho=pdf_autorizacao.artefatos[0],
        tipo=TipoDocumento.AUTORIZACAO,
        confianca=1.0,
        motivo="gerado pela automação",
    )
    entrada = [doc_autorizacao] + [
        d
        for d in ctx.processamento.documentos
        if d.tipo not in (TipoDocumento.AUTORIZACAO, TipoDocumento.XML_NFE)
    ]

    # Regra do usuário: o PDF final tem o MESMO nome do xlsx, para não
    # duplicar arquivo no servidor.
    destino = ctx.pasta_trabalho / f"{ctx.nome_base()}.pdf"
    resultado = montar_pdf.montar(entrada, destino, ctx.conta.id)

    if resultado.ok and resultado.artefatos:
        ctx.processamento.pdf_final = resultado.artefatos[0]
        estado.salvar_dados(
            ctx.conta.id, ctx.competencia, pdf_final=ctx.processamento.pdf_final
        )
    return resultado


def etapa_publicacao(ctx: Contexto) -> ResultadoEtapa:
    from automacao.entrega import publicador

    arquivos: list[Path] = []
    if ctx.processamento.xlsx and ctx.processamento.xlsx.is_file():
        arquivos.append(ctx.processamento.xlsx)
    if ctx.processamento.pdf_final and ctx.processamento.pdf_final.is_file():
        arquivos.append(ctx.processamento.pdf_final)
    # Os documentos de origem também vão para a pasta, como sempre foi feito.
    for d in ctx.processamento.documentos:
        if d.caminho.is_file() and d.origem is not OrigemDocumento.JA_NA_PASTA:
            arquivos.append(d.caminho)

    plano = publicador.planejar(
        ctx.conta, ctx.competencia, arquivos, ambiente_=ctx.ambiente
    )
    return publicador.publicar(
        plano,
        confirmado=ctx.autorizado(Etapa.PUBLICACAO),
        ambiente_=ctx.ambiente,
    )


def _mensagem_ja_saiu(proc: Processamento) -> str | None:
    """
    Por que esta fatura não deve ser enviada de novo — ou `None` se pode.

    Dois sinais, e o segundo importa tanto quanto o primeiro:

    * `email_enviado` — gravado no despacho e no registro de "enviei por fora";
    * a etapa de e-mail já ter fechado com sucesso. Cobre o mês inteiro que
      foi feito quando a ferramenta só rascunhava: ali `email_enviado` é 0,
      mas o rascunho virou mensagem enviada pela mão de alguém. Sem esta
      segunda checagem, ligar o envio automático transformaria um reprocesso
      dessas contas em segunda via para o financeiro.
    """
    if proc.email_enviado:
        quando = proc.enviado_em
        onde = "por fora da ferramenta" if proc.enviado_por_fora else "por aqui"
        return (
            f"esta fatura já foi enviada {onde}"
            + (f" em {quando:%d/%m/%Y}." if quando else ".")
        )

    anterior = proc.resultados.get(Etapa.EMAIL)
    if anterior and anterior.situacao in (Situacao.OK, Situacao.ATENCAO):
        quando = anterior.atualizado_em
        return (
            "a etapa de e-mail desta fatura já foi concluída"
            + (f" em {quando:%d/%m/%Y}" if quando else "")
            + " — a mensagem saiu ou está no seu Rascunhos."
        )
    return None


def etapa_email(ctx: Contexto) -> ResultadoEtapa:
    from automacao.entrega import email_outlook

    # O anexo é o PDF já publicado, se houver; senão, o da área local.
    destino_onedrive = ctx.ambiente.destino_onedrive(ctx.conta, ctx.competencia)
    pdf = destino_onedrive / f"{ctx.nome_base()}.pdf"
    if not pdf.is_file():
        pdf = ctx.processamento.pdf_final or (ctx.pasta_trabalho / f"{ctx.nome_base()}.pdf")

    if not pdf.is_file():
        return ResultadoEtapa.erro(Etapa.EMAIL, "o PDF final ainda não existe.")

    # Trava de repetição. Um rascunho a mais se apaga; um e-mail a mais já
    # chegou. Duplo clique, F5 no POST e reexecução da fila passam por aqui —
    # e nos registros deste projeto a etapa de e-mail já rodou duas vezes
    # seguidas na mesma conta.
    if email_outlook.deve_enviar(ctx.ambiente):
        ja = _mensagem_ja_saiu(ctx.processamento)
        if ja:
            return ResultadoEtapa.erro(
                Etapa.EMAIL,
                f"{ja} Não vou mandar de novo. Se precisa mesmo reenviar, use "
                "'Reabrir esta etapa' — aí a trava sai.",
            )

    # Não é mais trava de aprovação — o e-mail saiu de ETAPAS_SENSIVEIS.
    # É a diferença entre os dois botões do formulário: "Atualizar prévia"
    # chega sem confirmação e devolve o texto; "Enviar" chega com ela e fala
    # com o Outlook. Ler `confirmadas` direto (em vez de `autorizado`, que
    # agora sempre devolve True aqui) é o que impede a prévia de despachar.
    if Etapa.EMAIL not in ctx.confirmadas:
        previa = email_outlook.previa(
            ctx.conta,
            ctx.competencia,
            valor=ctx.valor,
            vencimento=ctx.vencimento,
            ambiente_=ctx.ambiente,
            para=ctx.email_para,
            copia=ctx.email_copia,
            assunto=ctx.email_assunto,
            corpo=ctx.email_corpo,
            assinatura=ctx.email_assinatura,
            corpo_pessoal=ctx.email_corpo_pessoal,
        )
        situacao = Situacao.PENDENTE
        mensagem = (
            "pronto para enviar — confira a prévia"
            if email_outlook.deve_enviar(ctx.ambiente)
            else "rascunho pronto para ser criado — confira a prévia"
        )
        if previa["enderecos_suspeitos"]:
            situacao = Situacao.ERRO
            mensagem = "endereço com formato inválido: " + ", ".join(
                previa["enderecos_suspeitos"]
            )
        elif not previa["para_enderecos"]:
            situacao = Situacao.ERRO
            mensagem = "o campo Para está vazio — informe pelo menos um destinatário."
        return ResultadoEtapa(
            etapa=Etapa.EMAIL,
            situacao=situacao,
            mensagem=mensagem,
            detalhes={**previa, "anexo": pdf.name, "anexo_caminho": str(pdf)},
        )

    resultado = email_outlook.criar_rascunho(
        ctx.conta,
        ctx.competencia,
        pdf=pdf,
        valor=ctx.valor,
        vencimento=ctx.vencimento,
        ambiente_=ctx.ambiente,
        para=ctx.email_para,
        copia=ctx.email_copia,
        assunto=ctx.email_assunto,
        corpo=ctx.email_corpo,
        assinatura=ctx.email_assinatura,
        corpo_pessoal=ctx.email_corpo_pessoal,
    )

    # Só quando saiu de verdade. Rascunho não conta como envio: a mensagem
    # ainda depende de alguém clicar no Outlook.
    if resultado.detalhes.get("enviada"):
        ctx.processamento.email_enviado = True
        ctx.processamento.enviado_em = date.today()
        estado.registrar_envio_pela_ferramenta(
            ctx.conta.id, ctx.competencia, date.today()
        )
    return resultado


def etapa_checklist(ctx: Contexto) -> ResultadoEtapa:
    from automacao.entrega import planilha_contas

    return planilha_contas.marcar_enviado(
        ctx.conta,
        ctx.competencia,
        ambiente_=ctx.ambiente,
        confirmado=ctx.autorizado(Etapa.CHECKLIST),
    )


EXECUTORES = {
    Etapa.COLETA: etapa_coleta,
    Etapa.PASTA: etapa_pasta,
    Etapa.AUTORIZACAO: etapa_autorizacao,
    Etapa.PDF_AUTORIZACAO: etapa_pdf_autorizacao,
    Etapa.PDF_FINAL: etapa_pdf_final,
    Etapa.PUBLICACAO: etapa_publicacao,
    Etapa.EMAIL: etapa_email,
    Etapa.CHECKLIST: etapa_checklist,
}


# --------------------------------------------------------------------------- #
# Execução
# --------------------------------------------------------------------------- #


def montar_contexto(
    conta: Conta | str,
    competencia: Competencia,
    *,
    confirmadas: set[Etapa] | None = None,
    valor: float | None = None,
    vencimento: date | None = None,
    numero_documento: str | None = None,
    descricao: str | None = None,
    pagante: str | None = None,
    beneficiario: str | None = None,
    meio_pgto: str | None = None,
    forma_pgto: str | None = None,
    ambiente_: Ambiente | None = None,
    email_para: str | None = None,
    email_copia: str | None = None,
    email_assunto: str | None = None,
    email_corpo: str | None = None,
    email_assinatura: str = "",
    email_corpo_pessoal: str = "",
) -> Contexto:
    alvo = conta if isinstance(conta, Conta) else conta_por_id(conta)
    return Contexto(
        conta=alvo,
        competencia=competencia,
        ambiente=ambiente_ or ambiente(),
        processamento=estado.carregar(alvo.id, competencia),
        confirmadas=set(confirmadas or ()),
        valor_manual=valor,
        vencimento_manual=vencimento,
        numero_documento_manual=numero_documento,
        descricao_manual=descricao,
        pagante_manual=pagante,
        beneficiario_manual=beneficiario,
        meio_pgto_manual=meio_pgto,
        forma_pgto_manual=forma_pgto,
        email_para=email_para,
        email_copia=email_copia,
        email_assunto=email_assunto,
        email_corpo=email_corpo,
        email_assinatura=email_assinatura,
        email_corpo_pessoal=email_corpo_pessoal,
    )


def executar_etapa(ctx: Contexto, etapa: Etapa) -> ResultadoEtapa:
    """Roda uma etapa isolada, tratando exceção como erro registrado."""
    executor = EXECUTORES[etapa]
    log.info("[%s/%s] %s", ctx.conta.id, ctx.competencia, etapa.rotulo)
    try:
        resultado = executor(ctx)
    except ImportError as exc:
        resultado = ResultadoEtapa.erro(
            etapa, f"módulo da etapa ainda não disponível: {exc}"
        )
    except Exception as exc:
        log.exception("etapa %s falhou", etapa.value)
        resultado = ResultadoEtapa.erro(etapa, f"{type(exc).__name__}: {exc}")

    ctx.processamento.registrar(resultado)
    estado.salvar_etapa(ctx.conta.id, ctx.competencia, resultado)
    recolher_entrada_se_concluiu(
        ctx.conta,
        ctx.competencia,
        processamento=ctx.processamento,
        ambiente_=ctx.ambiente,
    )
    return resultado


def recolher_entrada_se_concluiu(
    conta: Conta | str,
    competencia: Competencia,
    *,
    processamento: Processamento | None = None,
    ambiente_: Ambiente | None = None,
) -> None:
    """
    Fecharam as oito etapas: tira da entrada o que já está na pasta de destino.

    Fica aqui, e não dentro de `etapa_checklist`, porque "concluído" é o
    conjunto das oito — a última a fechar pode ser outra, quando você reabre
    uma etapa do meio e reprocessa, ou quando marca uma como resolvida por
    fora pelo painel. O gatilho é o estado, não a ordem.

    Nada acontece sem prova: a limpeza confere o SHA-256 de cada arquivo na
    pasta do OneDrive antes de apagar (ver `manutencao/limpeza.py`). E ela
    nunca derruba o fluxo — a fatura já saiu; sobrar arquivo na bancada é
    incômodo, perder a etapa por causa disso seria dano.
    """
    alvo = conta if isinstance(conta, Conta) else conta_por_id(conta)
    proc = processamento or estado.carregar(alvo.id, competencia)
    if not proc.concluido:
        return

    try:
        from automacao.manutencao import limpeza

        resultado = limpeza.limpar_entrada_da_conta(
            alvo, competencia, ambiente_=ambiente_ or ambiente()
        )
    except Exception as exc:  # noqa: BLE001 — limpeza nunca derruba o pipeline
        log.warning("não consegui recolher a entrada de %s: %s", alvo.id, exc)
        return

    if resultado.apagados:
        log.info(
            "[%s/%s] entrada recolhida: %d arquivo(s), %.1f MB",
            alvo.id,
            competencia,
            len(resultado.apagados),
            resultado.bytes_liberados / 1024 / 1024,
        )


def executar(
    ctx: Contexto,
    *,
    etapas: list[Etapa] | None = None,
    parar_no_erro: bool = True,
) -> Processamento:
    """Roda a sequência de etapas. Para no primeiro erro, por padrão."""
    for etapa in etapas or ORDEM_PADRAO:
        resultado = executar_etapa(ctx, etapa)

        if resultado.situacao is Situacao.ERRO and parar_no_erro:
            log.warning("parando em %s: %s", etapa.value, resultado.mensagem)
            break
        # Etapa sensível sem confirmação: para e devolve o plano para o painel.
        if resultado.situacao is Situacao.PENDENTE:
            log.info("aguardando confirmação em %s", etapa.value)
            break

    return ctx.processamento


def preparar(
    conta: Conta | str, competencia: Competencia, **kw
) -> Processamento:
    """Roda só as etapas que NÃO tocam o mundo externo. Sempre seguro."""
    ctx = montar_contexto(conta, competencia, **kw)
    return executar(
        ctx,
        etapas=[
            Etapa.COLETA,
            Etapa.PASTA,
            Etapa.AUTORIZACAO,
            Etapa.PDF_AUTORIZACAO,
            Etapa.PDF_FINAL,
        ],
    )
