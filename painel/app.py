"""
Painel web local da Automação Financeira.

Roda em http://127.0.0.1:8000 — só na sua máquina, sem exposição externa.

O painel é o lugar onde o julgamento humano entra: conferir se o boleto é o
certo, corrigir o valor, revisar o texto do e-mail e autorizar cada escrita.
Nenhuma rota que altera o OneDrive, o Outlook ou a planilha compartilhada
executa sem confirmação explícita vinda de um formulário.
"""

from __future__ import annotations

import logging
import os
import time as _time
from datetime import date, datetime
from html import escape
from pathlib import Path
from dataclasses import asdict
from urllib.parse import quote

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import (
    FileResponse,
    HTMLResponse,
    RedirectResponse,
    Response,
)
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from automacao import configurar_log
from automacao.nucleo import estado
import yaml

from automacao.nucleo.config import (
    DIR_CONFIG,
    beneficiario_por_nome,
    beneficiarios,
    pagante_por_nome,
    pagantes,
    salvar_beneficiarios,
    salvar_pagantes,
    salvar_settings,
    ErroConfiguracao,
    RAIZ_PROJETO,
    ambiente,
    conta_por_id,
    contas,
    recarregar,
)
from automacao.nucleo.modelos import (
    Beneficiario,
    Competencia,
    cnpj_confere,
    formatar_cnpj,
    FORMAS_PAGAMENTO,
    MEIOS_PAGAMENTO,
    Pagante,
    FORMATOS_PASTA_MES,
    Etapa,
    LIMIAR_CONFIANCA,
    ResultadoEtapa,
    Situacao,
    TipoDocumento,
)
from automacao.orquestrador import (
    ETAPAS_SENSIVEIS,
    ORDEM_PADRAO,
    executar,
    executar_etapa,
    montar_contexto,
    recolher_entrada_se_concluiu,
)

log = logging.getLogger("painel")

AQUI = Path(__file__).resolve().parent

# --------------------------------------------------------------------------- #
# Detecção de painel desatualizado
# --------------------------------------------------------------------------- #
#
# O Jinja relê os templates do disco a cada requisição; este módulo fica como
# estava quando o processo subiu. Atualizar o projeto com o painel aberto
# produz tela nova sobre servidor velho — e o sintoma engana: link novo devolve
# {"detail":"Not Found"}, variável nova derruba a página com erro interno.
#
# Comparar a data do arquivo em disco com o momento do import resolve sem
# precisar lembrar de incrementar número de versão nenhum.

# Relógio, não data de arquivo: a pergunta é "algum .py foi salvo depois que
# este processo carregou?". Usar a data do próprio app.py compararia com um
# instante no passado e acusaria desatualização o tempo todo.
MOMENTO_DO_IMPORT = _time.time()
ARQUIVOS_DE_CODIGO = [Path(__file__),
                      *(AQUI.parent / "automacao").rglob("*.py")]


def codigo_mudou_no_disco() -> bool:
    """True quando algum .py foi salvo depois deste processo carregar."""
    for arquivo in ARQUIVOS_DE_CODIGO:
        try:
            if arquivo.stat().st_mtime > MOMENTO_DO_IMPORT:
                return True
        except OSError:
            continue
    return False
app = FastAPI(title="Automação Financeira", docs_url=None, redoc_url=None)
app.mount("/static", StaticFiles(directory=AQUI / "static"), name="static")

#: Nome do cookie de sessão e por onde se entra sem estar logado.
COOKIE_SESSAO = "af_sessao"
SEM_LOGIN = ("/entrar", "/static", "/vivo", "/favicon.ico")


@app.middleware("http")
async def exigir_login(request: Request, proximo):
    """
    Nenhuma tela antes do login; nenhuma tela antes de trocar a senha inicial.

    O painel escuta só em 127.0.0.1, então isto não é defesa de rede — é
    defesa de mesa: a máquina é compartilhada e a tela mostra valor de fatura,
    CNPJ e dado bancário, além de poder despachar autorização.

    A troca de senha bloqueia tudo de propósito. Senha inicial é combinada por
    escrito e previsível; deixar navegar com ela seria não ter trocado nada.
    """
    caminho = request.url.path
    if caminho.startswith(SEM_LOGIN):
        return await proximo(request)

    from automacao.acesso import usuarios

    usuario = usuarios.de_cookie(request.cookies.get(COOKIE_SESSAO))
    if usuario is None:
        destino = request.url.path
        if request.url.query:
            destino += "?" + request.url.query
        return RedirectResponse(f"/entrar?destino={quote(destino)}", status_code=303)

    if usuario.precisa_trocar_senha and caminho != "/trocar-senha":
        return RedirectResponse("/trocar-senha", status_code=303)

    # Esconder o link no menu não protege nada: o endereço continua digitável,
    # e um POST nem precisa de tela. A trava de verdade é aqui.
    from automacao.acesso import permissoes

    if not permissoes.pode_acessar(usuario, caminho):
        porta = permissoes.primeiro_caminho(usuario)
        # Quem caiu na raiz sem ter a lista do mês é mandado para a primeira
        # tela que enxerga — é o que acontece logo depois do login.
        if caminho == "/" and porta and porta != "/":
            return RedirectResponse(porta, status_code=303)
        log.warning("%s tentou abrir %s, fora do grupo %s",
                    usuario.email, caminho, usuario.grupo)
        return _pagina_sem_permissao(usuario, caminho, porta)

    request.state.usuario = usuario
    return await proximo(request)


def _pagina_sem_permissao(usuario, caminho: str, porta: str) -> HTMLResponse:
    """
    Tela de 403 que diz o que fazer, em vez do JSON pelado do FastAPI.

    Não estende base.html: o cabeçalho de lá monta o menu a partir das
    permissões, e uma falha ali deixaria esta página cair junto.
    """
    from automacao.acesso import permissoes

    grupo = permissoes.grupo_de(usuario)
    saida = (f'<a href="{porta}">Voltar para o que você acessa</a>' if porta
             else '<a href="/sair">Sair do painel</a>')
    corpo = f"""<!doctype html><html lang="pt-BR"><head><meta charset="utf-8">
<title>Sem permissão</title><style>
 body{{margin:0;background:#14171a;color:#e6e8ea;
   font:15px/1.6 "Segoe UI",system-ui,sans-serif;padding:2.5rem 1.5rem}}
 .caixa{{max-width:38rem;margin:0 auto;background:#1c2024;border:1px solid #2c3238;
   border-radius:10px;padding:1.4rem 1.6rem}}
 h1{{font-size:1.15rem;margin:0 0 .8rem}}
 code{{background:#23282d;padding:.1rem .35rem;border-radius:4px;font-size:.86rem}}
 a{{color:#6fa8e8}}
 .apoio{{color:#9aa3ab;font-size:.86rem;margin-top:1rem}}
</style></head><body><div class="caixa">
<h1>Esta tela não está liberada para você</h1>
<p style="margin:0 0 .8rem">Seu acesso é do grupo
<strong>{escape(grupo.nome or grupo.chave or 'sem grupo')}</strong>, e
<code>{escape(caminho)}</code> está fora dele.</p>
<p style="margin:0 0 .8rem">{saida}</p>
<p class="apoio">Se você precisa desta tela para trabalhar, peça a um
administrador — ele libera em <strong>Configuração › Permissões</strong>.</p>
</div></body></html>"""
    return HTMLResponse(corpo, status_code=403)


@app.get("/entrar", response_class=HTMLResponse)
def tela_entrar(request: Request, destino: str = "/", erro: str = ""):
    from automacao.acesso import usuarios

    if usuarios.de_cookie(request.cookies.get(COOKIE_SESSAO)):
        return RedirectResponse(destino or "/", status_code=303)
    return templates.TemplateResponse(
        request,
        "entrar.html",
        {
            "request": request,
            "destino": destino or "/",
            "erro": erro,
            "sem_usuarios": not usuarios.carregar(),
            "tem_logo": (AQUI / "static" / "logo.png").is_file(),
            "tem_logo_escura": (AQUI / "static" / "logo-escura.png").is_file(),
            "tem_icone": (AQUI / "static" / "favicon.ico").is_file(),
        },
    )


@app.post("/entrar")
def fazer_login(
    email: str = Form(...), senha: str = Form(...), destino: str = Form("/")
):
    from automacao.acesso import usuarios

    usuario = usuarios.autenticar(email, senha)
    if usuario is None:
        # Uma mensagem só para os dois casos: dizer "este e-mail não existe"
        # entregaria quem tem acesso ao painel.
        return RedirectResponse(
            "/entrar?destino=" + quote(destino or "/")
            + "&erro=" + quote("E-mail ou senha não conferem."),
            status_code=303,
        )

    log.info("entrou no painel: %s", usuario.email)
    resposta = RedirectResponse(
        "/trocar-senha" if usuario.precisa_trocar_senha else (destino or "/"),
        status_code=303,
    )
    resposta.set_cookie(
        COOKIE_SESSAO,
        usuarios.assinar(usuario.email),
        httponly=True,      # fora do alcance de JavaScript
        samesite="lax",
        max_age=60 * 60 * 12,
    )
    return resposta


@app.get("/sair")
def sair(request: Request):
    resposta = RedirectResponse("/entrar", status_code=303)
    resposta.delete_cookie(COOKIE_SESSAO)
    return resposta


@app.get("/trocar-senha", response_class=HTMLResponse)
def tela_trocar_senha(request: Request, erro: str = "", guardado: str = ""):
    from automacao.acesso import usuarios

    usuario = getattr(request.state, "usuario", None)

    # No primeiro acesso a senha atual é a inicial — combinada por escrito e
    # acabada de digitar no login. Pedir que ele a repita aqui só cria chance
    # de errar, então o campo já vem preenchido e travado.
    #
    # Só se mostra o que se confere contra o resumo guardado: se um dia alguém
    # cadastrar com uma provisória diferente da inicial do ano, o campo volta a
    # ser pedido em vez de vir preenchido com o valor errado. E não há segredo
    # novo na tela — é a senha do próprio usuário logado, a mesma que a tela de
    # Usuários já exibe para o administrador.
    senha_atual = ""
    if usuario is not None and usuario.precisa_trocar_senha:
        inicial = usuarios.senha_inicial()
        if usuario.confere(inicial):
            senha_atual = inicial

    return templates.TemplateResponse(
        request,
        "trocar_senha.html",
        {
            **contexto_base(request, competencia_de(request)),
            "usuario": usuario,
            "primeira_vez": bool(usuario and usuario.precisa_trocar_senha),
            "senha_atual": senha_atual,
            "minimo": usuarios.TAMANHO_MINIMO_SENHA,
            "maximo": usuarios.TAMANHO_MAXIMO_SENHA,
            "erro": erro,
            "guardado": guardado,
        },
    )


@app.post("/trocar-senha")
def salvar_senha_nova(
    request: Request,
    atual: str = Form(...),
    nova: str = Form(...),
    repetida: str = Form(...),
):
    from automacao.acesso import usuarios

    usuario = request.state.usuario
    voltar = "/trocar-senha?erro="

    if not usuario.confere(atual):
        return RedirectResponse(voltar + quote("A senha atual não confere."), 303)
    if nova != repetida:
        return RedirectResponse(voltar + quote("As duas senhas novas não batem."), 303)
    try:
        usuarios.trocar_senha(usuario.email, nova)
    except usuarios.ErroUsuario as erro:
        return RedirectResponse(voltar + quote(f"Senha recusada: {erro}"), 303)

    return RedirectResponse("/?trocou=1", status_code=303)

templates = Jinja2Templates(directory=AQUI / "templates")


# --------------------------------------------------------------------------- #
# Filtros e helpers de apresentação
# --------------------------------------------------------------------------- #

# A partir de quantos dias do vencimento a conta entra em alerta no painel.
DIAS_DE_ALERTA = 5

CORES_SITUACAO = {
    Situacao.OK: ("ok", "concluído"),
    Situacao.ATENCAO: ("atencao", "revisar"),
    Situacao.ERRO: ("erro", "erro"),
    Situacao.PENDENTE: ("pendente", "aguardando"),
    Situacao.PULADO: ("pulado", "pulado"),
}


def moeda(valor) -> str:
    if valor in (None, ""):
        return "—"
    inteiro = f"{float(valor):,.2f}"
    return "R$ " + inteiro.replace(",", "·").replace(".", ",").replace("·", ".")


def data_br(valor) -> str:
    if not valor:
        return "—"
    if isinstance(valor, str):
        try:
            valor = date.fromisoformat(valor)
        except ValueError:
            return valor
    return valor.strftime("%d/%m/%Y")


def versao_do_arquivo(caminho) -> str:
    """
    Carimbo que muda quando o arquivo muda — para pôr no fim de um endereço.

    Serve para o navegador não reaproveitar do cache o PDF da geração
    anterior: o caminho do arquivo é sempre o mesmo, então sem isto ele não
    tem como perceber que o conteúdo é outro.
    """
    try:
        return str(int(Path(caminho).stat().st_mtime))
    except OSError:
        return "0"


templates.env.filters["moeda"] = moeda
templates.env.filters["data_br"] = data_br
templates.env.filters["versao_do_arquivo"] = versao_do_arquivo
# O que cada etapa faz, em uma frase, para quem está olhando a tela.
EXPLICACAO_ETAPA = {
    "pasta": (
        "Confere onde os arquivos vão parar no OneDrive. Nada é criado agora — "
        "a pasta nasce só na etapa de publicação, depois da sua confirmação."
    ),
    "pdf_autorizacao": (
        "Abre a planilha no Excel e exporta a aba da autorização em PDF, "
        "respeitando a área de impressão. Leva de 10 a 20 segundos."
    ),
    "pdf_final": (
        "Junta tudo num arquivo só, nesta ordem: autorização, demonstrativo, "
        "boleto e nota fiscal. O PDF recebe o mesmo nome da planilha."
    ),
    "publicacao": (
        "Cria a pasta do mês e copia os arquivos para o OneDrive. Primeiro "
        "mostra a lista do que será feito; só grava depois que você confirmar."
    ),
    "email": (
        "Monta o rascunho no Outlook com destinatário, cópia, assunto e o PDF "
        "anexado, e abre para você conferir. Quem clica em Enviar é você — o "
        "despacho automático não funciona nesta máquina (a mensagem para na "
        "Caixa de Saída do Outlook clássico, que não fica aberto)."
    ),
    "checklist": (
        "Marca a caixa \"Enviou para Financeiro?\" na linha desta conta, na aba "
        "do mês da planilha CONTAS E ACESSOS. É a caixa que deixa a linha "
        "verde — a planilha faz isso por formatação condicional."
    ),
}

templates.env.globals.update(
    DIAS_DE_ALERTA=DIAS_DE_ALERTA,
    ETAPAS=list(ORDEM_PADRAO),
    ETAPAS_SENSIVEIS=ETAPAS_SENSIVEIS,
    CORES_SITUACAO=CORES_SITUACAO,
    EXPLICACAO_ETAPA=EXPLICACAO_ETAPA,
    Situacao=Situacao,
    Etapa=Etapa,
    LIMIAR_CONFIANCA=LIMIAR_CONFIANCA,
    FORMAS_PAGAMENTO=FORMAS_PAGAMENTO,
    MEIOS_PAGAMENTO=MEIOS_PAGAMENTO,
)


@app.exception_handler(Exception)
def pagina_de_erro(request: Request, exc: Exception) -> HTMLResponse:
    """
    Transforma o "Internal Server Error" pelado numa tela que explica.

    O 500 do uvicorn é uma linha de texto sem contexto: não diz o que
    aconteceu, nem que na maioria das vezes basta reiniciar. Esta página diz —
    e, quando o código no disco está mais novo que o processo, coloca isso em
    primeiro lugar, porque é a causa mais comum enquanto o projeto muda.

    Não usa o base.html de propósito: se o erro veio de lá, a tela de erro cair
    junto deixaria a pessoa sem nenhuma pista.
    """
    desatualizado = codigo_mudou_no_disco()
    log.exception("erro em %s", request.url.path)

    motivo = (
        "<p style='margin:0 0 .8rem'><strong>O painel está rodando código antigo.</strong> "
        "Algum arquivo do projeto foi salvo depois que este processo subiu — as telas "
        "já são as novas, o programa não. Feche a janela preta do painel "
        "(<strong>Ctrl+C</strong> nela) e abra <code>painel.bat</code> de novo. "
        "Se ela foi fechada no X, rode <code>parar-painel.bat</code> antes.</p>"
        if desatualizado
        else "<p style='margin:0 0 .8rem'>A automação não conseguiu montar esta tela. "
             "Nada foi gravado. O detalhe técnico está em <code>logs/painel.log</code>.</p>"
    )

    corpo = f"""<!doctype html><html lang="pt-BR"><head><meta charset="utf-8">
<title>Erro no painel</title><style>
 body{{margin:0;background:#14171a;color:#e6e8ea;
   font:15px/1.6 "Segoe UI",system-ui,sans-serif;padding:2.5rem 1.5rem}}
 .caixa{{max-width:44rem;margin:0 auto;background:#1c2024;border:1px solid #f08a82;
   border-radius:10px;padding:1.4rem 1.6rem}}
 h1{{font-size:1.15rem;margin:0 0 .8rem}}
 code{{background:#23282d;padding:.1rem .35rem;border-radius:4px;font-size:.86rem}}
 .tecnico{{margin-top:1rem;color:#9aa3ab;font-size:.82rem}}
 pre{{background:#23282d;padding:.6rem .7rem;border-radius:6px;overflow-x:auto;
   font-size:.78rem;margin:.4rem 0 0}}
 a{{color:#6fa8e8}}
</style></head><body><div class="caixa">
<h1>Não deu para abrir esta tela</h1>
{motivo}
<p style="margin:0"><a href="/">← Voltar para Contas do mês</a></p>
<div class="tecnico">Em <code>{_escapar(request.url.path)}</code>
<pre>{_escapar(type(exc).__name__)}: {_escapar(str(exc))[:600]}</pre></div>
</div></body></html>"""
    return HTMLResponse(corpo, status_code=500)


def _escapar(texto: str) -> str:
    return (
        str(texto)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def competencia_de(request: Request) -> Competencia:
    bruto = request.query_params.get("competencia")
    if bruto:
        try:
            return Competencia.de_texto(bruto)
        except ValueError:
            log.warning("competência inválida na URL: %r", bruto)
    return Competencia.atual()


def contexto_base(request: Request, competencia: Competencia) -> dict:
    from automacao.acesso import permissoes

    amb = ambiente()
    usuario_logado = getattr(request.state, "usuario", None)
    return {
        # O menu é montado a partir disto: link que a pessoa não pode abrir não
        # aparece. Ver automacao/permissoes.py — a trava de verdade é no
        # middleware, isto aqui só evita oferecer porta fechada.
        "telas_visiveis": {t.id for t in permissoes.telas_de(usuario_logado)},
        "request": request,
        "competencia": competencia,
        "competencias": _competencias_vizinhas(competencia),
        "simulando": amb.simulando,
        "ambiente": amb,
        "agora": datetime.now(),
        # Por requisição, não como global do Jinja: assim a tela já mostra o
        # que acabou de ser salvo na Configuração, sem reiniciar o painel.
        "PAGANTES": pagantes(),
        "BENEFICIARIOS": beneficiarios(),
        "codigo_desatualizado": codigo_mudou_no_disco(),
        # Quem está logado: o cabeçalho mostra, e o menu de Usuários só
        # aparece para administrador.
        "usuario_logado": usuario_logado,
        # A logo é arte da empresa e fica fora do git. Sem ela — máquina nova,
        # clone limpo — o cabeçalho volta ao nome em texto em vez de exibir
        # uma imagem quebrada. `scripts/extrair_logo.py` gera os arquivos.
        # A tela precisa dizer se o botão despacha ou só rascunha. Import
        # local: email_outlook puxa o mundo do Outlook, e a maioria das
        # páginas não tem nada com isso.
        "envia_email": _envia_email(amb),
        "tem_logo": (AQUI / "static" / "logo.png").is_file(),
        # Versão para fundo escuro (azul → branco, laranja mantido). Sem ela,
        # o tema escuro cai na logo comum, que some contra o cabeçalho.
        "tem_logo_escura": (AQUI / "static" / "logo-escura.png").is_file(),
        "tem_icone": (AQUI / "static" / "favicon.ico").is_file(),
    }


def _envia_email(amb) -> bool:
    """Se a etapa de e-mail vai despachar ou apenas salvar em Rascunhos."""
    from automacao.entrega import email_outlook

    return email_outlook.deve_enviar(amb)


def _competencias_vizinhas(atual: Competencia) -> list[Competencia]:
    """
    Os meses oferecidos no seletor do topo.

    Corta tudo antes de `competencia_inicial` (settings.yaml): antes desse
    marco o processo era feito à mão, e mês vazio no painel dá a impressão
    falsa de trabalho pendente. O mês que está aberto entra sempre, mesmo
    fora da faixa — quem chegou por link não pode ficar sem o próprio mês na
    barra.
    """
    base = Competencia.atual()
    lista = [base.anterior().anterior(), base.anterior(), base, base.proxima()]

    piso = _competencia_inicial()
    if piso is not None:
        lista = [c for c in lista if c >= piso]

    if atual not in lista:
        lista.append(atual)
    return sorted(set(lista))


def _competencia_inicial() -> Competencia | None:
    bruto = ambiente().competencia_inicial
    if not bruto:
        return None
    try:
        return Competencia.de_texto(bruto)
    except ValueError:
        log.warning("competencia_inicial inválida em settings.yaml: %r", bruto)
        return None


# --------------------------------------------------------------------------- #
# Painel principal
# --------------------------------------------------------------------------- #


@app.get("/vivo")
def vivo():
    """
    Sinal de vida, sem renderizar nada.

    A tela de carregamento consulta este endereço quando a ação demora demais:
    se responder, o painel está de pé e a demora é da operação; se não
    responder, o processo caiu e a pessoa precisa saber disso em vez de
    encarar uma roda girando.
    """
    return Response(status_code=204)


def _resumo_financeiro(linhas: list[dict]) -> dict:
    """
    Quanto do mês já saiu e quanto ainda falta sair.

    A ferramenta acompanha a AUTORIZAÇÃO, não o extrato do banco: "pago" aqui
    quer dizer "autorização despachada para o financeiro". Por isso cada
    número vem com a contagem de contas do lado — quem lê sabe de onde saiu.

    Dois avisos viajam junto e não podem sumir, senão o total parece mais
    exato do que é:

      * `previsto` — parte da soma que veio do valor aproximado do cadastro,
        não de um boleto lido;
      * `sem_valor` — contas que não entraram em soma nenhuma por não terem
        valor nem no documento nem no cadastro.
    """

    def resumir(grupo: list[dict]) -> dict:
        com_valor = [l for l in grupo if l["valor"] is not None]
        return {
            "total": sum(l["valor"] for l in com_valor),
            "quantas": len(grupo),
            "previsto": sum(l["valor"] for l in com_valor if l["valor_estimado"]),
            "sem_valor": len(grupo) - len(com_valor),
        }

    pagas = [l for l in linhas if l["concluido"]]
    abertas = [l for l in linhas if not l["concluido"]]
    return {
        "pago": resumir(pagas),
        "aberto": resumir(abertas),
        "geral": resumir(linhas),
    }


@app.get("/", response_class=HTMLResponse)
def inicio(request: Request):
    competencia = competencia_de(request)

    try:
        todas = contas_ativas_de(competencia)
    except ErroConfiguracao as exc:
        return templates.TemplateResponse(
            request,
            "configuracao_pendente.html",
            {**contexto_base(request, competencia), "erro": str(exc)},
            status_code=200,
        )

    processamentos = estado.carregar_competencia(competencia)
    # Uma consulta só para todas as contas: perguntar mês a mês, conta a
    # conta, seriam 54 idas ao banco para montar uma tela de listagem.
    historico = estado.valores_anteriores(competencia)
    hoje = date.today()
    linhas = []
    for conta in todas:
        proc = processamentos.get(conta.id) or estado.carregar(conta.id, competencia)

        # Vencimento real da fatura, se já foi lido; senão o dia habitual da
        # conta ancorado nesta competência.
        vencimento = proc.vencimento
        if vencimento is None and conta.vencimento_dia:
            vencimento = competencia.dia_vencimento(conta.vencimento_dia)

        # Mesmo tratamento do vencimento: sem valor lido dos documentos,
        # mostra um previsto e marca como tal. A coluna ficava vazia em quase
        # tudo, o que dava a impressão de que o cadastro não servia para nada.
        #
        # A ordem do palpite importa. O último valor pago vem na frente do
        # número do cadastro porque acompanha a realidade sozinho — reajuste,
        # linha a mais, licença a menos —, enquanto o cadastro só muda quando
        # alguém lembra de revisar as 54 contas. O cadastro continua valendo
        # para conta nova, que ainda não tem mês nenhum atrás dela.
        valor, valor_do_mes = proc.valor, None
        if valor is None:
            anterior = historico.get(conta.id)
            if anterior is not None:
                valor, valor_do_mes = anterior
            else:
                valor = conta.valor_estimado

        despachada_em = _quando_foi_despachada(proc)
        # Prazo cumprido apaga o alerta: a linha vermelha existe para cobrar,
        # e não há mais o que cobrar. Envio em atraso continua vermelho, para
        # o registro do que saiu fora do prazo não desaparecer da tela.
        atrasada_no_envio = bool(
            despachada_em and vencimento and despachada_em > vencimento
        )

        linhas.append(
            {
                "conta": conta,
                "proc": proc,
                "despachada_em": despachada_em,
                "atrasada_no_envio": atrasada_no_envio,
                "valor": valor,
                # "é palpite", qualquer que seja a procedência — é o que o
                # fechamento do mês soma à parte para não parecer exato.
                "valor_estimado": proc.valor is None and valor is not None,
                # De que mês veio o palpite. `None` = do cadastro.
                "valor_do_mes": valor_do_mes,
                "vencimento": vencimento,
                "dias_para_vencer": (vencimento - hoje).days if vencimento else None,
                "vencimento_estimado": proc.vencimento is None and vencimento is not None,
                "situacoes": {e: proc.situacao_de(e) for e in ORDEM_PADRAO},
                "concluido": proc.concluido,
                "tem_erro": proc.tem_erro,
                "progresso": sum(
                    1
                    for e in ORDEM_PADRAO
                    if proc.situacao_de(e) in (Situacao.OK, Situacao.ATENCAO, Situacao.PULADO)
                ),
            }
        )

    # Ordem de trabalho: o que vence primeiro aparece primeiro. Conta já
    # concluída desce, e conta sem vencimento cadastrado vai para o fim —
    # não some, mas não disputa espaço com o que tem prazo.
    linhas.sort(
        key=lambda l: (
            l["concluido"],
            l["dias_para_vencer"] is None,
            l["dias_para_vencer"] if l["dias_para_vencer"] is not None else 0,
            l["conta"].rotulo,
        )
    )

    atrasadas = sum(
        1
        for l in linhas
        if not l["concluido"]
        and l["dias_para_vencer"] is not None
        and l["dias_para_vencer"] < 0
    )
    proximas = sum(
        1
        for l in linhas
        if not l["concluido"]
        and l["dias_para_vencer"] is not None
        and 0 <= l["dias_para_vencer"] <= DIAS_DE_ALERTA
    )

    return templates.TemplateResponse(
        request,
        "inicio.html",
        {
            **contexto_base(request, competencia),
            "linhas": linhas,
            "total": len(linhas),
            "concluidas": sum(1 for l in linhas if l["concluido"]),
            "com_erro": sum(1 for l in linhas if l["tem_erro"]),
            "atrasadas": atrasadas,
            "proximas": proximas,
            "resumo": _resumo_financeiro(linhas),
            "hoje": hoje,
        },
    )


CONCLUIDAS = (Situacao.OK, Situacao.ATENCAO, Situacao.PULADO)


def montar_trilha(proc) -> list[dict]:
    """
    Estado de cada etapa para o assistente passo a passo.

    Uma etapa só é liberada quando todas as anteriores estiverem concluídas —
    por automação ou marcadas como feitas à mão. É a regra de não pular etapa.
    """
    trilha: list[dict] = []
    anteriores_ok = True
    for indice, etapa in enumerate(ORDEM_PADRAO, start=1):
        resultado = proc.resultados.get(etapa)
        situacao = resultado.situacao if resultado else Situacao.PENDENTE
        concluida = situacao in CONCLUIDAS
        trilha.append(
            {
                "numero": indice,
                "etapa": etapa,
                "resultado": resultado,
                "situacao": situacao,
                "concluida": concluida,
                "liberada": anteriores_ok,
                "sensivel": etapa in ETAPAS_SENSIVEIS,
            }
        )
        anteriores_ok = anteriores_ok and concluida
    return trilha


def escolher_etapa(trilha: list[dict], pedida: str | None) -> dict:
    """A etapa em foco: a pedida na URL, se liberada; senão a primeira pendente."""
    if pedida:
        for item in trilha:
            if item["etapa"].value == pedida and item["liberada"]:
                return item
    for item in trilha:
        if item["liberada"] and not item["concluida"]:
            return item
    return trilha[-1]


def sugerir_dados_da_autorizacao(conta, competencia: Competencia) -> dict:
    """
    O que o painel já sabe sobre valor, vencimento e nº do documento.

    A etapa 3 abria com os campos vazios mesmo depois da coleta ter lido o
    boleto — o valor só era usado na hora de gerar, e você não tinha como
    conferir antes. Aqui a mesma cadeia de prioridade do `Contexto` roda na
    exibição, e cada campo diz de onde veio.
    """
    vazio = {
        "valor": None,
        "valor_origem": None,
        "valor_estimado": False,
        "vencimento": None,
        "vencimento_origem": None,
        "numero_documento": None,
        "numero_origem": None,
        "descricao": None,
        "descricao_origem": None,
        "descricao_sugerida": None,
        "descricao_conferir": False,
        "pagante": None,
        "pagante_cnpj": None,
        "pagante_aviso": None,
        "beneficiario": None,
        "beneficiario_dados": None,
        "meio_pgto": None,
        "forma_pgto": None,
    }
    try:
        ctx = montar_contexto(conta, competencia)
    except Exception as exc:  # registro incompleto não pode derrubar a tela
        log.warning("não consegui sugerir dados da autorização: %s", exc)
        return vazio

    def origem(salvo, achado, alternativa: str | None = None) -> str | None:
        """
        De onde veio o que está no campo.

        `salvo` é o que já está no banco. Ele pode ter vindo da sua correção
        OU da extração de um mês anterior — por isso a comparação com o que os
        documentos dizem hoje: só chamo de "corrigido por você" o que de fato
        diverge do documento.
        """
        do_doc = achado[0] if achado else None
        if salvo is not None and salvo != do_doc:
            return "corrigido por você"
        if achado:
            return f"{achado[1].tipo.rotulo.lower()} · {achado[1].nome}"
        return alternativa

    valor_doc = ctx.valor_do_documento
    venc_doc = ctx.vencimento_do_documento
    numero_doc = ctx.numero_do_documento
    proc = ctx.processamento

    from automacao.documentos.autorizacao import sugerir_descricao

    sugerida, como_descricao = sugerir_descricao(conta, competencia)
    if proc.descricao is not None:
        descricao, como_descricao = proc.descricao, "reescrito por você"
    else:
        descricao = sugerida

    # Pagante: o cadastro casa tanto pelo nome atual quanto pelo antigo que
    # ainda está na tabela da planilha, então a conta migra sem retrabalho.
    registrado = pagante_por_nome(ctx.pagante)
    aviso_pagante = None
    if ctx.pagante and not registrado:
        aviso_pagante = (
            f"{ctx.pagante!r} não está no cadastro de pagantes — o CNPJ vai sair "
            f"como #N/D. Escolha um da lista ou cadastre em Configuração."
        )

    return {
        "descricao": descricao,
        "descricao_origem": como_descricao,
        "descricao_sugerida": sugerida,
        "descricao_conferir": bool(como_descricao and "Confira" in como_descricao),
        "pagante": registrado.nome if registrado else ctx.pagante,
        "pagante_cnpj": registrado.cnpj if registrado else None,
        "pagante_aviso": aviso_pagante,
        "beneficiario": ctx.beneficiario,
        "beneficiario_dados": beneficiario_por_nome(ctx.beneficiario),
        "meio_pgto": ctx.meio_pgto,
        "forma_pgto": ctx.forma_pgto,
        "valor": ctx.valor,
        "valor_origem": origem(
            proc.valor,
            valor_doc,
            f"{ctx.origem_do_palpite} — confira" if ctx.valor_veio_do_estimado else None,
        ),
        "valor_estimado": ctx.valor_veio_do_estimado,
        "vencimento": ctx.vencimento,
        "vencimento_origem": origem(
            proc.vencimento,
            venc_doc,
            f"dia {conta.vencimento_dia} do registro" if conta.vencimento_dia else None,
        ),
        "numero_documento": ctx.numero_documento,
        "numero_origem": origem(proc.numero_documento, numero_doc),
    }


def buscar_conta(conta_id: str):
    """Conta do registro, ou 404 — link velho não deve virar erro interno."""
    try:
        return conta_por_id(conta_id)
    except KeyError:
        raise HTTPException(
            status_code=404,
            detail=f"conta {conta_id!r} não está no registro (config/fornecedores.yaml)",
        ) from None


def previa_do_email(conta, competencia: Competencia, etapa: Etapa, usuario=None) -> dict:
    """
    Texto do e-mail como ele está agora, para os campos abrirem preenchidos.

    Só monta na etapa do e-mail: é barato (não encosta no Outlook), mas não há
    por que calcular nas outras sete.

    `usuario` precisa chegar aqui, e não só no POST: é esta prévia que enche a
    textarea. Montá-la sem assinatura faria o operador enviar o texto da tela
    — que passa a valer como corpo digitado — sem as três linhas do fim.
    """
    if etapa is not Etapa.EMAIL:
        return {}
    try:
        from automacao.entrega import email_outlook
        from automacao.acesso import usuarios as mod_usuarios

        ctx = montar_contexto(conta, competencia)
        return email_outlook.previa(
            conta,
            competencia,
            valor=ctx.valor,
            vencimento=ctx.vencimento,
            ambiente_=ctx.ambiente,
            assinatura=mod_usuarios.assinatura_de(
                usuario, email_outlook.empresa_da_assinatura(ctx.ambiente)
            ),
            corpo_pessoal=mod_usuarios.corpo_de(usuario),
        )
    except Exception as exc:  # a tela não pode cair por causa da prévia
        log.warning("não consegui montar a prévia do e-mail: %s", exc)
        return {}


@app.get("/conta/{conta_id}", response_class=HTMLResponse)
def detalhe(request: Request, conta_id: str):
    competencia = competencia_de(request)
    conta = buscar_conta(conta_id)
    proc = estado.carregar(conta_id, competencia)
    amb = ambiente()

    trilha = montar_trilha(proc)
    foco = escolher_etapa(trilha, request.query_params.get("etapa"))

    # Os arquivos da fatura moram na pasta DESTA competência, e só nela. A
    # pasta `entrada/<conta_id>/` não participa mais da coleta: ela era comum
    # a todos os meses, e o que sobrava de um mês entrava no seguinte.
    pasta_da_competencia = amb.pasta_trabalho(conta.id, competencia)
    pasta_da_competencia.mkdir(parents=True, exist_ok=True)

    # PDFs que precisam de senha — o painel pede na etapa de coleta.
    try:
        from automacao.acesso import senhas_pdf

        candidatos = [d.caminho for d in proc.documentos if d.caminho.is_file()]
        candidatos += sorted(pasta_da_competencia.glob("*.pdf"))
        pdfs_travados = sorted({p.name for p in senhas_pdf.travados(candidatos)})
        tem_senha_guardada = conta.id in senhas_pdf.contas_com_senha()
    except Exception as exc:
        log.warning("checagem de PDF protegido falhou: %s", exc)
        pdfs_travados, tem_senha_guardada = [], False

    destino = amb.destino_onedrive(conta, competencia)
    return templates.TemplateResponse(
        request,
        "conta.html",
        {
            **contexto_base(request, competencia),
            "conta": conta,
            "proc": proc,
            "sugestao": sugerir_dados_da_autorizacao(conta, competencia),
            "previa_email": previa_do_email(
                conta, competencia, foco["etapa"],
                getattr(request.state, "usuario", None),
            ),
            "recusados": [
                n for n in request.query_params.get("recusados", "").split("|") if n
            ],
            "trilha": trilha,
            "foco": foco,
            "concluidas": sum(1 for i in trilha if i["concluida"]),
            "destino": destino,
            "destino_existe": destino.is_dir(),
            # Só para o aviso "este nome já está aqui" na fila de envio.
            "arquivos_da_competencia": sorted(
                p.name for p in pasta_da_competencia.iterdir() if p.is_file()
            ),
            "pdfs_travados": pdfs_travados,
            "tem_senha_guardada": tem_senha_guardada,
            "pasta_trabalho": amb.pasta_trabalho(conta_id, competencia),
            "nome_base": conta.nome_base_arquivo(competencia),
            "tipos": list(TipoDocumento),
        },
    )


EXTENSOES_ACEITAS = {".pdf", ".xml", ".xlsx", ".xls", ".png", ".jpg", ".jpeg"}


@app.post("/conta/{conta_id}/upload")
async def upload(
    conta_id: str,
    competencia: str = Form(...),
    arquivos: list[UploadFile] = File(...),
):
    """
    Recebe os arquivos da fatura e guarda na pasta DESTA competência.

    Este é o único caminho de entrada de documento. Não há pasta a vigiar nem
    lista a marcar: o que você manda aqui é o que esta fatura tem.

    Antes o arquivo passava por `entrada/<conta_id>/`, que era a mesma pasta
    para todos os meses. O que ninguém recolhia de um mês continuava lá e
    entrava no mês seguinte — foi assim que uma autorização saiu com o valor,
    o vencimento e a nota do mês anterior. Gravando em
    `trabalho/<competencia>/<conta_id>/`, um mês não enxerga o outro.
    """
    comp = Competencia.de_texto(competencia)
    conta = conta_por_id(conta_id)
    pasta = ambiente().pasta_trabalho(conta.id, comp)
    pasta.mkdir(parents=True, exist_ok=True)

    guardados: list[str] = []
    recusados: list[str] = []

    for enviado in arquivos:
        nome = Path(enviado.filename or "").name
        if not nome:
            continue
        if Path(nome).suffix.lower() not in EXTENSOES_ACEITAS:
            recusados.append(nome)
            continue

        conteudo = await enviado.read()

        # Reenviar o mesmo arquivo não cria cópia. Antes criava: o contador
        # abaixo só olhava o NOME, e reenviar o boleto quatro vezes deixou
        # quatro PDFs idênticos disputando lugar no PDF final. Nome igual com
        # conteúdo DIFERENTE continua ganhando sufixo — aí são coisas distintas.
        igual = next(
            (
                existente
                for existente in pasta.iterdir()
                if existente.is_file()
                and existente.stat().st_size == len(conteudo)
                and existente.read_bytes() == conteudo
            ),
            None,
        )
        if igual is not None:
            guardados.append(igual.name)
            log.info("%s já estava em %s — reaproveitando", igual.name, pasta)
            continue

        alvo = pasta / nome
        contador = 2
        while alvo.exists():
            alvo = pasta / f"{Path(nome).stem}_{contador}{Path(nome).suffix}"
            contador += 1

        alvo.write_bytes(conteudo)
        guardados.append(alvo.name)
        log.info("recebi %s (%d bytes) para %s/%s", alvo.name, len(conteudo), conta_id, comp)

    if recusados:
        log.warning("arquivos recusados por extensão: %s", recusados)

    if guardados:
        # A coleta classifica o que está na pasta da competência — inclusive o
        # que acabou de chegar. Não precisa vincular nada: já está no lugar.
        ctx = montar_contexto(conta_id, comp)
        executar_etapa(ctx, Etapa.COLETA)

    destino = _voltar(conta_id, comp, Etapa.COLETA)
    if recusados:
        # Antes isso só ia para o log: o arquivo sumia da tela sem explicação.
        destino.headers["location"] += "&recusados=" + quote("|".join(recusados))
    return destino


@app.get("/conta/{conta_id}/ver")
def ver_arquivo(conta_id: str, arquivo: str, baixar: int = 0):
    """
    Serve um arquivo gerado, para o painel mostrar o PDF sem sair da tela.

    Só entrega o que está dentro de `dados/` — a mesma fronteira que limita a
    escrita. Assim o painel nunca vira um atalho para ler o OneDrive inteiro
    pelo navegador.
    """
    alvo = Path(arquivo)
    amb = ambiente()
    if not amb.e_area_local(alvo) or not alvo.is_file():
        log.warning("recusei servir %s (fora de dados/ ou inexistente)", alvo)
        raise HTTPException(status_code=404, detail="arquivo não disponível")

    # Sem `no-store` o navegador reaproveita o PDF da geração anterior: o
    # endereço é o mesmo a cada reprocessamento, e ele não tem como saber que
    # o conteúdo mudou. A pessoa corrige a autorização, gera de novo e vê o
    # arquivo velho na tela.
    cabecalhos = {"Cache-Control": "no-store, must-revalidate"}
    if not baixar:
        # `inline` deixa o navegador desenhar o PDF no lugar de baixar.
        cabecalhos["Content-Disposition"] = f'inline; filename="{alvo.name}"'

    return FileResponse(
        alvo,
        media_type="application/pdf" if alvo.suffix.lower() == ".pdf" else None,
        filename=alvo.name if baixar else None,
        headers=cabecalhos,
    )


@app.post("/conta/{conta_id}/enviada-fora")
def registrar_envio_externo(
    conta_id: str,
    competencia: str = Form(...),
    data_envio: str = Form(""),
    observacao: str = Form(""),
    voltar_para: str = Form("lista"),
):
    """
    Registra que a fatura já foi enviada, por fora da ferramenta.

    O processo existia antes da automação e boa parte do mês já foi despachada
    na mão. Sem isto, o painel mostraria como pendente algo que o financeiro
    já recebeu — e o risco é mandar duas vezes.
    """
    comp = Competencia.de_texto(competencia)
    quando = _para_data(data_envio) or date.today()

    estado.registrar_envio_externo(conta_id, comp, quando, observacao.strip())

    if voltar_para == "conta":
        return _voltar(conta_id, comp)
    return RedirectResponse(f"/?competencia={comp}", status_code=303)


@app.post("/conta/{conta_id}/desfazer-envio-fora")
def desfazer_envio_externo(conta_id: str, competencia: str = Form(...)):
    """Tira a marcação de enviada e reabre as etapas."""
    comp = Competencia.de_texto(competencia)
    estado.desfazer_envio_externo(conta_id, comp)
    return RedirectResponse(f"/?competencia={comp}", status_code=303)


@app.post("/conta/{conta_id}/senha-pdf")
def salvar_senha_pdf(
    conta_id: str,
    competencia: str = Form(...),
    senha: str = Form(...),
):
    """
    Guarda a senha usada para abrir os PDFs desta conta.

    Fatura de operadora vem protegida — na base atual, 59 arquivos. A senha
    é a mesma todo mês para o mesmo contrato, então fica no cofre cifrado da
    máquina e vale para as próximas competências.
    """
    from automacao.acesso import senhas_pdf

    comp = Competencia.de_texto(competencia)
    senha = senha.strip()
    if not senha:
        return _voltar(conta_id, comp, Etapa.COLETA)

    # Testa contra um PDF travado de verdade antes de guardar, para o usuário
    # descobrir na hora que errou a senha — e não três etapas adiante.
    pasta = ambiente().pasta_trabalho(conta_id, comp)
    travados = senhas_pdf.travados(sorted(pasta.glob("*.pdf")))
    conferida, recado = (True, "guardada sem teste — nenhum PDF travado aqui.")
    if travados:
        conferida, recado = senhas_pdf.testar(travados[0], senha)

    if not conferida:
        log.warning("senha de PDF recusada para %s: %s", conta_id, recado)
        estado.salvar_etapa(
            conta_id,
            comp,
            ResultadoEtapa.atencao(Etapa.COLETA, f"Senha do PDF não serviu: {recado}"),
        )
        return _voltar(conta_id, comp, Etapa.COLETA)

    senhas_pdf.guardar(conta_id, senha)
    log.info("senha de PDF aceita para %s: %s", conta_id, recado)

    # Reexecuta a coleta com a senha disponível: os PDFs que estavam ilegíveis
    # passam a ser classificados de verdade, com valor e vencimento.
    estado.reabrir(conta_id, comp, Etapa.PDF_FINAL)
    ctx = montar_contexto(conta_id, comp)
    executar_etapa(ctx, Etapa.COLETA)
    return _voltar(conta_id, comp, Etapa.COLETA)


@app.post("/conta/{conta_id}/etapa/{etapa}/manual")
def concluir_manualmente(
    conta_id: str,
    etapa: str,
    competencia: str = Form(...),
    observacao: str = Form(""),
):
    """
    Marca a etapa como resolvida por fora da automação.

    Existe porque nem tudo dá para automatizar: fatura com senha, portal que
    exige token, autorização preenchida na mão. Sem essa saída, a conta trava
    e o assistente não deixa seguir.
    """
    comp = Competencia.de_texto(competencia)
    alvo = Etapa(etapa)
    nota = observacao.strip() or "resolvido fora da automação"

    estado.salvar_etapa(
        conta_id,
        comp,
        ResultadoEtapa.pulado(alvo, f"concluída por você — {nota}"),
    )
    # Esta rota grava a etapa direto no banco, sem passar pelo orquestrador —
    # então o gancho de fim de fluxo precisa ser chamado à mão. Sem isto, a
    # conta que fecha pela mão do operador ficaria com o arquivo parado na
    # entrada para sempre.
    recolher_entrada_se_concluiu(conta_id, comp)
    return _voltar(conta_id, comp, alvo)


# --------------------------------------------------------------------------- #
# Ações — preparar (seguro) e confirmar (escreve)
# --------------------------------------------------------------------------- #


@app.post("/conta/{conta_id}/preparar")
def preparar(
    request: Request,
    conta_id: str,
    competencia: str = Form(...),
    valor: str = Form(""),
    vencimento: str = Form(""),
    numero_documento: str = Form(""),
):
    """Roda as etapas que NÃO tocam o mundo externo. Sempre seguro."""
    comp = Competencia.de_texto(competencia)
    ctx = montar_contexto(
        conta_id,
        comp,
        valor=_para_float(valor),
        vencimento=_para_data(vencimento),
        numero_documento=numero_documento.strip() or None,
    )
    executar(
        ctx,
        etapas=[
            Etapa.COLETA,
            Etapa.PASTA,
            Etapa.AUTORIZACAO,
            Etapa.PDF_AUTORIZACAO,
            Etapa.PDF_FINAL,
        ],
    )
    return _voltar(conta_id, comp)


@app.post("/conta/{conta_id}/etapa/{etapa}")
def rodar_etapa(
    request: Request,
    conta_id: str,
    etapa: str,
    competencia: str = Form(...),
    confirmar: str = Form(""),
    valor: str = Form(""),
    vencimento: str = Form(""),
    numero_documento: str = Form(""),
    descricao: str = Form(""),
    pagante: str = Form(""),
    beneficiario: str = Form(""),
    meio_pgto: str = Form(""),
    forma_pgto: str = Form(""),
    editou_email: str = Form(""),
    email_para: str = Form(""),
    email_copia: str = Form(""),
    email_assunto: str = Form(""),
    email_corpo: str = Form(""),
):
    """
    Roda uma etapa isolada.

    Etapas sensíveis (publicação, e-mail, checklist) só executam de verdade
    quando o formulário vem com `confirmar=sim`. Sem isso, a etapa devolve o
    plano do que faria — e é isso que o painel mostra.
    """
    comp = Competencia.de_texto(competencia)
    alvo = Etapa(etapa)
    confirmadas = {alvo} if confirmar == "sim" else set()

    # Campo em branco é diferente de campo ausente: só o formulário do e-mail
    # manda `editou_email`, e aí uma cópia vazia significa "tire todo mundo",
    # não "use o settings.yaml".
    do_painel = bool(editou_email)

    # Quem está logado assina — e é o molde dele que vale, quando tem um.
    # Nada disto passa pelo formulário: assinatura vinda do POST seria
    # assinatura que qualquer um edita pelo navegador.
    from automacao.entrega import email_outlook
    from automacao.acesso import usuarios as mod_usuarios

    quem = getattr(request.state, "usuario", None)
    assinatura = mod_usuarios.assinatura_de(
        quem, email_outlook.empresa_da_assinatura(ambiente())
    )

    ctx = montar_contexto(
        conta_id,
        comp,
        confirmadas=confirmadas,
        valor=_para_float(valor),
        vencimento=_para_data(vencimento),
        numero_documento=numero_documento.strip() or None,
        descricao=descricao.strip() or None,
        pagante=pagante.strip() or None,
        beneficiario=beneficiario.strip() or None,
        meio_pgto=meio_pgto.strip() or None,
        forma_pgto=forma_pgto.strip() or None,
        email_para=email_para.strip() if do_painel else None,
        email_copia=email_copia.strip() if do_painel else None,
        email_assunto=email_assunto.strip() if do_painel else None,
        email_corpo=email_corpo if do_painel else None,
        email_assinatura=assinatura,
        email_corpo_pessoal=mod_usuarios.corpo_de(quem),
    )
    executar_etapa(ctx, alvo)
    return _voltar(conta_id, comp, alvo)


@app.post("/conta/{conta_id}/beneficiario-novo")
async def cadastrar_beneficiario_na_hora(conta_id: str, request: Request):
    """
    Cadastra um beneficiário sem sair da etapa 3.

    Ir até Configuração › Beneficiários no meio do preenchimento custa perder
    o que já estava na tela. Aqui o registro é gravado no mesmo
    `config/beneficiarios.yaml` — é o mesmo cadastro, só que alcançável de
    onde a falta é notada — e já fica escolhido para esta conta.

    Nome vazio não grava nada: seria uma linha fantasma no cadastro. CNPJ com
    dígito que não fecha entra assim mesmo, com aviso: há CNPJ correto que a
    conferência não reconhece, e recusar travaria o trabalho.
    """
    form = await request.form()
    comp = Competencia.de_texto(str(form.get("competencia")))

    def texto(campo: str) -> str:
        return (form.get(campo) or "").strip()

    nome = texto("nome")
    if not nome:
        return _voltar(conta_id, comp, Etapa.AUTORIZACAO)

    novo = Beneficiario(
        nome=nome,
        cnpj=formatar_cnpj(texto("cnpj")) or None,
        contato=texto("contato") or None,
        telefone=texto("telefone") or None,
        banco=texto("banco") or None,
        agencia=texto("agencia") or None,
        conta=texto("conta") or None,
    )

    atuais = list(beneficiarios())
    if beneficiario_por_nome(nome):
        # Já existe com esse nome (ignorando acento e pontuação): atualiza em
        # vez de duplicar. Dois registros com o mesmo nome quebrariam a
        # escolha na etapa 3, que casa justamente pelo nome.
        atuais = [b for b in atuais if not _mesmo_nome(b.nome, nome)]
        log.info("beneficiário %r já existia — substituído pelo novo", nome)
    salvar_beneficiarios(sorted(atuais + [novo], key=lambda b: b.nome))

    # Deixa escolhido para esta conta, senão a etapa 3 recarrega no anterior.
    estado.salvar_dados(conta_id, comp, beneficiario=nome)
    log.info("beneficiário %r cadastrado pela etapa 3 de %s", nome, conta_id)
    return _voltar(conta_id, comp, Etapa.AUTORIZACAO)


def _mesmo_nome(a: str, b: str) -> bool:
    """Compara como o cadastro compara: sem acento, sem pontuação."""
    from automacao.nucleo.config import _chave_de_nome

    return _chave_de_nome(a) == _chave_de_nome(b)


@app.post("/conta/{conta_id}/reabrir")
def reabrir(conta_id: str, competencia: str = Form(...), etapa: str = Form("")):
    """
    "Refazer": esquece o resultado da etapa — e, na coleta, zera os documentos.

    Nas outras sete etapas, refazer é só esquecer o resultado e rodar de novo:
    a etapa lê os mesmos dados e chega ao mesmo lugar. Na coleta não, porque o
    que ela lê são ARQUIVOS, e esquecer o resultado não tira arquivo de lugar
    nenhum — a etapa reencontrava exatamente os mesmos documentos, e refazer
    parecia não fazer nada.

    Então aqui refazer significa o que o nome diz: a coleta volta ao zero. Os
    arquivos coletados saem da pasta da competência e o registro deles some.
    A tela avisa antes, porque o caminho de volta é reenviar.

    Sai só o que está registrado como documento coletado. O que a automação
    gerou — o .xlsx, o PDF da autorização, o PDF final — não é documento de
    entrada e continua onde está; quem o refaz é a etapa que o criou.
    """
    comp = Competencia.de_texto(competencia)
    alvo = Etapa(etapa) if etapa else None

    if alvo is Etapa.COLETA:
        _zerar_coleta(conta_id, comp)

    estado.reabrir(conta_id, comp, alvo)
    return _voltar(conta_id, comp, alvo)


def _zerar_coleta(conta_id: str, competencia: Competencia) -> None:
    """Tira da pasta da competência os documentos coletados, e do estado."""
    from automacao.manutencao import limpeza

    proc = estado.carregar(conta_id, competencia)
    if not proc.documentos:
        return

    presos: list[str] = []
    for doc in proc.documentos:
        try:
            r = limpeza.descartar_da_bancada(
                doc.caminho, conta_id=conta_id, competencia=competencia
            )
        except PermissionError as exc:
            presos.append(f"{doc.nome}: {exc}")
            continue
        if r.falhas:
            presos.append(f"{doc.nome}: {'; '.join(r.falhas)}")

    # Arquivo que não saiu do disco continua no estado. Tirá-lo da tela sem
    # tirar do disco faria a coleta trazê-lo de volta sozinha na execução
    # seguinte — o defeito que esta etapa inteira existe para não repetir.
    restantes = [
        d for d in proc.documentos
        if any(f.startswith(f"{d.nome}:") for f in presos)
    ]
    estado.salvar_documentos(conta_id, competencia, restantes)

    if presos:
        log.error("refazer coleta de %s/%s: %d arquivo(s) não saíram — %s",
                  conta_id, competencia, len(presos), "; ".join(presos))
    else:
        log.info("refazer coleta de %s/%s: %d documento(s) zerado(s)",
                 conta_id, competencia, len(proc.documentos))


@app.post("/conta/{conta_id}/documento")
def reclassificar(
    conta_id: str,
    competencia: str = Form(...),
    caminho: str = Form(...),
    tipo: str = Form(...),
):
    """Correção manual da classificação de um documento."""
    comp = Competencia.de_texto(competencia)
    proc = estado.carregar(conta_id, comp)
    for doc in proc.documentos:
        if str(doc.caminho) == caminho:
            doc.tipo = TipoDocumento(tipo)
            doc.confianca = 1.0
            doc.motivo = "classificado manualmente por você"
            break
    estado.salvar_documentos(conta_id, comp, proc.documentos)
    estado.reabrir(conta_id, comp, Etapa.PDF_FINAL)
    return _voltar(conta_id, comp)


@app.post("/conta/{conta_id}/documento/descartar")
def descartar_documento(
    conta_id: str,
    competencia: str = Form(...),
    caminho: str = Form(...),
):
    """
    Tira um documento desta competência — o desfazer do envio errado.

    Some do estado E da pasta da competência. Só do estado não bastaria: a
    coleta varre essa pasta, então o arquivo voltaria na execução seguinte.

    Como o upload é o único caminho de entrada, não sobra cópia em lugar
    nenhum: descartou, some. Para trazer de volta é só mandar o arquivo de
    novo em "Enviar e coletar" — por isso a tela confirma antes.
    """
    comp = Competencia.de_texto(competencia)
    proc = estado.carregar(conta_id, comp)

    alvo = next((d for d in proc.documentos if str(d.caminho) == caminho), None)
    if alvo is None:
        log.warning("descarte pedido para documento fora do estado: %s", caminho)
        return _voltar(conta_id, comp, Etapa.COLETA)

    from automacao.manutencao import limpeza

    try:
        r = limpeza.descartar_da_bancada(
            alvo.caminho, conta_id=conta_id, competencia=comp
        )
    except PermissionError as exc:
        # A trava recusou. Não mexe no estado: deixar o documento na lista é o
        # que permite ver que o descarte não aconteceu.
        log.error("descarte recusado pela trava: %s", exc)
        return _voltar(conta_id, comp, Etapa.COLETA)

    # Arquivo preso (OneDrive sincronizando, PDF aberto no visualizador) é a
    # única razão para parar aqui. Tirar do estado assim mesmo deixaria o
    # arquivo na pasta da competência e a coleta o traria de volta na execução
    # seguinte — some da tela, volta sozinho, ninguém entende.
    if r.falhas:
        log.error("não consegui apagar %s: %s", alvo.nome, "; ".join(r.falhas))
        return _voltar(conta_id, comp, Etapa.COLETA)

    restantes = [d for d in proc.documentos if str(d.caminho) != caminho]
    estado.salvar_documentos(conta_id, comp, restantes)
    estado.reabrir(conta_id, comp, Etapa.PDF_FINAL)
    log.info(
        "descartado de %s/%s: %s (%d documento(s) restante(s))",
        conta_id, comp, alvo.nome, len(restantes),
    )
    return _voltar(conta_id, comp, Etapa.COLETA)


def _voltar(
    conta_id: str, competencia: Competencia, etapa: Etapa | None = None
) -> RedirectResponse:
    destino = f"/conta/{conta_id}?competencia={competencia}"
    if etapa:
        destino += f"&etapa={etapa.value}"
    return RedirectResponse(destino, status_code=303)


def _para_float(bruto: str) -> float | None:
    bruto = (bruto or "").strip()
    if not bruto:
        return None
    # Aceita "1.234,56", "1234,56" e "1234.56".
    limpo = bruto.replace("R$", "").strip()
    if "," in limpo:
        limpo = limpo.replace(".", "").replace(",", ".")
    try:
        return float(limpo)
    except ValueError:
        log.warning("valor não reconhecido: %r", bruto)
        return None


def _para_data(bruto: str) -> date | None:
    bruto = (bruto or "").strip()
    if not bruto:
        return None
    for formato in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(bruto, formato).date()
        except ValueError:
            continue
    log.warning("data não reconhecida: %r", bruto)
    return None


# --------------------------------------------------------------------------- #
# Entrada manual, auditoria e diagnóstico
# --------------------------------------------------------------------------- #


@app.get("/entrada", response_class=HTMLResponse)
def entrada(request: Request):
    competencia = competencia_de(request)
    amb = ambiente()

    documentos = []
    erro = ""
    try:
        from automacao.coleta import manual

        documentos = manual.varrer_entrada()
    except Exception as exc:
        erro = f"{type(exc).__name__}: {exc}"

    sugestoes = {}
    if documentos:
        try:
            from automacao.coleta import manual

            ativas = [c for c in contas() if c.ativo]
            for doc in documentos:
                sugestoes[str(doc.caminho)] = manual.sugerir_conta(doc, ativas)[:3]
        except Exception as exc:
            log.warning("sugestão de conta indisponível: %s", exc)

    return templates.TemplateResponse(
        request,
        "entrada.html",
        {
            **contexto_base(request, competencia),
            "documentos": documentos,
            "sugestoes": sugestoes,
            "pasta_entrada": amb.caminhos.entrada,
            "erro": erro,
        },
    )


@app.get("/auditoria", response_class=HTMLResponse)
def auditoria(request: Request):
    competencia = competencia_de(request)
    return templates.TemplateResponse(
        request,
        "auditoria.html",
        {
            **contexto_base(request, competencia),
            "registros": estado.ultimas_auditorias(200),
        },
    )


PRAZO_CHECAGEM_COM = 25  # segundos


def _checar_com(rotulo: str, modulo: str, funcao: str) -> tuple[str, bool, str]:
    """Roda uma checagem COM com prazo, para o diagnóstico nunca travar."""
    import importlib
    from concurrent.futures import ThreadPoolExecutor, TimeoutError as PrazoEsgotado

    def executar_checagem():
        alvo = getattr(importlib.import_module(modulo), funcao)
        return alvo()

    with ThreadPoolExecutor(max_workers=1) as executor:
        futuro = executor.submit(executar_checagem)
        try:
            ok, detalhe = futuro.result(timeout=PRAZO_CHECAGEM_COM)
            return rotulo, bool(ok), str(detalhe)
        except PrazoEsgotado:
            return (
                rotulo,
                False,
                f"não respondeu em {PRAZO_CHECAGEM_COM}s — o aplicativo pode estar "
                "ocupado, com uma caixa de diálogo aberta ou rodando com nível de "
                "permissão diferente do Python.",
            )
        except Exception as exc:
            return rotulo, False, f"{type(exc).__name__}: {exc}"


@app.get("/diagnostico", response_class=HTMLResponse)
def diagnostico(request: Request):
    competencia = competencia_de(request)
    amb = ambiente()

    verificacoes = []

    for problema in amb.caminhos.validar():
        verificacoes.append(("Caminhos", False, problema))
    if not amb.caminhos.validar():
        verificacoes.append(
            ("Caminhos", True, f"OneDrive acessível em {amb.caminhos.onedrive}")
        )

    try:
        registro = contas()
        verificacoes.append(
            ("Registro de contas", True, f"{len(registro)} contas carregadas")
        )
    except ErroConfiguracao as exc:
        verificacoes.append(("Registro de contas", False, str(exc)))

    # Excel e Outlook são checados via COM, que pode demorar (ou travar) quando
    # o aplicativo está ocupado. Cada checagem roda com prazo — se estourar,
    # o painel diz isso em vez de ficar carregando para sempre.
    verificacoes.append(_checar_com("Excel (COM)", "automacao.documentos.exportar_pdf", "excel_disponivel"))
    verificacoes.append(_checar_com("Outlook (COM)", "automacao.coleta.outlook", "outlook_disponivel"))

    verificacoes.append(
        (
            "Modo simulação",
            True,
            "LIGADO — nada é gravado fora de dados/"
            if amb.simulando
            else "DESLIGADO — a automação pode gravar no OneDrive",
        )
    )

    return templates.TemplateResponse(
        request,
        "diagnostico.html",
        {**contexto_base(request, competencia), "verificacoes": verificacoes},
    )


# --------------------------------------------------------------------------- #
# Configuração — contas ativas do mês e cadastro de conta nova
# --------------------------------------------------------------------------- #


def contas_ativas_de(competencia: Competencia) -> list:
    """
    Contas que valem nesta competência.

    Se o mês já foi configurado no painel, manda a configuração. Se não, vale
    o `ativo` do registro — assim um mês novo nasce com o padrão em vez de
    nascer vazio.
    """
    escolhidas = estado.contas_do_mes(competencia)
    if escolhidas is None:
        return [c for c in contas() if c.ativo]
    return [c for c in contas() if c.id in escolhidas]


@app.get("/configuracao", response_class=HTMLResponse)
def configuracao(request: Request):
    competencia = competencia_de(request)

    try:
        todas = sorted(contas(), key=lambda c: c.rotulo.lower())
    except ErroConfiguracao as exc:
        return templates.TemplateResponse(
            request,
            "configuracao_pendente.html",
            {**contexto_base(request, competencia), "erro": str(exc)},
        )

    escolhidas = estado.contas_do_mes(competencia)
    ativas_ids = (
        escolhidas if escolhidas is not None else {c.id for c in todas if c.ativo}
    )

    return templates.TemplateResponse(
        request,
        "configuracao.html",
        {
            **contexto_base(request, competencia),
            "todas": todas,
            "ativas_ids": ativas_ids,
            "mes_configurado": escolhidas is not None,
            # Só os OUTROS meses: copiar o mês para ele mesmo não faz nada, e
            # a lista vinha vazia quando agosto era o único configurado.
            "meses_para_copiar": [
                m for m in estado.meses_configurados() if m != str(competencia)
            ],
            "grupos_existentes": sorted(
                {c.email.grupo for c in todas if c.email.grupo}
            ),
            "formatos": list(FORMATOS_PASTA_MES),
            "pastas": sorted(
                {c.pasta for c in todas}
                | {
                    p.name
                    for p in ambiente().caminhos.autorizacoes.iterdir()
                    if p.is_dir()
                }
            ),
        },
    )


# --------------------------------------------------------------------------- #
# Configuração > cadastros de pagantes e beneficiários
# --------------------------------------------------------------------------- #

# `largura` em px é o ponto de partida de cada coluna; dá para arrastar a
# borda do cabeçalho e mudar. Razão social é longa — cortar o nome deixa a
# tabela ilegível, então ela nasce larga e o resto encolhe.
COLUNAS_PAGANTE = [
    {"campo": "nome", "rotulo": "Empresa", "obrigatorio": True,
     "largura": 400, "exemplo": "RAZAO SOCIAL LTDA"},
    {"campo": "cnpj", "rotulo": "CNPJ", "largura": 170,
     "exemplo": "00.000.000/0001-00"},
    {"campo": "banco", "rotulo": "Banco", "largura": 160},
    {"campo": "agencia", "rotulo": "Agência", "largura": 100},
    {"campo": "conta", "rotulo": "Conta", "largura": 120},
]

COLUNAS_BENEFICIARIO = [
    {"campo": "nome", "rotulo": "Nome", "obrigatorio": True, "largura": 340},
    {"campo": "cnpj", "rotulo": "CNPJ", "largura": 170,
     "exemplo": "00.000.000/0001-00"},
    {"campo": "contato", "rotulo": "Contato", "largura": 140},
    {"campo": "telefone", "rotulo": "Telefone", "largura": 140},
    {"campo": "banco", "rotulo": "Banco", "largura": 150},
    {"campo": "agencia", "rotulo": "Agência", "largura": 100},
    {"campo": "conta", "rotulo": "Conta", "largura": 120},
]


def _linhas_do_formulario(form, colunas: list[dict], existentes: list[dict]) -> list[dict]:
    """
    Lê a tabela editável de volta: uma linha por índice, mais a linha "novo".

    Linha marcada para remover não entra no resultado; linha nova só entra se
    o nome foi preenchido — campo em branco no fim da tabela é o estado normal
    de quem não quis cadastrar nada.

    Campos que existem no cadastro mas **não** estão na tela são copiados do
    registro original. Sem isso, tirar uma coluna da tabela apagaria aquele
    dado de todo mundo na primeira gravação — e o formulário não teria como
    avisar, porque nem sabe que o campo existe.
    """
    itens: list[dict] = []
    na_tela = {c["campo"] for c in colunas}

    for i, original in enumerate(existentes):
        if form.get(f"remover__{i}"):
            continue
        registro = dict(original)
        for campo in na_tela:
            registro[campo] = (form.get(f"{campo}__{i}") or "").strip() or None
        if registro.get("nome"):
            itens.append(registro)

    novo = {
        c["campo"]: (form.get(f'novo__{c["campo"]}') or "").strip() or None
        for c in colunas
    }
    if novo.get("nome"):
        itens.append(novo)
    return itens


def _conferir_cnpjs(itens: list[dict]) -> tuple[list[dict], list[str]]:
    """
    Padroniza a pontuação do CNPJ e aponta os que não fecham a conta.

    Reformata só quando os dígitos verificadores conferem — assim
    `11222333000181` vira `11.222.333/0001-81`, mas número errado fica como
    foi digitado, para o aviso ter o que mostrar. Nada é recusado: um CNPJ
    estranho pode ser só um cadastro incompleto, e travar o salvamento
    perderia o resto do que você digitou.
    """
    suspeitos: list[str] = []
    for item in itens:
        bruto = item.get("cnpj")
        if not bruto:
            continue
        if cnpj_confere(bruto):
            item["cnpj"] = formatar_cnpj(bruto)
        else:
            suspeitos.append(f"{item.get('nome')} ({bruto})")
    return itens, suspeitos


def _quando_foi_despachada(proc) -> date | None:
    """
    Data em que a fatura chegou ao financeiro, ou `None` se ainda não chegou.

    Três origens, da melhor para a pior:

    1. `enviado_em` — a data gravada quando a mensagem realmente saiu, seja
       pela ferramenta, seja no registro manual de "enviei por fora". É a
       única que fala do e-mail em si;
    2. o momento em que o checklist foi marcado, para as faturas fechadas
       antes de a ferramenta passar a despachar (aí não há `enviado_em`);
    3. `None`, quando ainda não saiu.

    Sem isso não dá para dizer se saiu no prazo — e comparar o vencimento com
    *hoje* marcaria como atrasada uma fatura despachada a tempo, semanas atrás.
    """
    if proc.enviado_em:
        return proc.enviado_em
    if not proc.concluido:
        return None
    ultima = proc.resultados.get(Etapa.CHECKLIST)
    return ultima.atualizado_em.date() if ultima and ultima.atualizado_em else None


def _pasta_para_registro(amb, caminho: str) -> tuple[str | None, str | None]:
    """
    Traduz um caminho digitado nos dois campos que o registro entende.

    Devolve `(pasta, raiz_onedrive)`:

      * dentro de "Autorizações de pagamento" -> `(subpasta, None)`, o arranjo
        normal, em que a raiz padrão vale;
      * fora dela -> `(último trecho, caminho absoluto)`. O
        `pasta` continua preenchido porque é dele que sai o nome exibido
        quando não há apelido.

    Não adivinha: quem decide é o caminho digitado, e o campo na tela mostra
    o lugar real, então não há ambiguidade sobre o que se está editando.
    """
    alvo = Path(os.path.expandvars(caminho.strip()))
    try:
        dentro = alvo.resolve().relative_to(amb.caminhos.autorizacoes.resolve())
    except (ValueError, OSError):
        return (alvo.name or None), str(alvo)
    return (str(dentro) or None), None


def _pastas_das_contas(amb, competencia: Competencia) -> list[dict]:
    """
    Para onde cada conta ativa arquiva — o que a tela de Parâmetros edita.

    `caminho` é a pasta do fornecedor (a que guarda todos os meses), e
    `exemplo` mostra onde o mês corrente cairia dentro dela. Ver os dois
    juntos é o que evita apontar para o lugar errado sem perceber.
    """
    linhas = []
    for conta in contas_ativas_de(competencia):
        pasta = amb.pasta_do_fornecedor(conta)
        linhas.append(
            {
                "id": conta.id,
                "rotulo": conta.rotulo,
                "pastas": conta.rotulo_das_pastas,
                "caminho": str(pasta),
                "existe": pasta.is_dir(),
                "fora_do_padrao": bool(conta.raiz_onedrive),
                "exemplo": str(amb.destino_onedrive(conta, competencia)),
            }
        )
    return sorted(linhas, key=lambda l: l["rotulo"].lower())


def _lista_editavel(bruto) -> str:
    """`[{nome, endereco}]` -> `'Nome <a@b>; c@d'`, vazio quando não há ninguém."""
    from automacao.entrega import email_outlook

    pessoas = email_outlook._pessoas(bruto)
    return "; ".join(f"{n} <{e}>" if n else e for n, e in pessoas)


@app.get("/configuracao/parametros", response_class=HTMLResponse)
def tela_parametros(request: Request):
    """
    O que a automação está usando: caminhos, travas, e-mail padrão, cadastros.

    Só leitura. Editar caminho pelo navegador é atalho para apontar a
    automação para o lugar errado sem perceber — o arquivo é curto e fica
    versionado junto do projeto.
    """
    competencia = competencia_de(request)
    amb = ambiente()
    c = amb.caminhos

    # Os campos guardam o valor CRU do settings.yaml (o relativo continua
    # relativo); a marca de "encontrado" usa o caminho já resolvido.
    cru = amb.bruto.get("caminhos") or {}
    campos_caminho = [
        ("caminho_onedrive", "OneDrive (raiz)", str(cru.get("onedrive", "")).strip(),
         c.onedrive.is_dir(), "pasta sincronizada da empresa"),
        ("caminho_autorizacoes", "Autorizações de pagamento",
         str(cru.get("autorizacoes", "")).strip(), c.autorizacoes.is_dir(),
         "relativa à raiz — uma subpasta por fornecedor, e dentro dela uma por mês"),
        ("caminho_planilha", "Planilha CONTAS E ACESSOS",
         str(cru.get("planilha_contas", "")).strip(), c.planilha_contas.is_file(),
         "relativa à raiz — checklist do mês e cadastro de acessos"),
        ("caminho_trabalho", "Área de trabalho local",
         str(cru.get("trabalho", "dados/trabalho")).strip(), c.trabalho.is_dir(),
         "onde a autorização é montada antes de publicar"),
        ("caminho_entrada", "Entrada manual",
         str(cru.get("entrada", "dados/entrada")).strip(), c.entrada.is_dir(),
         "uma subpasta por conta; é para cá que vão os arquivos que você sobe"),
        ("caminho_backups", "Backups",
         str(cru.get("backups", "dados/backups")).strip(), c.backups.is_dir(),
         "cópia do que seria sobrescrito, guardada antes de gravar"),
    ]

    try:
        quantas_contas = len(contas())
    except ErroConfiguracao:
        quantas_contas = 0

    from automacao.entrega import email_outlook

    return templates.TemplateResponse(
        request,
        "parametros.html",
        {
            **contexto_base(request, competencia),
            # Barra normal: é referência a arquivo do projeto, não caminho do
            # Windows — `config\settings.yaml` lê pior do que `config/settings.yaml`.
            "arquivo_settings": (DIR_CONFIG / "settings.yaml")
            .relative_to(RAIZ_PROJETO)
            .as_posix(),
            "campos_caminho": campos_caminho,
            "pastas_das_contas": _pastas_das_contas(amb, competencia),
            "problemas": c.validar(),
            "erro": request.query_params.get("erro", ""),
            "guardado": request.query_params.get("guardado", ""),
            # No campo editável, "(ninguém configurado)" atrapalharia: seria
            # salvo de volta como se fosse um endereço.
            "email_para_editavel": _lista_editavel(amb.email.get("para")),
            "email_copia_editavel": _lista_editavel(amb.email.get("copia")),
            "ordem_pdf": " › ".join(amb.pdf.get("ordem") or []) or "—",
            "cor_checklist": amb.checklist.get("cor_enviado", "—"),
            "cadastros": [
                ("Contas (fornecedores)", quantas_contas,
                 "config/fornecedores.yaml", "/configuracao"),
                ("Pagantes", len(pagantes()),
                 "config/pagantes.yaml", "/configuracao/pagantes"),
                ("Beneficiários", len(beneficiarios()),
                 "config/beneficiarios.yaml", "/configuracao/beneficiarios"),
            ],
        },
    )


@app.post("/configuracao/parametros")
async def salvar_parametros(request: Request):
    """
    Grava caminhos e e-mail no `settings.yaml`.

    Duas cautelas, porque um caminho errado aqui derruba tudo:

    * **valida antes de gravar** — se a pasta do OneDrive, a de autorizações
      ou a planilha não existirem com os valores novos, nada é salvo e a tela
      volta dizendo qual falhou. Salvar primeiro e descobrir depois deixaria o
      painel inutilizável, e o conserto seria por editor de texto.
    * **cópia de segurança** — o arquivo anterior vai para `dados/backups/`
      antes de qualquer escrita.
    """
    from automacao.entrega import email_outlook

    form = await request.form()
    competencia = (form.get("competencia") or "").strip()
    voltar = f"/configuracao/parametros?competencia={competencia}"

    def texto(campo: str) -> str:
        return (form.get(campo) or "").strip()

    # --- caminhos: confere o conjunto novo antes de aceitar ---------------- #
    onedrive = Path(os.path.expandvars(texto("caminho_onedrive")))
    autorizacoes = texto("caminho_autorizacoes")
    planilha = texto("caminho_planilha")

    problemas: list[str] = []
    if not onedrive.is_dir():
        problemas.append(f"pasta do OneDrive não encontrada: {onedrive}")
    else:
        if not (onedrive / autorizacoes).is_dir():
            problemas.append(f"subpasta de autorizações não encontrada: {autorizacoes}")
        if not (onedrive / planilha).is_file():
            problemas.append(f"planilha não encontrada: {planilha}")

    # --- e-mail: endereço malformado não entra ----------------------------- #
    pessoas_para = email_outlook._pessoas_de_texto(texto("email_para"))
    pessoas_copia = email_outlook._pessoas_de_texto(texto("email_copia"))
    if not pessoas_para:
        problemas.append("o campo Para ficou vazio — precisa de ao menos um destinatário")
    suspeitos = email_outlook._suspeitos(pessoas_para + pessoas_copia)
    if suspeitos:
        problemas.append("endereço com formato inválido: " + ", ".join(suspeitos))

    janela = texto("outlook_janela_dias")
    if not janela.isdigit() or not 1 <= int(janela) <= 365:
        problemas.append("janela de busca do Outlook deve ser um número de 1 a 365")

    def bloco(pessoas) -> list[dict]:
        return [{"nome": nome, "endereco": endereco} for nome, endereco in pessoas]

    mudancas = {
        "caminhos.onedrive": texto("caminho_onedrive"),
        "caminhos.autorizacoes": autorizacoes,
        "caminhos.planilha_contas": planilha,
        "caminhos.trabalho": texto("caminho_trabalho") or "dados/trabalho",
        "caminhos.entrada": texto("caminho_entrada") or "dados/entrada",
        "caminhos.backups": texto("caminho_backups") or "dados/backups",
        "email.para": bloco(pessoas_para),
        "email.copia": bloco(pessoas_copia),
        "email.assunto": texto("email_assunto"),
        "email.corpo": form.get("email_corpo") or "",
        "email.corpo_agrupado": form.get("email_corpo_agrupado") or "",
        "outlook.janela_dias": int(janela),
    }

    # --- pasta de cada conta: confere ANTES de gravar ---------------------- #
    pastas_novas: dict[str, str] = {}
    for chave, valor in form.items():
        if not chave.startswith("pasta__"):
            continue
        conta_id = chave[len("pasta__"):]
        digitado = (valor or "").strip()
        if not digitado:
            continue
        alvo = Path(os.path.expandvars(digitado))
        if not alvo.is_dir():
            problemas.append(f"pasta de {conta_id} não encontrada: {alvo}")
            continue
        pastas_novas[conta_id] = digitado

    if problemas:
        log.warning("parâmetros recusados: %s", problemas)
        return RedirectResponse(
            voltar + "&erro=" + quote(" | ".join(problemas)), status_code=303
        )

    trocas_de_pasta = _gravar_pastas_das_contas(pastas_novas)

    try:
        aplicadas = salvar_settings(mudancas)
    except Exception as exc:  # noqa: BLE001 — a tela precisa dizer o que houve
        log.exception("falha ao gravar settings.yaml")
        return RedirectResponse(
            voltar + "&erro=" + quote(f"não consegui gravar: {exc}"), status_code=303
        )

    if not aplicadas and not trocas_de_pasta:
        return RedirectResponse(voltar + "&guardado=" + quote("nada mudou."), status_code=303)

    log.info("settings.yaml atualizado: %s", aplicadas)
    partes = []
    if aplicadas:
        partes.append(f"{len(aplicadas)} campo(s) alterados")
    if trocas_de_pasta:
        partes.append(f"{len(trocas_de_pasta)} pasta(s) de conta")
    recado = ", ".join(partes) + ". Cópia do arquivo anterior em dados/backups/."
    return RedirectResponse(voltar + "&guardado=" + quote(recado), status_code=303)


def _gravar_pastas_das_contas(pastas: dict[str, str]) -> list[str]:
    """
    Aponta cada conta para a pasta informada, em config/fornecedores.yaml.

    Só grava o que mudou de fato. Mexer no caminho não move arquivo nenhum:
    daqui para a frente a conta procura e publica no lugar novo, e o que já
    estava no antigo continua lá — mudar de pasta é decisão de quem opera, e
    mover histórico automaticamente seria irreversível.
    """
    if not pastas:
        return []

    amb = ambiente()
    caminho = DIR_CONFIG / "fornecedores.yaml"
    doc = yaml.safe_load(caminho.read_text(encoding="utf-8"))
    mudou: list[str] = []

    for conta in doc["contas"]:
        digitado = pastas.get(conta["id"])
        if not digitado:
            continue
        nova_pasta, nova_raiz = _pasta_para_registro(amb, digitado)
        if nova_pasta is None:
            continue
        if conta.get("pasta") == nova_pasta and (conta.get("raiz_onedrive") or None) == nova_raiz:
            continue
        conta["pasta"] = nova_pasta
        if nova_raiz:
            conta["raiz_onedrive"] = nova_raiz
        else:
            conta.pop("raiz_onedrive", None)
        mudou.append(f"{conta['id']}: pasta → {digitado}")

    if mudou:
        with caminho.open("w", encoding="utf-8") as fh:
            yaml.safe_dump(doc, fh, allow_unicode=True, sort_keys=False, width=120)
        recarregar()
        for item in mudou:
            log.info("configuração: %s", item)
    return mudou


# --------------------------------------------------------------------------- #
# Configuração › E-mail (o que é de cada pessoa)
# --------------------------------------------------------------------------- #


def _usuario_logado(request: Request):
    """
    Quem está na sessão, ou 403.

    O middleware já barra quem não entrou, então na prática isto nunca dispara
    — mas as rotas abaixo gravam no cadastro de alguém, e gravar em `None`
    seria falha calada.
    """
    usuario = getattr(request.state, "usuario", None)
    if usuario is None:
        raise HTTPException(status_code=403, detail="entre no painel primeiro.")
    return usuario


@app.get("/configuracao/email", response_class=HTMLResponse)
def tela_email_usuario(request: Request, guardado: str = "", erro: str = ""):
    """
    A mensagem que o painel monta no Outlook, na parte que é sua.

    Função, setor e o molde do corpo ficam no cadastro da pessoa, não no
    settings.yaml: a máquina é compartilhada e cada um assina com o seu. O
    nome da empresa é o oposto — é o mesmo para todo mundo, e três grafias
    diferentes saindo para o mesmo fornecedor seria o resultado de deixar
    cada um digitar o seu.
    """
    from automacao.entrega import email_outlook
    from automacao.acesso import usuarios as mod

    usuario = _usuario_logado(request)
    amb = ambiente()
    empresa = email_outlook.empresa_da_assinatura(amb)
    corpo_pessoal = mod.corpo_de(usuario)
    corpo_global = email_outlook._molde(amb, "corpo", email_outlook.CORPO_PADRAO)
    assinatura = mod.assinatura_de(usuario, empresa)

    return templates.TemplateResponse(
        request,
        "email_usuario.html",
        {
            **contexto_base(request, competencia_de(request)),
            "usuario": usuario,
            "empresa": empresa,
            "pode_editar_empresa": usuario.administrador,
            "assinatura": assinatura,
            "corpo_pessoal": corpo_pessoal,
            # O texto que vale hoje quando o molde pessoal está vazio: sem ele
            # à vista, "usar o padrão do painel" é uma promessa sem conteúdo.
            "corpo_global": corpo_global,
            # O settings.yaml em uso tem a assinatura batida à mão dentro do
            # corpo. Enquanto ela não bater com a montada aqui, o fornecedor
            # recebe o nome duas vezes — e isso só se vê no rascunho.
            "molde_ja_assinado": email_outlook.duplicaria_assinatura(
                corpo_pessoal or corpo_global, assinatura
            ),
            "guardado": guardado,
            "erro": erro,
        },
    )


@app.post("/configuracao/email")
async def salvar_email_usuario(request: Request):
    """
    Grava função, setor e molde pessoal — e, só para administrador, a empresa.

    Cada campo é gravado **se veio no formulário**. Campo ausente não é campo
    em branco: um POST parcial (aba fechada, campo desabilitado, teste) lido
    com `.get()` apagaria a função e o setor de quem já os tinha. Já aconteceu
    neste projeto, em outra tela.
    """
    from automacao.entrega import email_outlook
    from automacao.acesso import usuarios as mod

    usuario = _usuario_logado(request)
    form = await request.form()
    competencia = (form.get("competencia") or "").strip()
    voltar = f"/configuracao/email?competencia={competencia}"

    mudou: list[str] = []
    for campo, rotulo in (("funcao", "função"), ("setor", "setor")):
        if campo not in form:
            continue
        novo = str(form.get(campo) or "").strip()
        if novo != getattr(usuario, campo):
            setattr(usuario, campo, novo)
            mudou.append(rotulo)

    if "corpo_email" in form:
        # Sem `.strip()` no meio: o molde é texto de várias linhas e a
        # indentação faz parte dele. Só as pontas sobrando é que saem, para
        # "campo em branco" não virar um molde de uma linha vazia.
        novo_corpo = str(form.get("corpo_email") or "").strip()
        if novo_corpo != usuario.corpo_email:
            usuario.corpo_email = novo_corpo
            mudou.append("texto padrão" if novo_corpo else "texto padrão (voltou ao global)")

    if mudou:
        try:
            mod.salvar(usuario)
        except mod.ErroUsuario as erro:
            return RedirectResponse(voltar + "&erro=" + quote(str(erro)), status_code=303)

    # --- nome da empresa: global, e por isso só do administrador ---------- #
    recusado = ""
    if "assinatura_empresa" in form:
        pedido = str(form.get("assinatura_empresa") or "").strip()
        atual = email_outlook.empresa_da_assinatura(ambiente())
        if pedido != atual:
            if not usuario.administrador:
                # Esconder o campo não protege nada: o endereço continua
                # digitável. A recusa é aqui.
                log.warning(
                    "usuário %s tentou mudar email.assinatura_empresa sem ser "
                    "administrador — ignorado",
                    usuario.email,
                )
                recusado = (
                    " O nome da empresa não foi alterado: ele vale para todo "
                    "mundo e só um administrador muda."
                )
            else:
                try:
                    if salvar_settings({"email.assinatura_empresa": pedido}):
                        mudou.append("nome da empresa")
                except Exception as exc:  # noqa: BLE001 — a tela tem que dizer
                    log.exception("falha ao gravar email.assinatura_empresa")
                    return RedirectResponse(
                        voltar + "&erro=" + quote(f"não consegui gravar: {exc}"),
                        status_code=303,
                    )

    if not mudou:
        return RedirectResponse(
            voltar + "&guardado=" + quote("nada mudou." + recusado), status_code=303
        )
    log.info("configuração de e-mail de %s: %s", usuario.email, ", ".join(mudou))
    return RedirectResponse(
        voltar + "&guardado=" + quote("salvo: " + ", ".join(mudou) + "." + recusado),
        status_code=303,
    )


# --------------------------------------------------------------------------- #
# Manutenção da bancada
# --------------------------------------------------------------------------- #


def _mb(bytes_: int) -> str:
    """Tamanho legível. Abaixo de 1 MB, KB — "0,0 MB" não informa nada."""
    if bytes_ >= 1024 * 1024:
        return f"{bytes_ / 1024 / 1024:.1f} MB"
    return f"{bytes_ / 1024:.0f} KB"


@app.get("/configuracao/manutencao", response_class=HTMLResponse)
def tela_manutencao(request: Request, guardado: str = "", erro: str = ""):
    """
    O que a bancada está ocupando e o que pode sair sem perder nada.

    Montar o plano custa alguns SHA-256, então a tela pode demorar um segundo
    em mês cheio. É de propósito: a lista que você vê é exatamente a que o
    botão executa, conferida arquivo por arquivo contra a pasta de destino.
    """
    from automacao.manutencao import limpeza

    competencia = competencia_de(request)
    amb = ambiente()
    regras = limpeza.Regras.de_ambiente(amb)

    plano = limpeza.PlanoLimpeza()
    falha_do_plano = ""
    try:
        plano = limpeza.planejar_mensal(ambiente_=amb)
    except Exception as exc:  # noqa: BLE001 — a tela precisa dizer o que houve
        log.exception("falha ao planejar a limpeza")
        falha_do_plano = f"{type(exc).__name__}: {exc}"

    return templates.TemplateResponse(
        request,
        "manutencao.html",
        {
            **contexto_base(request, competencia),
            "ocupacao": limpeza.ocupacao_da_bancada(ambiente_=amb),
            "plano": plano,
            "regras": regras,
            "ultima_limpeza": limpeza.ultima_limpeza(),
            "falha_do_plano": falha_do_plano,
            "guardado": guardado,
            "erro": erro,
            "mb": _mb,
        },
    )


@app.post("/configuracao/manutencao")
def limpar_bancada(request: Request, competencia: str = Form(...),
                   confirmar: str = Form("")):
    """
    Apaga o que o plano listou. Só com `confirmar=sim`.

    O plano é remontado aqui, e não reaproveitado da tela: entre carregar a
    página e clicar, um processamento pode ter mudado de situação. Executar
    uma lista velha apagaria arquivo que voltou a ser necessário.
    """
    from automacao.manutencao import limpeza

    voltar = f"/configuracao/manutencao?competencia={competencia}"
    if confirmar != "sim":
        return RedirectResponse(
            voltar + "&erro=" + quote("nada foi feito: faltou marcar a confirmação."),
            status_code=303,
        )

    amb = ambiente()
    try:
        plano = limpeza.planejar_mensal(ambiente_=amb)
        resultado = limpeza.executar(plano, confirmado=True, ambiente_=amb)
    except Exception as exc:  # noqa: BLE001
        log.exception("falha ao limpar a bancada")
        return RedirectResponse(
            voltar + "&erro=" + quote(f"{type(exc).__name__}: {exc}"), status_code=303
        )

    if not resultado.apagados and not resultado.falhas:
        return RedirectResponse(
            voltar + "&guardado=" + quote("nada a apagar — a bancada já está limpa."),
            status_code=303,
        )

    recado = (
        f"{len(resultado.apagados)} arquivo(s) apagados, "
        f"{_mb(resultado.bytes_liberados)} liberados"
    )
    if resultado.pastas_removidas:
        recado += f", {resultado.pastas_removidas} pasta(s) vazia(s) recolhida(s)"
    if resultado.falhas:
        recado += f". {len(resultado.falhas)} falharam: " + "; ".join(resultado.falhas[:3])
    log.info("limpeza pela tela: %s", recado)
    return RedirectResponse(voltar + "&guardado=" + quote(recado + "."), status_code=303)


def _exigir_administrador(request: Request):
    """
    Barra quem não é administrador no cadastro de usuários.

    Sem isto, qualquer pessoa logada poderia criar um usuário para si com
    outro e-mail, ou remover os colegas.
    """
    usuario = getattr(request.state, "usuario", None)
    if usuario is None or not usuario.administrador:
        raise HTTPException(
            status_code=403,
            detail="só um administrador mexe no cadastro de usuários.",
        )
    return usuario


@app.get("/configuracao/usuarios", response_class=HTMLResponse)
def tela_usuarios(
    request: Request, guardado: str = "", erro: str = "", avisar: str = ""
):
    from automacao.acesso import permissoes, usuarios as mod

    _exigir_administrador(request)
    return templates.TemplateResponse(
        request,
        "usuarios.html",
        {
            **contexto_base(request, competencia_de(request)),
            "usuarios": sorted(mod.carregar(), key=lambda u: u.email.lower()),
            "grupos": permissoes.carregar(),
            "senha_provisoria": mod.senha_inicial(),
            "guardado": guardado,
            "erro": erro,
            # Quem acabou de ser cadastrado: a tela abre um aviso do navegador
            # com a senha provisória, para o administrador ter o que passar
            # adiante sem ir procurar em outro lugar.
            "avisar": avisar,
        },
    )


@app.post("/configuracao/usuarios/novo")
def usuario_novo(
    request: Request,
    email: str = Form(...),
    nome: str = Form(""),
    grupo: str = Form(""),
):
    from automacao.acesso import usuarios as mod

    _exigir_administrador(request)
    try:
        novo = mod.criar(email, nome=nome, grupo=grupo)
    except mod.ErroUsuario as erro:
        return RedirectResponse(
            "/configuracao/usuarios?erro=" + quote(str(erro)), status_code=303
        )
    return RedirectResponse(
        "/configuracao/usuarios?avisar=" + quote(novo.email) + "&guardado="
        + quote(
            f"{novo.email} cadastrado com a senha provisória "
            f"{mod.senha_inicial()} — ele troca no primeiro acesso."
        ),
        status_code=303,
    )


@app.post("/configuracao/usuarios/grupo")
def usuario_grupo(
    request: Request, email: str = Form(...), grupo: str = Form(...)
):
    """Muda o nível de permissão de alguém direto na lista."""
    from automacao.acesso import permissoes, usuarios as mod

    _exigir_administrador(request)
    try:
        alvo = mod.trocar_grupo(email, grupo)
    except mod.ErroUsuario as erro:
        return RedirectResponse(
            "/configuracao/usuarios?erro=" + quote(str(erro)), status_code=303
        )
    nome_grupo = permissoes.grupo_de(alvo).nome
    return RedirectResponse(
        "/configuracao/usuarios?guardado="
        + quote(f"{alvo.email} agora é do grupo {nome_grupo}."),
        status_code=303,
    )


@app.post("/configuracao/usuarios/redefinir")
def usuario_redefinir(request: Request, email: str = Form(...)):
    from automacao.acesso import usuarios as mod

    _exigir_administrador(request)
    try:
        provisoria = mod.redefinir_senha(email)
    except mod.ErroUsuario as erro:
        return RedirectResponse(
            "/configuracao/usuarios?erro=" + quote(str(erro)), status_code=303
        )
    return RedirectResponse(
        "/configuracao/usuarios?guardado="
        + quote(f"Senha de {email} voltou para {provisoria}; a troca é obrigatória."),
        status_code=303,
    )


@app.post("/configuracao/usuarios/remover")
def usuario_remover(request: Request, email: str = Form(...)):
    from automacao.acesso import usuarios as mod

    eu = _exigir_administrador(request)
    if mod.normalizar_email(email) == mod.normalizar_email(eu.email):
        return RedirectResponse(
            "/configuracao/usuarios?erro="
            + quote("Você não pode remover a si mesmo."),
            status_code=303,
        )
    try:
        mod.remover(email)
    except mod.ErroUsuario as erro:
        return RedirectResponse(
            "/configuracao/usuarios?erro=" + quote(str(erro)), status_code=303
        )
    return RedirectResponse(
        "/configuracao/usuarios?guardado=" + quote(f"{email} removido."),
        status_code=303,
    )


# --------------------------------------------------------------------------- #
# Permissões: o que cada grupo enxerga
# --------------------------------------------------------------------------- #


@app.get("/configuracao/permissoes", response_class=HTMLResponse)
def tela_permissoes(request: Request, guardado: str = "", erro: str = ""):
    from automacao.acesso import permissoes

    _exigir_administrador(request)
    grupos = permissoes.carregar()
    return templates.TemplateResponse(
        request,
        "permissoes.html",
        {
            **contexto_base(request, competencia_de(request)),
            "grupos": grupos,
            "telas": permissoes.TELAS_LIBERAVEIS,
            "telas_admin": [t for t in permissoes.TELAS if t.so_admin],
            # Quantas pessoas em cada grupo: mexer numa linha com gente dentro
            # é diferente de mexer numa linha vazia, e a tela precisa mostrar.
            "quantos": {g.chave: len(permissoes.em_uso(g.chave)) for g in grupos},
            "guardado": guardado,
            "erro": erro,
        },
    )


@app.post("/configuracao/permissoes")
async def salvar_permissoes(request: Request):
    """
    Grava a grade de telas por grupo.

    Lê o marcador `grupo__<chave>` antes das caixas: um grupo sem nenhuma tela
    marcada não manda nenhum campo, e sem o marcador não daria para distinguir
    "desmarcou tudo" de "esta linha nem estava na tela" — o segundo caso
    apagaria as telas de quem não foi editado.
    """
    from automacao.acesso import permissoes

    _exigir_administrador(request)
    formulario = await request.form()
    grupos = permissoes.carregar()

    mexidos: list[str] = []
    for grupo in grupos:
        if grupo.e_admin or f"grupo__{grupo.chave}" not in formulario:
            continue
        nome = str(formulario.get(f"nome__{grupo.chave}") or "").strip()
        if nome:
            grupo.nome = nome
        antes = set(grupo.telas)
        grupo.telas = {
            tela.id for tela in permissoes.TELAS_LIBERAVEIS
            if f"tela__{grupo.chave}__{tela.id}" in formulario
        }
        if grupo.telas != antes:
            mexidos.append(grupo.nome)

    permissoes.gravar(grupos)
    for nome in mexidos:
        log.info("permissões do grupo %s alteradas", nome)
    return RedirectResponse(
        "/configuracao/permissoes?guardado="
        + quote("Permissões guardadas. Quem já está logado vê a mudança na "
                "próxima tela que abrir."),
        status_code=303,
    )


@app.post("/configuracao/permissoes/novo")
def grupo_novo(request: Request, nome: str = Form(...)):
    from automacao.acesso import permissoes

    _exigir_administrador(request)
    chave = permissoes.normalizar(nome)
    if not chave:
        return RedirectResponse(
            "/configuracao/permissoes?erro="
            + quote("Dê um nome ao grupo."), status_code=303
        )
    grupos = permissoes.carregar()
    if any(g.chave == chave for g in grupos):
        return RedirectResponse(
            "/configuracao/permissoes?erro="
            + quote(f"Já existe um grupo chamado {nome.strip()}."),
            status_code=303,
        )
    # Nasce sem nenhuma tela: liberar por engano é pior que a pessoa voltar
    # aqui e marcar o que precisa.
    grupos.append(permissoes.Grupo(chave, nome.strip(), set()))
    permissoes.gravar(grupos)
    log.info("grupo de permissão %s criado", chave)
    return RedirectResponse(
        "/configuracao/permissoes?guardado="
        + quote(f"Grupo {nome.strip()} criado, ainda sem nenhuma tela. "
                "Marque o que ele enxerga e guarde."),
        status_code=303,
    )


@app.post("/configuracao/permissoes/remover")
def grupo_remover(request: Request, chave: str = Form(...)):
    from automacao.acesso import permissoes

    _exigir_administrador(request)
    alvo = permissoes.normalizar(chave)
    if alvo == permissoes.GRUPO_ADMIN:
        return RedirectResponse(
            "/configuracao/permissoes?erro="
            + quote("O grupo Administrador não pode ser removido."),
            status_code=303,
        )
    dentro = permissoes.em_uso(alvo)
    if dentro:
        return RedirectResponse(
            "/configuracao/permissoes?erro="
            + quote(f"{len(dentro)} usuário(s) ainda estão neste grupo "
                    f"({', '.join(dentro[:3])}{'…' if len(dentro) > 3 else ''}). "
                    "Mude o grupo dessas pessoas antes de remover."),
            status_code=303,
        )
    restantes = [g for g in permissoes.carregar() if g.chave != alvo]
    permissoes.gravar(restantes)
    log.warning("grupo de permissão %s removido", alvo)
    return RedirectResponse(
        "/configuracao/permissoes?guardado=" + quote("Grupo removido."),
        status_code=303,
    )


@app.get("/configuracao/pagantes", response_class=HTMLResponse)
def tela_pagantes(request: Request, guardado: str = ""):
    return templates.TemplateResponse(
        request,
        "cadastro.html",
        {
            **contexto_base(request, competencia_de(request)),
            "titulo": "Pagantes",
            "explicacao": (
                "Empresas que pagam a fatura. O nome escolhido aqui vai para o "
                "campo PAGANTE da autorização, e o CNPJ preenche a célula abaixo "
                "dele."
            ),
            "acao": "/configuracao/pagantes",
            "arquivo": "config/pagantes.yaml",
            "colunas": COLUNAS_PAGANTE,
            "chave_larguras": "pagantes",
            "itens": [asdict(p) for p in pagantes()],
            "guardado": guardado,
            "nota": (
                "A tabela de consulta da planilha (aba INFORMAÇÕES DE PROGRAMAÇÃO) "
                "guarda as razões sociais antigas. Quando o nome atual não está "
                "lá, o VLOOKUP do CNPJ devolve #N/D — foi o que saiu nas "
                "autorizações dos últimos meses. Com o CNPJ preenchido aqui, a "
                "automação escreve o número direto na célula e avisa que fez isso. "
                "O reconhecimento ignora acento, pontuação e maiúsculas, então "
                "'FLIX INTELIGÊNCIA' e 'FLIX INTELIGENCIA' são a mesma empresa."
            ),
        },
    )


@app.post("/configuracao/pagantes")
async def salvar_tela_pagantes(request: Request):
    form = await request.form()
    itens, suspeitos = _conferir_cnpjs(
        _linhas_do_formulario(form, COLUNAS_PAGANTE, [asdict(p) for p in pagantes()])
    )
    salvar_pagantes([Pagante.de_dict(d) for d in itens])
    log.info("cadastro de pagantes salvo: %d registro(s)", len(itens))

    recado = f"{len(itens)} pagante(s) salvos."
    if suspeitos:
        log.warning("CNPJ que não confere: %s", suspeitos)
        recado += " CNPJ que não fecha os dígitos verificadores: " + "; ".join(suspeitos)
    return RedirectResponse(
        f"/configuracao/pagantes?competencia={form.get('competencia', '')}"
        f"&guardado={quote(recado)}",
        status_code=303,
    )


@app.get("/configuracao/beneficiarios", response_class=HTMLResponse)
def tela_beneficiarios(request: Request, guardado: str = ""):
    return templates.TemplateResponse(
        request,
        "cadastro.html",
        {
            **contexto_base(request, competencia_de(request)),
            "titulo": "Beneficiários",
            "explicacao": (
                "Quem recebe o pagamento. O nome vai para FORNECEDOR na "
                "autorização; contato, telefone e dados bancários ficam "
                "disponíveis na etapa de preenchimento."
            ),
            "acao": "/configuracao/beneficiarios",
            "arquivo": "config/beneficiarios.yaml",
            "colunas": COLUNAS_BENEFICIARIO,
            "chave_larguras": "beneficiarios",
            "itens": [asdict(b) for b in beneficiarios()],
            "guardado": guardado,
            "nota": (
                "Campo deixado em branco não é escrito na planilha: a célula "
                "fica como estava no modelo do mês anterior, em vez de ser "
                "apagada. O meio de pagamento não está aqui de propósito — "
                "ele muda de uma fatura para outra do mesmo fornecedor, então "
                "é escolhido na etapa 3, ao preencher a autorização."
            ),
        },
    )


@app.post("/configuracao/beneficiarios")
async def salvar_tela_beneficiarios(request: Request):
    form = await request.form()
    itens, suspeitos = _conferir_cnpjs(
        _linhas_do_formulario(form, COLUNAS_BENEFICIARIO, [asdict(b) for b in beneficiarios()])
    )
    salvar_beneficiarios([Beneficiario.de_dict(d) for d in itens])
    log.info("cadastro de beneficiários salvo: %d registro(s)", len(itens))

    recado = f"{len(itens)} beneficiário(s) salvos."
    if suspeitos:
        log.warning("CNPJ que não confere: %s", suspeitos)
        recado += " CNPJ que não fecha os dígitos verificadores: " + "; ".join(suspeitos)
    return RedirectResponse(
        f"/configuracao/beneficiarios?competencia={form.get('competencia', '')}"
        f"&guardado={quote(recado)}",
        status_code=303,
    )


def _dia_do_mes(bruto: str) -> int | None:
    bruto = (bruto or "").strip()
    if not bruto.isdigit():
        return None
    dia = int(bruto)
    return dia if 1 <= dia <= 31 else None


@app.post("/configuracao/ativas")
async def salvar_ativas(request: Request):
    """
    Grava, de uma vez: quem está ativo no mês e as edições da tabela.

    O conjunto ativo é por competência (fica no banco); vencimento, grupo e
    valor estimado são da conta em si (ficam no fornecedores.yaml) e valem
    para todos os meses.
    """
    formulario = await request.form()
    comp = Competencia.de_texto(str(formulario["competencia"]))
    marcadas = {str(v) for v in formulario.getlist("ativa")}

    estado.definir_contas_do_mes(comp, marcadas)

    caminho = DIR_CONFIG / "fornecedores.yaml"
    doc = yaml.safe_load(caminho.read_text(encoding="utf-8"))
    alteracoes: list[str] = []

    for conta in doc["contas"]:
        id_conta = conta["id"]

        # Campo AUSENTE no formulário é diferente de campo em branco. Sem esta
        # distinção, qualquer envio parcial — um POST de teste, uma coluna
        # tirada da tela, um campo desabilitado — apagaria silenciosamente o
        # grupo e o valor de todas as contas.
        if f"venc__{id_conta}" in formulario:
            dia = _dia_do_mes(str(formulario[f"venc__{id_conta}"]))
            if dia != conta.get("vencimento_dia"):
                conta["vencimento_dia"] = dia
                alteracoes.append(f"{id_conta}: vencimento → {dia or '—'}")

        # Campo AUSENTE continua sendo diferente de campo em branco: em
        # branco aqui significa "volte a usar o nome da pasta".
        if f"nome__{id_conta}" in formulario:
            apelido = str(formulario[f"nome__{id_conta}"]).strip() or None
            if apelido != conta.get("apelido"):
                conta["apelido"] = apelido
                alteracoes.append(f"{id_conta}: nome → {apelido or '(o da pasta)'}")

        if f"grupo__{id_conta}" in formulario:
            grupo = str(formulario[f"grupo__{id_conta}"]).strip() or None
            bloco_email = conta.setdefault("email", {})
            if grupo != bloco_email.get("grupo"):
                bloco_email["grupo"] = grupo
                alteracoes.append(f"{id_conta}: grupo → {grupo or '—'}")

        if f"valor__{id_conta}" in formulario:
            valor = _para_float(str(formulario[f"valor__{id_conta}"]))
            if valor != conta.get("valor_estimado"):
                conta["valor_estimado"] = valor
                alteracoes.append(f"{id_conta}: valor estimado → {valor or '—'}")

    if alteracoes:
        with caminho.open("w", encoding="utf-8") as fh:
            yaml.safe_dump(doc, fh, allow_unicode=True, sort_keys=False, width=120)
        recarregar()
        for item in alteracoes:
            log.info("configuração: %s", item)

    log.info(
        "competência %s: %d conta(s) ativas, %d edição(ões)",
        comp,
        len(marcadas),
        len(alteracoes),
    )
    return RedirectResponse(f"/configuracao?competencia={comp}", status_code=303)


@app.post("/configuracao/copiar")
def copiar_mes(competencia: str = Form(...), origem: str = Form(...)):
    """Repete neste mês o conjunto de contas de outro mês."""
    destino = Competencia.de_texto(competencia)
    quantas = estado.copiar_contas_do_mes(Competencia.de_texto(origem), destino)
    log.info("copiei %d conta(s) de %s para %s", quantas, origem, destino)
    return RedirectResponse(f"/configuracao?competencia={destino}", status_code=303)


@app.post("/configuracao/nova")
def nova_conta(
    competencia: str = Form(...),
    pasta: str = Form(...),
    subunidade: str = Form(""),
    chave_planilha: str = Form(""),
    formato_mes: str = Form("MM-AAAA"),
    vencimento_dia: str = Form(""),
    ativar_neste_mes: str = Form(""),
):
    """
    Cadastra uma conta nova no registro (config/fornecedores.yaml).

    Serve para o fornecedor que entrou agora ou para o que existe na planilha
    mas nunca teve pasta de autorização — hoje é o caso de três fornecedores
    da base.
    """
    comp = Competencia.de_texto(competencia)
    pasta = pasta.strip()
    subunidade = subunidade.strip()
    if not pasta:
        return RedirectResponse(f"/configuracao?competencia={comp}", status_code=303)

    novo_id = gerar_id_conta(pasta, subunidade)
    caminho = DIR_CONFIG / "fornecedores.yaml"
    doc = yaml.safe_load(caminho.read_text(encoding="utf-8"))

    if any(c["id"] == novo_id for c in doc["contas"]):
        log.warning("conta %r já existe — nada a fazer", novo_id)
        return RedirectResponse(f"/configuracao?competencia={comp}", status_code=303)

    doc["contas"].append(
        {
            "id": novo_id,
            "ativo": True,
            "pasta": pasta,
            "subunidade": subunidade or None,
            "formato_mes": formato_mes,
            "padrao_nome_arquivo": None,
            "aba_autorizacao": "AUTORIZAÇÃO",
            "modelo_base": None,
            "vencimento_dia": int(vencimento_dia) if vencimento_dia.strip().isdigit() else None,
            "autorizacao": {},
            "referencia": {},
            "revisar": [
                "conta cadastrada pelo painel — falta apontar 'modelo_base' "
                "(a planilha de autorização usada como base) e conferir os "
                "campos em 'autorizacao'."
            ],
            "email": {"eh_operadora": False, "cidade": None, "nome_empresa": pasta, "grupo": None},
            "coleta": {"remetentes": [], "assunto_contem": [], "identificadores": []},
            "planilha_contas": {"chaves": [chave_planilha.strip()] if chave_planilha.strip() else []},
            "observacao": "cadastrada pelo painel",
        }
    )

    with caminho.open("w", encoding="utf-8") as fh:
        yaml.safe_dump(doc, fh, allow_unicode=True, sort_keys=False, width=120)
    recarregar()
    log.info("conta %r cadastrada", novo_id)

    if ativar_neste_mes == "sim":
        atuais = estado.contas_do_mes(comp) or {c.id for c in contas() if c.ativo}
        estado.definir_contas_do_mes(comp, set(atuais) | {novo_id})

    return RedirectResponse(f"/configuracao?competencia={comp}", status_code=303)


def gerar_id_conta(pasta: str, subunidade: str = "") -> str:
    """Mesmo formato de id usado pelos scripts de geração do registro."""
    import re
    import unicodedata

    bruto = "-".join(p for p in (pasta, subunidade) if p)
    bruto = "".join(
        c for c in unicodedata.normalize("NFKD", bruto) if not unicodedata.combining(c)
    ).lower()
    return re.sub(r"-{2,}", "-", re.sub(r"[^a-z0-9]+", "-", bruto).strip("-"))


@app.post("/recarregar")
def recarregar_config(request: Request, competencia: str = Form("")):
    recarregar()
    # Volta para a tela de onde o botão foi clicado, não sempre para a home:
    # recarregar a partir dos Parâmetros é justamente para conferir ali.
    origem = request.headers.get("referer") or ""
    if "/configuracao/parametros" in origem:
        return RedirectResponse(
            f"/configuracao/parametros?competencia={competencia}", status_code=303
        )
    destino = f"/?competencia={competencia}" if competencia else "/"
    return RedirectResponse(destino, status_code=303)


def preparar_pastas_de_entrada() -> int:
    """
    Cria `dados/entrada/<id-da-conta>/` para cada conta ativa.

    Sem essas pastas, a entrada manual não tem onde receber o arquivo e o
    "Preparar" morre na primeira etapa dizendo que não encontrou documento —
    sem deixar claro onde o arquivo deveria estar. Com as pastas prontas,
    você abre o Explorador, vê o nome de cada conta e arrasta o PDF.
    """
    try:
        amb = ambiente()
        ativas = [c for c in contas() if c.ativo]
    except ErroConfiguracao as exc:
        log.warning("registro de contas indisponível: %s", exc)
        return 0

    criadas = 0
    for conta in ativas:
        pasta = amb.caminhos.entrada / conta.id
        if not pasta.exists():
            pasta.mkdir(parents=True, exist_ok=True)
            criadas += 1

    leiame = amb.caminhos.entrada / "LEIA-ME.txt"
    if not leiame.exists():
        leiame.write_text(
            "PASTA DE ENTRADA — onde você larga as faturas baixadas na mão\n"
            "=============================================================\n\n"
            "Cada subpasta tem o nome (id) de uma conta. Baixou a fatura no\n"
            "portal? Arraste o PDF para a subpasta correspondente e clique em\n"
            "PREPARAR no painel.\n\n"
            "Pode jogar tudo de uma vez: boleto, nota fiscal e demonstrativo.\n"
            "A automação identifica o que é cada um e monta na ordem certa.\n\n"
            "Arquivo solto aqui na raiz também funciona — o painel mostra um\n"
            "palpite de qual conta é, na aba Entrada manual.\n\n"
            "Nada aqui é apagado. Depois de publicado, o arquivo é movido para\n"
            "_processados/.\n",
            encoding="utf-8",
        )

    if criadas:
        log.info("criei %d pasta(s) de entrada em %s", criadas, amb.caminhos.entrada)
    return criadas


#: Nome amigável do painel. Só resolve nesta máquina, e só depois de o
#: endereço entrar no arquivo hosts — `python scripts/configurar_endereco.py`,
#: como administrador. Sem isso o painel continua em 127.0.0.1, que sempre
#: funciona.
NOME_LOCAL = "organizacao.financeira.local"

#: A 80 é a porta padrão de http: com ela o endereço fica sem `:porta` no
#: fim. Ocupada (IIS, outro servidor), o painel cai na 8000 sozinho.
PORTA_PADRAO = 80
PORTA_RESERVA = 8000
def _limpar_bancada_do_mes() -> None:
    """
    Varre a bancada uma vez por mês, na primeira subida do painel.

    Roda antes de o servidor atender qualquer requisição: são alguns SHA-256
    numa pasta sincronizada, e fazer isso durante um clique deixaria a tela
    parada sem explicação. Falhar aqui não impede o painel de subir — o pior
    que acontece é a bancada continuar cheia, e a tela de Manutenção existe
    justamente para resolver isso à mão.
    """
    try:
        from automacao.manutencao import limpeza

        resultado = limpeza.rodar_limpeza_mensal_se_for_a_hora()
    except Exception as exc:  # noqa: BLE001 — limpeza nunca impede a subida
        log.warning("limpeza mensal não rodou: %s", exc)
        return

    if resultado and resultado.apagados:
        print(
            f"   Limpeza do mes: {len(resultado.apagados)} arquivo(s) da bancada, "
            f"{resultado.bytes_liberados / 1024 / 1024:.1f} MB liberados.",
            flush=True,
        )


TENTATIVAS_DE_PORTA = 12


def _porta_de_subida(pedida: int) -> int:
    """
    A porta onde o painel realmente vai subir.

    Tenta a pedida; se for a 80 e ela estiver tomada — comum no Windows, em
    que o `http.sys` a reserva para o IIS —, vai direto para a faixa 8000 em
    vez de tentar 81, 82, 83, que ninguém espera num endereço de painel.
    """
    if porta_livre(pedida) == pedida:
        return pedida
    if pedida == 80:
        return porta_livre(PORTA_RESERVA)
    return porta_livre(pedida)


def porta_livre(inicial: int = PORTA_PADRAO) -> int:
    """
    Primeira porta livre a partir de `inicial`.

    Um painel que ficou preso de uma execução anterior continua segurando a
    8000, e às vezes nem dá para encerrá-lo (processo com outro nível de
    permissão devolve "Acesso negado"). Em vez de morrer com um erro de bind
    ilegível, sobe na próxima porta e avisa qual é.
    """
    import socket

    for deslocamento in range(TENTATIVAS_DE_PORTA):
        porta = inicial + deslocamento
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as teste:
            teste.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                teste.bind(("127.0.0.1", porta))
                return porta
            except OSError:
                continue
    raise SystemExit(
        f"nenhuma porta livre entre {inicial} e {inicial + TENTATIVAS_DE_PORTA - 1}. "
        "Encerre os painéis abertos e tente de novo."
    )


#: Para onde vai o que as bibliotecas escreveriam no console, quando não há
#: console. Cresceu além disto, recomeça — são só linhas de subida.
LIMITE_CONSOLE = 1024 * 1024


def _console_para_arquivo() -> None:
    """
    Dá um `stdout`/`stderr` de verdade ao painel que subiu sem console.

    Sob `pythonw.exe` — que é como o painel sobe junto com o Windows — os dois
    são `None`, e qualquer biblioteca que escreva neles quebra. O uvicorn é
    uma delas, e o sintoma é cruel: o log registra "painel em http://..." e o
    processo morre logo depois **sem deixar rastro nenhum**, porque o
    traceback iria justamente para o stderr que não existe.

    Apontar os dois para um arquivo resolve as duas coisas de uma vez: as
    bibliotecas voltam a ter onde escrever, e o que der errado fica gravado
    em `logs/painel-console.log` em vez de se perder.
    """
    import sys

    if sys.stdout is not None and sys.stderr is not None:
        return

    destino = RAIZ_PROJETO / "logs" / "painel-console.log"
    destino.parent.mkdir(parents=True, exist_ok=True)
    modo = "w" if (destino.is_file() and destino.stat().st_size > LIMITE_CONSOLE) else "a"
    # buffering=1 (linha a linha): sem isso, o painel que trava deixa o
    # arquivo vazio — e vazio é exatamente o que não ajuda a investigar.
    fluxo = open(destino, modo, encoding="utf-8", buffering=1, errors="replace")
    if sys.stdout is None:
        sys.stdout = fluxo
    if sys.stderr is None:
        sys.stderr = fluxo


def main() -> None:
    import os
    import uvicorn

    _console_para_arquivo()
    configurar_log(arquivo=str(RAIZ_PROJETO / "logs" / "painel.log"))
    preparar_pastas_de_entrada()
    _limpar_bancada_do_mes()

    # Sem isto, a primeira execução mostraria uma tela de login sem ninguém
    # para entrar. Só roda quando o cadastro está vazio.
    from automacao.acesso import usuarios

    # Quem é o primeiro administrador vem do settings.yaml, não daqui: este
    # arquivo vai para o GitHub, e nome e e-mail de pessoa da empresa não.
    bruto = (ambiente().bruto.get("acesso") or {})
    primeiro = usuarios.garantir_primeiro_usuario(
        str(bruto.get("primeiro_administrador") or "").strip(),
        nome=str(bruto.get("nome_do_administrador") or "").strip(),
    )
    if primeiro:
        print()
        print("   PRIMEIRO ACESSO")
        print(f"   usuario: {primeiro.email}")
        print(f"   senha  : {usuarios.senha_inicial()}   (troca obrigatoria ao entrar)")
        print()

    pedida = int(os.environ.get("PAINEL_PORTA", PORTA_PADRAO))
    porta = _porta_de_subida(pedida)
    if porta != pedida:
        log.warning(
            "a porta %d está ocupada (IIS, outro servidor ou um painel antigo) "
            "— subindo na %d",
            pedida, porta,
        )

    sufixo = "" if porta == 80 else f":{porta}"
    endereco = f"http://127.0.0.1{sufixo}"
    amigavel = f"http://{NOME_LOCAL}{sufixo}"
    log.info("painel em %s (nome local: %s)", endereco, amigavel)
    print(f"\n   >>> Painel no ar: {amigavel}", flush=True)
    print(f"       (também em {endereco})\n", flush=True)

    # Deixa o .bat abrir o navegador na porta certa.
    (RAIZ_PROJETO / "logs" / "porta.txt").write_text(str(porta), encoding="utf-8")

    uvicorn.run(app, host="127.0.0.1", port=porta, log_level="warning")


if __name__ == "__main__":
    main()
