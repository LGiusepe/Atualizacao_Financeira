"""
Gera o manual de uso da ferramenta em PDF.

    python scripts/gerar_manual.py

Sai em `docs/Manual - Automacao Financeira.pdf`. O arquivo é pensado para duas
leituras: a de quem vai operar (índice, passo a passo, o que fazer quando dá
errado) e a de uma ferramenta de perguntas e respostas como o NotebookLM — por
isso os títulos são explícitos, cada seção se sustenta sozinha e há um
glossário no fim ligando o vocabulário do painel ao da planilha.

`docs/` fica fora do git: o manual cita caminhos e fornecedores reais.
"""

from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from reportlab.lib import colors  # noqa: E402
from reportlab.lib.enums import TA_JUSTIFY  # noqa: E402
from reportlab.lib.pagesizes import A4  # noqa: E402
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet  # noqa: E402
from reportlab.lib.units import mm  # noqa: E402
from reportlab.platypus import (  # noqa: E402
    BaseDocTemplate,
    Frame,
    Image,
    ListFlowable,
    ListItem,
    NextPageTemplate,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)

RAIZ = Path(__file__).resolve().parents[1]
LOGO = RAIZ / "painel" / "static" / "logo.png"
SAIDA = RAIZ / "docs" / "Manual - Automacao Financeira.pdf"

MARINHO = colors.HexColor("#17325c")
LARANJA = colors.HexColor("#f07020")
CINZA = colors.HexColor("#5b6169")
CINZA_CLARO = colors.HexColor("#eef0f2")


def _estilos() -> dict:
    base = getSampleStyleSheet()
    return {
        "titulo_capa": ParagraphStyle(
            "titulo_capa", parent=base["Title"], fontSize=28, leading=33,
            textColor=MARINHO, spaceAfter=6, alignment=0),
        "sub_capa": ParagraphStyle(
            "sub_capa", parent=base["Normal"], fontSize=13, leading=19,
            textColor=CINZA, spaceAfter=4),
        "h1": ParagraphStyle(
            "h1", parent=base["Heading1"], fontSize=17, leading=21,
            textColor=MARINHO, spaceBefore=18, spaceAfter=8),
        "h2": ParagraphStyle(
            "h2", parent=base["Heading2"], fontSize=12.5, leading=16,
            textColor=MARINHO, spaceBefore=13, spaceAfter=5),
        "corpo": ParagraphStyle(
            "corpo", parent=base["BodyText"], fontSize=10, leading=15,
            alignment=TA_JUSTIFY, spaceAfter=7),
        "item": ParagraphStyle(
            "item", parent=base["BodyText"], fontSize=10, leading=14.5,
            spaceAfter=3),
        "codigo": ParagraphStyle(
            "codigo", parent=base["Code"], fontSize=9, leading=13,
            backColor=CINZA_CLARO, borderPadding=6, spaceBefore=4,
            spaceAfter=9, leftIndent=0),
        "nota": ParagraphStyle(
            "nota", parent=base["BodyText"], fontSize=9.5, leading=14,
            textColor=CINZA, spaceAfter=8, leftIndent=8),
        "celula": ParagraphStyle(
            "celula", parent=base["BodyText"], fontSize=9, leading=12.5,
            spaceAfter=0),
        "celula_forte": ParagraphStyle(
            "celula_forte", parent=base["BodyText"], fontSize=9, leading=12.5,
            spaceAfter=0, textColor=MARINHO, fontName="Helvetica-Bold"),
    }


E = _estilos()


def p(texto: str, estilo: str = "corpo") -> Paragraph:
    return Paragraph(texto, E[estilo])


def lista(itens: list[str]) -> ListFlowable:
    return ListFlowable(
        [ListItem(p(i, "item"), leftIndent=14) for i in itens],
        bulletType="bullet", bulletChar="\u2022", bulletFontSize=8,
        leftIndent=12, spaceAfter=8,
    )


def tabela(linhas: list[list[str]], larguras: list[float]) -> Table:
    dados = [
        [Paragraph(c, E["celula_forte"] if i == 0 else E["celula"]) for c in linha]
        for i, linha in enumerate(linhas)
    ]
    t = Table(dados, colWidths=larguras, hAlign="LEFT", repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), CINZA_CLARO),
        ("LINEBELOW", (0, 0), (-1, 0), 0.8, MARINHO),
        ("LINEBELOW", (0, 1), (-1, -2), 0.3, colors.HexColor("#dfe3e8")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
    ]))
    return t


def passo(numero: str, titulo: str, corpo: list) -> list:
    """Bloco de uma etapa do assistente, com o número em destaque."""
    saida = [p('<font color="#f07020">' + numero + "</font>&nbsp;&nbsp;" + titulo, "h2")]
    saida.extend(corpo)
    return saida


def _rodape(canvas, doc) -> None:
    canvas.saveState()
    canvas.setFont("Helvetica", 7.5)
    canvas.setFillColor(CINZA)
    canvas.drawString(20 * mm, 12 * mm,
                      "Automação Financeira — Autorizações de Pagamento")
    canvas.drawRightString(A4[0] - 20 * mm, 12 * mm,
                           "página " + str(canvas.getPageNumber()))
    canvas.setStrokeColor(colors.HexColor("#dfe3e8"))
    canvas.line(20 * mm, 15 * mm, A4[0] - 20 * mm, 15 * mm)
    canvas.restoreState()


def _faixa_da_capa(canvas, doc) -> None:
    canvas.saveState()
    canvas.setFillColor(MARINHO)
    canvas.rect(0, A4[1] - 8 * mm, A4[0], 8 * mm, stroke=0, fill=1)
    canvas.setFillColor(LARANJA)
    canvas.rect(0, A4[1] - 8 * mm, A4[0] * 0.34, 8 * mm, stroke=0, fill=1)
    canvas.restoreState()


def _documento() -> BaseDocTemplate:
    doc = BaseDocTemplate(
        str(SAIDA), pagesize=A4,
        leftMargin=20 * mm, rightMargin=20 * mm,
        topMargin=20 * mm, bottomMargin=22 * mm,
        title="Manual — Automação Financeira",
        author="T.I.",
        subject="Como usar a automação de autorizações de pagamento",
    )
    quadro = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height,
                   id="corpo")
    doc.addPageTemplates([
        PageTemplate(id="capa", frames=[quadro], onPage=_faixa_da_capa),
        PageTemplate(id="miolo", frames=[quadro], onPage=_rodape),
    ])
    return doc


def _capa() -> list:
    blocos: list = [Spacer(1, 28 * mm)]
    if LOGO.is_file():
        blocos += [
            Image(str(LOGO), width=52 * mm, height=52 * mm * 130 / 468,
                  hAlign="LEFT"),
            Spacer(1, 16 * mm),
        ]
    blocos += [
        p("Automação Financeira", "titulo_capa"),
        p("Autorizações de Pagamento — manual de uso", "sub_capa"),
        Spacer(1, 8 * mm),
        p("Este manual explica como usar a ferramenta que monta as autorizações "
          "de pagamento das contas de T.I.: da coleta da fatura até o rascunho "
          "do e-mail para o financeiro e a marcação no checklist."),
        Spacer(1, 4 * mm),
        p("Documento gerado em " + date.today().strftime("%d/%m/%Y")
          + ". Uso interno.", "nota"),
    ]
    return blocos


def _visao_geral() -> list:
    return [
        p("1. O que a ferramenta faz", "h1"),
        p("Todo mês, cada conta de T.I. — telefonia, links, licenças de software — "
          "exige a mesma sequência de tarefas manuais: achar a fatura, criar a "
          "pasta do mês, copiar a autorização anterior e reescrever valor, "
          "vencimento e número do documento, exportar em PDF, juntar com boleto e "
          "nota fiscal, enviar ao financeiro e marcar a linha como enviada no "
          "checklist."),
        p("São sete passos vezes o número de contas ativas. O trabalho é "
          "repetitivo, mas não é mecânico: cada operadora nomeia arquivo de um "
          "jeito, o valor está em lugar diferente em cada boleto, e errar "
          "significa pagar a conta errada. A ferramenta faz a parte repetitiva e "
          "para nos pontos em que alguém precisa olhar."),
        p("O caminho de uma conta, do começo ao fim", "h2"),
        tabela([
            ["Etapa", "O que a ferramenta faz", "O que você faz"],
            ["1. Coletar documentos",
             "Procura a fatura no Outlook e aceita arquivos arrastados na tela. "
             "Identifica o que é boleto, nota fiscal e demonstrativo.",
             "Sobe os arquivos, se não vieram por e-mail. Corrige a classificação "
             "se ela errou."],
            ["2. Criar pasta do mês",
             "Mostra onde os arquivos vão parar no OneDrive.",
             "Confere o caminho."],
            ["3. Preencher autorização",
             "Copia o modelo do mês anterior e reescreve valor, vencimento, número "
             "do documento e observações, preservando todas as fórmulas.",
             "Confere os campos, que já vêm preenchidos, e corrige o que precisar."],
            ["4. Exportar em PDF",
             "Abre o Excel e exporta a aba respeitando a área de impressão.",
             "Confere o PDF, que aparece na própria tela."],
            ["5. Montar PDF único",
             "Junta autorização, demonstrativo, boleto e nota fiscal, nessa ordem.",
             "Confere o resultado."],
            ["6. Publicar no OneDrive",
             "Mostra a lista do que será gravado e espera confirmação.",
             "Lê a lista e confirma."],
            ["7. Rascunho do e-mail",
             "Monta a mensagem no Outlook com o PDF anexado.",
             "Ajusta o que precisar e clica em enviar — a ferramenta nunca envia."],
            ["8. Marcar no checklist",
             "Marca a caixa de enviado na planilha de controle, que deixa a "
             "linha verde.",
             "Confirma."],
        ], [38 * mm, 66 * mm, 66 * mm]),
        p("As etapas 6, 7 e 8 são as únicas que tocam o mundo fora do computador. "
          "Nenhuma delas roda sem você confirmar.", "nota"),
    ]


def _comecar() -> list:
    return [
        PageBreak(),
        p("2. Abrir e fechar a ferramenta", "h1"),
        p("Para abrir", "h2"),
        p("Dê duplo clique em <b>painel.bat</b>, na pasta do projeto. Ou, pelo "
          "terminal:"),
        p("cd C:\\Users\\SEU.USUARIO\\developer\\AutomacaoFinanceira<br/>"
          "python -m painel", "codigo"),
        p("O navegador abre sozinho em <b>http://127.0.0.1:8000</b>. Esse endereço "
          "existe só na sua máquina — ninguém na rede acessa."),
        p("Para fechar", "h2"),
        p("Tecle <b>Ctrl+C</b> na janela preta. Fechar a janela no X não encerra o "
          "programa: ele continua rodando invisível e segura a porta, e o painel "
          "seguinte sobe em outro endereço (8001, 8002…) enquanto você olha uma "
          "aba antiga que não responde mais."),
        p("Se isso acontecer, rode <b>parar-painel.bat</b>: ele varre as portas e "
          "encerra o que estiver preso."),
        p("Quando algo muda no programa", "h2"),
        p("Depois de qualquer atualização do código, feche e abra o painel. As "
          "telas são lidas do disco a cada acesso, mas o programa fica na memória "
          "desde que subiu — a mistura das duas coisas causa erros estranhos. O "
          "painel detecta isso sozinho e mostra um aviso vermelho no topo pedindo "
          "o reinício."),
    ]


def _tela_inicial() -> list:
    return [
        PageBreak(),
        p("3. A tela inicial: Contas do mês", "h1"),
        p("É a lista de trabalho do mês, ordenada por vencimento — o que vence "
          "antes aparece primeiro. Conta já concluída desce para o fim."),
        tabela([
            ["Coluna", "O que mostra"],
            ["Conta", "Nome do fornecedor e, quando é operadora, a cidade atendida."],
            ["Progresso", "Oito quadradinhos, um por etapa. Verde é concluída, "
                          "âmbar é revisar, vermelho é erro."],
            ["Situação", "Não iniciada, em andamento, concluída, com erro, ou "
                         "enviada por fora."],
            ["Valor", "O valor lido do boleto. Quando ainda não há boleto, mostra o "
                      "valor aproximado do cadastro, marcado como previsto."],
            ["Vencimento", "A data e quanto falta. Vermelho quando já venceu, âmbar "
                           "quando falta pouco."],
        ], [30 * mm, 140 * mm]),
        p("Os três botões de cada linha", "h2"),
        lista([
            "<b>Preparar</b> — roda de uma vez as etapas que não tocam nada fora "
            "da máquina (coleta, pasta, autorização, os dois PDFs). É seguro e é "
            "por onde se começa.",
            "<b>Marca de confirmado</b> — registra que a fatura já foi enviada por "
            "fora da ferramenta, com a data. Serve para o painel não cobrar algo "
            "que o financeiro já recebeu.",
            "<b>Abrir</b> — entra no assistente daquela conta, etapa por etapa.",
        ]),
        p("O seletor de mês", "h2"),
        p("No canto superior direito. Cada mês tem sua própria lista de contas "
          "ativas e seu próprio andamento. Trocar de mês não perde nada do mês "
          "anterior."),
        p("O fechamento do mês", "h2"),
        p("No pé da lista, três números: quanto já saiu, quanto falta sair e o "
          "total. \"Já pagas\" conta as faturas cuja autorização foi despachada "
          "ao financeiro — a ferramenta acompanha o despacho, não o extrato do "
          "banco. Quando alguma conta não tem valor nem no documento nem no "
          "cadastro, o bloco avisa em laranja que a soma está incompleta; o total "
          "real é maior do que o exibido."),
        p("Tema claro e escuro", "h2"),
        p("A chave fica à direita do seletor de mês: sol para clarear, lua para "
          "escurecer. Sem escolha, o painel segue o tema do Windows. Feita a "
          "escolha, ela vale sempre naquele navegador, mesmo que o Windows mude. "
          "A logomarca acompanha o tema."),
    ]


def _assistente() -> list:
    blocos: list = [
        PageBreak(),
        p("4. O assistente, etapa por etapa", "h1"),
        p("Ao abrir uma conta, as oito etapas aparecem em fila no topo. A seguinte "
          "só libera quando a anterior fecha — por automação ou porque você marcou "
          "<b>Já resolvi por fora</b>. Isso existe para nada ser pulado sem alguém "
          "notar."),
    ]

    blocos += passo("1", "Coletar documentos", [
        p("Arraste os arquivos para a área tracejada, ou clique para escolher. "
          "Pode mandar boleto, nota fiscal e demonstrativo de uma vez."),
        p("Cada arquivo escolhido aparece numa lista com nome, tamanho e o que há "
          "de errado com ele: tipo não aceito, arquivo vazio, nome repetido, ou já "
          "presente na pasta. O botão de enviar só libera quando há pelo menos um "
          "arquivo válido, e os recusados não sobem."),
        p("Depois do envio, a ferramenta classifica cada documento e mostra a "
          "confiança da decisão. Se errou, corrija pelo seletor da coluna à "
          "direita."),
        p("<b>PDF com senha:</b> faturas de operadora costumam vir protegidas. "
          "Quando isso acontece, a tela pede a senha, testa antes de guardar e a "
          "mantém cifrada nesta máquina — vale para os próximos meses, uma vez por "
          "conta.", "nota"),
    ])

    blocos += passo("2", "Criar pasta do mês", [
        p("Só mostra onde os arquivos vão ficar. Nada é criado agora: a pasta nasce "
          "na etapa 6, depois da sua confirmação."),
    ])

    blocos += passo("3", "Preencher autorização", [
        p("Os campos já vêm preenchidos com o que foi lido dos documentos, e cada "
          "um diz de onde veio — por exemplo <i>boleto · Boletos-123.pdf</i>. O "
          "boleto tem preferência sobre os outros documentos, porque é ele que "
          "informa quanto e quando pagar."),
        p("Tudo é editável. O que você digitar vale sobre o que foi extraído, e o "
          "botão <b>Desfazer edição</b> volta ao valor lido."),
        tabela([
            ["Campo", "De onde vem", "Observação"],
            ["Valor", "Boleto, depois nota fiscal, depois demonstrativo.",
             "Sem documento, cai no valor aproximado do cadastro e avisa."],
            ["Vencimento", "Mesma ordem.",
             "Sem documento, usa o dia habitual da conta."],
            ["Nº do documento", "Nota fiscal, depois boleto.", ""],
            ["Pagante", "Cadastro de pagantes.",
             "Define também o CNPJ que sai na autorização."],
            ["Beneficiário", "Cadastro de beneficiários.",
             "Traz contato, telefone e dados bancários."],
            ["Forma e meio de pgto", "Listas do próprio modelo da planilha.",
             "Muda de fatura para fatura; por isso fica aqui."],
            ["Observações", "Texto do mês anterior, com o mês atualizado.",
             "Confira quando a ferramenta avisar que a distância parece estranha."],
        ], [32 * mm, 62 * mm, 76 * mm]),
    ])

    blocos += passo("4 e 5", "Exportar e montar os PDFs", [
        p("A etapa 4 abre o Excel e exporta a aba da autorização. Leva de dez a "
          "vinte segundos — é o Excel abrindo, não travamento."),
        p("A etapa 5 junta tudo num arquivo só, nesta ordem: autorização, "
          "demonstrativo, boleto, nota fiscal. O PDF recebe o mesmo nome da "
          "planilha, para não duplicar arquivo no servidor."),
        p("Nas duas, o arquivo gerado aparece na própria tela. Confira antes de "
          "seguir: é esse arquivo que vai para o OneDrive e para o anexo do "
          "e-mail."),
    ])

    blocos += passo("6", "Publicar no OneDrive", [
        p("Primeiro a ferramenta mostra a lista completa do que faria — cada "
          "arquivo, cada destino, com aviso quando algo seria sobrescrito. Nada é "
          "gravado até você clicar em confirmar."),
        p("Arquivo que seria sobrescrito vira cópia de segurança antes, e toda "
          "gravação fica registrada na aba Auditoria."),
    ])

    blocos += passo("7", "Criar rascunho do e-mail", [
        p("Destinatários, cópia, assunto e texto aparecem editáveis. O que você "
          "mudar vale só para aquela mensagem — a configuração padrão não é "
          "alterada."),
        p("O rascunho é salvo em <i>Rascunhos</i> do Outlook e aberto para "
          "conferência. <b>Quem clica em Enviar é você.</b> Não há mais "
          "confirmação na tela do painel: um clique já monta o rascunho."),
        p("<b>Por que não envia sozinho.</b> Foi tentado em 02/09/2026 e não "
          "funciona nesta máquina: a automação conversa com o Outlook clássico, "
          "que não fica aberto aqui, e o envio para na Caixa de Saída dele sem "
          "nunca ser transmitido — o Outlook novo tem armazenamento próprio e "
          "não esvazia aquela fila. Despacho automático exigiria Microsoft "
          "Graph, que é outro projeto."),
        p("Endereço com formato inválido ou campo <i>Para</i> vazio barram antes de "
          "chamar o Outlook.", "nota"),
    ])

    blocos += passo("8", "Marcar verde no checklist", [
        p("Marca a caixa <b>Enviou para Financeiro?</b> na linha da conta, na aba "
          "do mês da planilha de controle. Como nas anteriores, mostra qual aba, "
          "qual linha e qual caixa antes de gravar."),
        p("A linha fica verde por causa da caixa: a planilha tem formatação "
          "condicional que pinta a linha inteira quando ela está marcada. Até "
          "01/09/2026 a automação pintava a linha por conta própria — o que "
          "deixava a coluna dizendo \"não enviado\" numa linha verde, e não "
          "alcançava as colunas H a K.", "nota"),
        p("Nas abas antigas, que não têm essa coluna, o preenchimento continua "
          "sendo aplicado como antes.", "nota"),
    ])
    return blocos


def _configuracao() -> list:
    return [
        PageBreak(),
        p("5. Configuração", "h1"),
        p("O menu <b>Configuração</b> abre quatro telas. O que você salva nelas "
          "fica gravado em arquivo e continua valendo depois de fechar o painel."),
        p("Faturas ativas", "h2"),
        p("Quais contas valem no mês, com um interruptor redondo por linha. A "
          "escolha é <b>por mês</b>: desmarcar aqui não apaga nada nem mexe nos "
          "meses anteriores. Há um atalho para repetir a configuração de outro mês."),
        p("Na mesma tabela dá para editar o dia de vencimento, o valor aproximado e "
          "o grupo de e-mail. Esses três são da conta em si e valem para todos os "
          "meses."),
        p("<b>Grupo de e-mail:</b> contas com o mesmo grupo saem numa única "
          "mensagem, com um anexo por fatura. É o caso de fornecedor que emite "
          "várias faturas por mês.", "nota"),
        p("Pagantes", "h2"),
        p("As empresas do grupo que pagam a fatura, com CNPJ. O nome escolhido na "
          "etapa 3 vai para o campo PAGANTE da autorização."),
        p("A tabela de consulta dentro da planilha guarda as razões sociais "
          "antigas. Quando o nome atual não está lá, a fórmula do CNPJ falharia e "
          "sairia <i>#N/D</i> no PDF — foi o que aconteceu por vários meses. Com o "
          "CNPJ cadastrado aqui, a ferramenta escreve o número direto na célula e "
          "avisa na tela que fez isso."),
        p("Beneficiários", "h2"),
        p("Quem recebe o pagamento: nome, CNPJ, contato, telefone e dados "
          "bancários. Campo deixado em branco não é escrito na planilha — a célula "
          "fica como estava no modelo do mês anterior, em vez de ser apagada."),
        p("O CNPJ é conferido pelos dígitos verificadores ao salvar. Se não fechar, "
          "a tela avisa mas salva o resto do que você digitou."),
        p("Parâmetros do sistema", "h2"),
        p("Caminhos das pastas e configuração do e-mail padrão, ambos editáveis. "
          "Cada caminho mostra se foi encontrado. Antes de salvar, a ferramenta "
          "confere: se a pasta ou a planilha não existirem, nada é gravado e a tela "
          "diz qual falhou. O arquivo anterior vira cópia de segurança."),
        p("A mesma tela mostra, só para leitura, as travas de segurança, a ordem "
          "das páginas do PDF e a contagem dos cadastros."),
        p("Nas duas tabelas de cadastro", "h2"),
        lista([
            "A última linha, em azul, é para cadastrar um registro novo.",
            "A lixeira marca o registro para sair — ele só some depois que você "
            "salvar, e desmarcar antes disso desfaz.",
            "As colunas são arrastáveis pela borda do cabeçalho, e a largura "
            "escolhida fica guardada neste navegador.",
        ]),
    ]


def _seguranca() -> list:
    return [
        PageBreak(),
        p("6. O que a ferramenta nunca faz", "h1"),
        p("O processo mexe com pagamento. As garantias abaixo são estruturais — "
          "não dependem de alguém lembrar de tomar cuidado."),
        tabela([
            ["Garantia", "Como funciona"],
            ["Nunca apaga nada fora da área local",
             "Qualquer tentativa de exclusão fora da pasta de trabalho é recusada "
             "pelo código, e essa trava não tem chave: desligar o modo simulação "
             "libera gravação, jamais exclusão. Há um script que prova isso "
             "(scripts/provar_nao_apaga.py)."],
            ["Nunca envia e-mail",
             "Não existe chamada de envio no código. O máximo é salvar em "
             "Rascunhos."],
            ["Nunca grava sem mostrar antes",
             "As três etapas que tocam OneDrive, Outlook e planilha mostram a lista "
             "completa do que fariam e esperam confirmação."],
            ["Sempre copia antes de sobrescrever",
             "O arquivo original vai para a pasta de backups antes de ser "
             "substituído."],
            ["Registra tudo",
             "Cada gravação fica na aba Auditoria com origem, destino, cópia de "
             "segurança e uma assinatura do conteúdo."],
        ], [44 * mm, 126 * mm]),
        p("Modo simulação", "h2"),
        p("Enquanto o modo simulação estiver ligado, <b>nada</b> é gravado no "
          "OneDrive nem na planilha compartilhada. As etapas sensíveis mostram o "
          "que fariam e param aí. O rascunho do e-mail, esse sim, é criado de "
          "verdade, para você conferir texto e anexo."),
        p("É o estado recomendado até a primeira validação completa. Para liberar a "
          "gravação, mude <b>seguranca.simulacao</b> para <b>false</b> no arquivo "
          "de configuração. A faixa azul no topo do painel sempre diz em qual dos "
          "dois estados a ferramenta está."),
    ]


def _problemas() -> list:
    return [
        PageBreak(),
        p("7. Quando algo dá errado", "h1"),
        tabela([
            ["Sintoma", "Causa provável", "O que fazer"],
            ["A tela devolve Not Found ou erro interno",
             "O painel está rodando código antigo.",
             "Feche e abra o painel. Se a janela já foi fechada no X, rode "
             "parar-painel.bat antes."],
            ["A roda de carregamento não para",
             "O painel caiu, ou uma caixa do navegador está esperando resposta "
             "atrás dela.",
             "Depois de 25 segundos a própria tela diz o que houve e oferece "
             "recarregar."],
            ["O endereço abriu em 8001 em vez de 8000",
             "Sobrou um painel antigo segurando a porta.",
             "Rode parar-painel.bat e depois painel.bat."],
            ["A etapa 1 não encontra documento",
             "A conta não tem remetente configurado e a pasta de entrada está "
             "vazia.",
             "Suba os arquivos pela tela, arrastando."],
            ["O PDF final saiu com menos páginas",
             "Algum PDF está protegido por senha.",
             "Informe a senha na etapa 1 e refaça a etapa 5."],
            ["O CNPJ saiu como #N/D",
             "O pagante escolhido não está no cadastro.",
             "Cadastre-o em Configuração, aba Pagantes, com o CNPJ."],
            ["A planilha não abre para gravação",
             "Ela está aberta no Excel, aqui ou em outra máquina.",
             "Feche a planilha e repita a etapa."],
        ], [45 * mm, 60 * mm, 65 * mm]),
        p("Diagnóstico do ambiente", "h2"),
        p("A aba <b>Diagnóstico</b> testa, de uma vez, o acesso às pastas, o "
          "registro de contas, o Excel e o Outlook. É o primeiro lugar a olhar "
          "quando alguma coisa parou de funcionar sem motivo aparente."),
    ]


def _glossario() -> list:
    return [
        PageBreak(),
        p("8. Glossário", "h1"),
        p("O painel e a planilha usam palavras diferentes para as mesmas coisas. "
          "Esta tabela liga as duas."),
        tabela([
            ["Termo", "Significa"],
            ["Competência", "O mês de referência da fatura. É a chave de tudo: cada "
                            "conta tem um andamento por competência."],
            ["Conta", "Uma linha do checklist — um serviço contratado de um "
                      "fornecedor. Um mesmo fornecedor pode ter várias contas."],
            ["Sub-unidade", "Quando o fornecedor emite faturas separadas por cidade "
                            "ou serviço, cada uma vira uma subpasta e uma conta."],
            ["Pagante", "A empresa do grupo que paga a fatura. Vai para o campo "
                        "PAGANTE da autorização e determina o CNPJ."],
            ["Beneficiário", "Quem recebe o pagamento. Vai para o campo FORNECEDOR."],
            ["Autorização", "A planilha preenchida — o formulário interno que "
                            "autoriza o pagamento. Também chamada de folha de rosto."],
            ["Demonstrativo", "O detalhamento da cobrança que a operadora envia "
                              "junto com o boleto."],
            ["PDF único", "Autorização, demonstrativo, boleto e nota fiscal "
                          "reunidos num arquivo só, nessa ordem, com o mesmo nome "
                          "da planilha."],
            ["Área local de trabalho", "A pasta na sua máquina onde tudo é montado "
                                       "antes de ir para o OneDrive."],
            ["Modo simulação", "Estado em que nada é gravado fora da área local."],
            ["Enviada por fora", "Marcação de que a fatura foi despachada sem usar "
                                 "a ferramenta, para não cobrar duas vezes."],
        ], [38 * mm, 132 * mm]),
    ]


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    SAIDA.parent.mkdir(parents=True, exist_ok=True)

    historia: list = []
    historia += _capa()
    historia.append(NextPageTemplate("miolo"))
    historia.append(PageBreak())
    historia += _visao_geral()
    historia += _comecar()
    historia += _tela_inicial()
    historia += _assistente()
    historia += _configuracao()
    historia += _seguranca()
    historia += _problemas()
    historia += _glossario()

    _documento().build(historia)
    print("manual : " + str(SAIDA))
    print("tamanho: {:.0f} KB".format(SAIDA.stat().st_size / 1024))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
