"""
Contratos de dados compartilhados por todos os módulos.

Este arquivo é a fonte da verdade das estruturas que circulam no pipeline.
Alterou aqui, alterou para todo mundo — mexa com cuidado.
"""

from __future__ import annotations

import calendar
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime
from enum import Enum
from pathlib import Path

MESES_PT = [
    "janeiro", "fevereiro", "março", "abril", "maio", "junho",
    "julho", "agosto", "setembro", "outubro", "novembro", "dezembro",
]


def sem_acento(texto: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFKD", texto) if not unicodedata.combining(c)
    )


def normalizar(texto: str) -> str:
    """Normaliza para comparação: sem acento, minúsculo, espaços colapsados."""
    return re.sub(r"\s+", " ", sem_acento(texto or "").lower()).strip()


# --------------------------------------------------------------------------- #
# Competência (o mês de referência da fatura)
# --------------------------------------------------------------------------- #

FORMATOS_PASTA_MES = {
    "MM-AAAA": "{mes:02d}-{ano:04d}",
    "AAAA-MM": "{ano:04d}-{mes:02d}",
    "MM_AAAA": "{mes:02d}_{ano:04d}",
    "MM - AAAA": "{mes:02d} - {ano:04d}",
}


@dataclass(frozen=True, order=True)
class Competencia:
    """Um mês de referência. É a chave de tudo no pipeline."""

    ano: int
    mes: int

    def __post_init__(self) -> None:
        if not 1 <= self.mes <= 12:
            raise ValueError(f"mês inválido: {self.mes}")

    @classmethod
    def de_texto(cls, texto: str) -> Competencia:
        """Aceita '2026-08', '08-2026', '08/2026'."""
        nums = re.findall(r"\d+", texto)
        if len(nums) != 2:
            raise ValueError(f"competência não reconhecida: {texto!r}")
        a, b = int(nums[0]), int(nums[1])
        return cls(a, b) if a > 12 else cls(b, a)

    @classmethod
    def de_data(cls, quando: date) -> Competencia:
        return cls(quando.year, quando.month)

    @classmethod
    def atual(cls) -> Competencia:
        return cls.de_data(date.today())

    def somar(self, meses: int) -> Competencia:
        """Anda `meses` para frente (ou para trás, se negativo), virando o ano."""
        absoluto = self.ano * 12 + (self.mes - 1) + meses
        return Competencia(absoluto // 12, absoluto % 12 + 1)

    def proxima(self) -> Competencia:
        return Competencia(self.ano + 1, 1) if self.mes == 12 else Competencia(self.ano, self.mes + 1)

    def anterior(self) -> Competencia:
        return Competencia(self.ano - 1, 12) if self.mes == 1 else Competencia(self.ano, self.mes - 1)

    def nome_pasta(self, formato: str) -> str:
        """Nome da subpasta do mês, no formato que aquele fornecedor usa."""
        molde = FORMATOS_PASTA_MES.get(formato)
        if molde is None:
            raise ValueError(
                f"formato de pasta desconhecido: {formato!r}. "
                f"Válidos: {sorted(FORMATOS_PASTA_MES)}"
            )
        return molde.format(ano=self.ano, mes=self.mes)

    @property
    def nome_mes(self) -> str:
        return MESES_PT[self.mes - 1]

    @property
    def extenso(self) -> str:
        return f"{self.nome_mes}/{self.ano}"

    @property
    def aba_checklist(self) -> str:
        return f"Checklist-{self.nome_mes.capitalize()}"

    def dia_vencimento(self, dia: int) -> date:
        """Ancora um dia do mês, respeitando meses curtos."""
        ultimo = calendar.monthrange(self.ano, self.mes)[1]
        return date(self.ano, self.mes, min(dia, ultimo))

    def __str__(self) -> str:
        return f"{self.ano:04d}-{self.mes:02d}"


# --------------------------------------------------------------------------- #
# Documentos
# --------------------------------------------------------------------------- #


class TipoDocumento(str, Enum):
    AUTORIZACAO = "autorizacao"
    DEMONSTRATIVO = "demonstrativo"
    BOLETO = "boleto"
    NOTA_FISCAL = "nota_fiscal"
    XML_NFE = "xml_nfe"
    DESCONHECIDO = "desconhecido"

    @property
    def rotulo(self) -> str:
        return {
            "autorizacao": "Autorização de pagamento",
            "demonstrativo": "Demonstrativo da cobrança",
            "boleto": "Boleto",
            "nota_fiscal": "Nota fiscal",
            "xml_nfe": "XML da NF-e",
            "desconhecido": "Não identificado",
        }[self.value]


class OrigemDocumento(str, Enum):
    OUTLOOK = "outlook"
    ENTRADA_MANUAL = "entrada_manual"
    JA_NA_PASTA = "ja_na_pasta"


@dataclass
class Documento:
    """Um arquivo coletado que vai compor a autorização do mês."""

    caminho: Path
    tipo: TipoDocumento = TipoDocumento.DESCONHECIDO
    origem: OrigemDocumento = OrigemDocumento.ENTRADA_MANUAL
    # 0.0 a 1.0 — abaixo de LIMIAR_CONFIANCA o painel pede confirmação humana.
    confianca: float = 0.0
    # Por que foi classificado assim (aparece no painel).
    motivo: str = ""
    paginas: int | None = None
    # Metadados extraídos do próprio documento, quando dá.
    valor: float | None = None
    vencimento: date | None = None
    numero_documento: str | None = None
    remetente: str | None = None
    assunto_email: str | None = None

    @property
    def nome(self) -> str:
        return self.caminho.name


LIMIAR_CONFIANCA = 0.65


# --------------------------------------------------------------------------- #
# Conta (uma linha do checklist = uma pasta que recebe autorização por mês)
# --------------------------------------------------------------------------- #


@dataclass
class DadosAutorizacao:
    """Campos fixos da Autorização de Pagamento, por conta."""

    departamento: str | None = None
    pagante: str | None = None
    beneficiario: str | None = None
    forma_pgto: str | None = None
    meio_pgto: str | None = None
    setor: str | None = None
    motivo: str | None = None
    regional: str | None = None
    cooperativa: str | None = None
    rotulo_documento: str | None = None
    centro_custo: str | None = None
    natureza: str | None = None
    pgto_previsto: str | None = None
    # Pode conter {competencia_extenso}, {mes_anterior_extenso}, {cidade}.
    descricao: str | None = None

    @classmethod
    def de_dict(cls, dados: dict | None) -> DadosAutorizacao:
        dados = dados or {}
        campos = {f for f in cls.__dataclass_fields__}
        return cls(**{k: v for k, v in dados.items() if k in campos})


#: Listas do próprio template (nomes definidos FORMASPGTO e MEIOSPGTO na aba
#: `INFORMAÇÕES DE PROGRAMAÇÃO`). Ficam aqui porque o painel precisa delas
#: antes de abrir a planilha.
FORMAS_PAGAMENTO = ("À VISTA", "À PRAZO", "PARCELADO", "OUTRO")

#: `Carnê` não existe na lista do template — foi acrescentado a pedido.
MEIOS_PAGAMENTO = (
    "BOLETO", "DINHEIRO", "PIX", "CHEQUE",
    "TRANSFERENCIA", "CREDITO", "DEBITO", "Carnê",
)


def so_digitos(texto: str | None) -> str:
    return re.sub(r"\D", "", texto or "")


def cnpj_confere(cnpj: str | None) -> bool:
    """
    Confere os dois dígitos verificadores do CNPJ.

    Pega dígito trocado e número inventado, que é o erro que passa batido:
    formato certo, empresa errada. Não diz se o CNPJ existe na Receita.
    """
    numero = so_digitos(cnpj)
    if len(numero) != 14 or len(set(numero)) == 1:
        return False
    for tamanho in (12, 13):
        pesos = [(i % 8) + 2 for i in range(tamanho - 1, -1, -1)]
        resto = sum(int(d) * peso for d, peso in zip(numero, pesos)) % 11
        if int(numero[tamanho]) != (0 if resto < 2 else 11 - resto):
            return False
    return True


def formatar_cnpj(cnpj: str | None) -> str | None:
    """
    `11222333000181` -> `11.222.333/0001-81`.

    Só reformata quando os 14 dígitos fecham a conta. Número que não confere
    volta como foi digitado — reescrever um CNPJ suspeito esconderia o erro.
    """
    if not cnpj:
        return cnpj
    if not cnpj_confere(cnpj):
        return cnpj
    n = so_digitos(cnpj)
    return f"{n[:2]}.{n[2:5]}.{n[5:8]}/{n[8:12]}-{n[12:]}"


@dataclass
class Pagante:
    """Empresa que paga a fatura — o campo PAGANTE (B7) da autorização."""

    nome: str
    cnpj: str | None = None
    #: Como a empresa aparece na tabela de consulta do template. As razões
    #: sociais mudaram e a tabela não acompanhou; com isto a automação sabe
    #: se o `=VLOOKUP` de B9 ainda resolve ou se precisa escrever o CNPJ.
    nome_na_planilha: str | None = None
    banco: str | None = None
    agencia: str | None = None
    conta: str | None = None

    @property
    def id(self) -> str:
        return _identificador(self.nome)

    @property
    def rotulo(self) -> str:
        return f"{self.nome} — {self.cnpj}" if self.cnpj else self.nome

    @classmethod
    def de_dict(cls, dados: dict) -> Pagante:
        campos = {f for f in cls.__dataclass_fields__}
        return cls(**{k: v for k, v in (dados or {}).items() if k in campos})


@dataclass
class Beneficiario:
    """Quem recebe — o campo FORNECEDOR (C13) e os dados bancários do rodapé."""

    nome: str
    cnpj: str | None = None
    contato: str | None = None
    telefone: str | None = None
    banco: str | None = None
    agencia: str | None = None
    conta: str | None = None
    # Meio de pagamento NÃO mora aqui: muda de uma fatura para outra do mesmo
    # fornecedor, então é escolhido na etapa da autorização.

    @property
    def id(self) -> str:
        return _identificador(self.nome)

    @property
    def rotulo(self) -> str:
        return f"{self.nome} — {self.cnpj}" if self.cnpj else self.nome

    @classmethod
    def de_dict(cls, dados: dict) -> Beneficiario:
        campos = {f for f in cls.__dataclass_fields__}
        return cls(**{k: v for k, v in (dados or {}).items() if k in campos})


def _identificador(nome: str) -> str:
    """`FORNECEDOR S/A` -> `fornecedor-s-a`. Chave estável nos formulários."""
    limpo = unicodedata.normalize("NFKD", nome or "").encode("ascii", "ignore").decode()
    return re.sub(r"-+", "-", re.sub(r"[^a-z0-9]+", "-", limpo.lower())).strip("-")


@dataclass
class ConfigEmail:
    eh_operadora: bool = False
    cidade: str | None = None
    # Sobrescreve o nome da empresa no assunto, se preenchido.
    nome_empresa: str | None = None
    # Contas com o mesmo grupo saem num e-mail só, com um anexo por fatura.
    # É o caso de quem tem 4 faturas: uma mensagem, quatro PDFs.
    grupo: str | None = None


@dataclass
class ConfigColeta:
    """Como reconhecer os e-mails desta conta no Outlook."""

    remetentes: list[str] = field(default_factory=list)
    assunto_contem: list[str] = field(default_factory=list)
    # Termos que aparecem no anexo/corpo e distinguem esta conta das irmãs
    # (ex.: o número da linha, para separar as várias contas da operadora).
    identificadores: list[str] = field(default_factory=list)


@dataclass
class Conta:
    """Uma conta a pagar: fornecedor + (opcionalmente) unidade/linha/cidade."""

    id: str
    pasta: str
    subunidade: str | None = None
    #: Nome exibido no painel. Vazio = monta a partir de pasta/subunidade.
    #: É SÓ aparência: não entra em caminho de arquivo nem em nome de
    #: pasta. Renomear aqui nunca move nada no OneDrive.
    apelido: str | None = None
    ativo: bool = True
    formato_mes: str = "MM-AAAA"
    # Quantos meses a PASTA (e o nome do arquivo) ficam à frente da
    # competência. Quase toda conta guarda a autorização na pasta do próprio
    # mês de vencimento — aí isto é 0. Há uma exceção conhecida na base: a
    # fatura que vence em 04/08 mora em "2026-09", um mês à frente. Só o nome da
    # pasta e do arquivo andam; vencimento, valor, descrição e checklist
    # continuam presos à competência de verdade.
    #: Onde esta conta arquiva no OneDrive, quando NÃO é sob "Autorizações
    #: de pagamento". Caminho absoluto, ou relativo à raiz do OneDrive. Como
    #: já é a pasta do fornecedor, o nome dele não se repete no caminho: o
    #: destino fica `<raiz>/<mês>`. É o caso de quem mora fora da árvore.
    raiz_onedrive: str | None = None
    deslocamento_pasta: int = 0
    # Ex.: "TI - FORNECEDOR - {MM}-{AAAA}"
    padrao_nome_arquivo: str | None = None
    #: Molde da PASTA do mês, quando ela não é só o mês. Há fornecedor que usa
    #: "Fornecedor - {AAAA}-{MM}". Vazio = só o mês, no `formato_mes`.
    #: Note que pasta e arquivo podem discordar na ordem: nessa mesma conta a
    #: pasta é AAAA-MM e o arquivo é MM-AAAA.
    padrao_nome_pasta: str | None = None
    aba_autorizacao: str = "AUTORIZAÇÃO"
    # Caminho relativo (dentro de "Autorizações de pagamento") do xlsx usado
    # como base para gerar o do mês novo.
    modelo_base: str | None = None
    vencimento_dia: int | None = None
    # Quanto costuma vir na fatura. Serve para pré-preencher a autorização
    # quando o valor não foi extraído do boleto, e para o painel avisar quando
    # o valor do mês fugir muito do de sempre.
    valor_estimado: float | None = None
    autorizacao: DadosAutorizacao = field(default_factory=DadosAutorizacao)
    email: ConfigEmail = field(default_factory=ConfigEmail)
    coleta: ConfigColeta = field(default_factory=ConfigColeta)
    # Textos da coluna FORNECEDOR na planilha CONTAS E ACESSOS.
    chaves_planilha: list[str] = field(default_factory=list)
    observacao: str = ""

    @property
    def rotulo(self) -> str:
        """Como a conta aparece no painel — o apelido, se houver."""
        if self.apelido:
            return self.apelido
        return f"{self.pasta} — {self.subunidade}" if self.subunidade else self.pasta

    @property
    def rotulo_das_pastas(self) -> str:
        """O nome derivado do caminho, ignorando o apelido. Para a tela de
        Parâmetros mostrar de onde o caminho sai."""
        return f"{self.pasta} — {self.subunidade}" if self.subunidade else self.pasta

    @property
    def nome_empresa(self) -> str:
        return self.email.nome_empresa or self.pasta

    def competencia_da_pasta(self, competencia: Competencia) -> Competencia:
        """
        O mês que aparece na pasta e no nome do arquivo.

        Igual à competência para quase todo mundo; adiantado em
        `deslocamento_pasta` meses para quem arquiva fora de fase.
        """
        return competencia.somar(self.deslocamento_pasta)

    def nome_pasta_do_mes(self, competencia: Competencia) -> str:
        """
        Nome da subpasta do mês — só o mês, ou o molde da conta.

        `padrao_nome_pasta` existe porque nem todo fornecedor batizou a pasta
        igual: a maioria usa "09-2026", e há quem use "Fornecedor - 2026-09".
        """
        na_pasta = self.competencia_da_pasta(competencia)
        if self.padrao_nome_pasta:
            return (
                self.padrao_nome_pasta
                .replace("{AAAA}", f"{na_pasta.ano:04d}")
                .replace("{MM}", f"{na_pasta.mes:02d}")
            )
        return na_pasta.nome_pasta(self.formato_mes)

    def nome_base_arquivo(self, competencia: Competencia) -> str:
        """Nome (sem extensão) do xlsx e do PDF final do mês."""
        na_pasta = self.competencia_da_pasta(competencia)
        if self.padrao_nome_arquivo:
            return (
                self.padrao_nome_arquivo
                .replace("{AAAA}", f"{na_pasta.ano:04d}")
                .replace("{MM}", f"{na_pasta.mes:02d}")
            )
        alvo = self.subunidade or self.pasta
        return f"{alvo} - {na_pasta.nome_pasta(self.formato_mes)}"

    def caminho_relativo(self, competencia: Competencia) -> Path:
        """
        Pasta de destino, relativa à raiz que esta conta usa.

        Com `raiz_onedrive`, a raiz JÁ é a pasta do fornecedor — repetir o
        nome dela daria `.../Fornecedor/Fornecedor/07-2026`. Sem ela, a raiz é
        "Autorizações de pagamento" e o nome do fornecedor abre o caminho.
        """
        na_pasta = self.competencia_da_pasta(competencia)
        partes = [] if self.raiz_onedrive else [self.pasta]
        partes.append(self.nome_pasta_do_mes(competencia))
        if self.subunidade:
            partes.append(self.subunidade)
        return Path(*partes)


# --------------------------------------------------------------------------- #
# Estado do processamento
# --------------------------------------------------------------------------- #


class Etapa(str, Enum):
    COLETA = "coleta"
    PASTA = "pasta"
    AUTORIZACAO = "autorizacao"
    PDF_AUTORIZACAO = "pdf_autorizacao"
    PDF_FINAL = "pdf_final"
    PUBLICACAO = "publicacao"
    EMAIL = "email"
    CHECKLIST = "checklist"

    @property
    def rotulo(self) -> str:
        return {
            "coleta": "Coletar documentos",
            "pasta": "Criar pasta do mês",
            "autorizacao": "Preencher autorização",
            "pdf_autorizacao": "Exportar autorização em PDF",
            "pdf_final": "Montar PDF único",
            "publicacao": "Publicar no OneDrive",
            "email": "Criar rascunho do e-mail",
            "checklist": "Marcar verde no checklist",
        }[self.value]


class Situacao(str, Enum):
    PENDENTE = "pendente"
    OK = "ok"
    ATENCAO = "atencao"   # concluído, mas precisa de olho humano
    ERRO = "erro"
    PULADO = "pulado"


# Ordem canônica do pipeline. Fica aqui, e não no orquestrador, para que
# módulos de baixo nível (estado) possam usá-la sem importar o de cima.
ORDEM_ETAPAS_PADRAO = [
    Etapa.COLETA,
    Etapa.PASTA,
    Etapa.AUTORIZACAO,
    Etapa.PDF_AUTORIZACAO,
    Etapa.PDF_FINAL,
    Etapa.PUBLICACAO,
    Etapa.EMAIL,
    Etapa.CHECKLIST,
]


@dataclass
class ResultadoEtapa:
    etapa: Etapa
    situacao: Situacao
    mensagem: str = ""
    detalhes: dict = field(default_factory=dict)
    # Arquivos produzidos ou tocados por esta etapa.
    artefatos: list[Path] = field(default_factory=list)
    #: Quando esta etapa foi gravada. Preenchido na leitura do banco; em
    #: memória fica vazio. É o que permite saber QUANDO a fatura saiu, e não
    #: só que saiu — a lista do mês usa isso para julgar atraso.
    atualizado_em: datetime | None = None

    @property
    def ok(self) -> bool:
        return self.situacao in (Situacao.OK, Situacao.ATENCAO, Situacao.PULADO)

    @classmethod
    def sucesso(cls, etapa: Etapa, mensagem: str = "", **kw) -> ResultadoEtapa:
        return cls(etapa, Situacao.OK, mensagem, **kw)

    @classmethod
    def atencao(cls, etapa: Etapa, mensagem: str, **kw) -> ResultadoEtapa:
        return cls(etapa, Situacao.ATENCAO, mensagem, **kw)

    @classmethod
    def erro(cls, etapa: Etapa, mensagem: str, **kw) -> ResultadoEtapa:
        return cls(etapa, Situacao.ERRO, mensagem, **kw)

    @classmethod
    def pulado(cls, etapa: Etapa, mensagem: str, **kw) -> ResultadoEtapa:
        return cls(etapa, Situacao.PULADO, mensagem, **kw)


@dataclass
class Processamento:
    """Estado de uma conta em uma competência."""

    conta_id: str
    competencia: Competencia
    resultados: dict[Etapa, ResultadoEtapa] = field(default_factory=dict)
    documentos: list[Documento] = field(default_factory=list)
    valor: float | None = None
    vencimento: date | None = None
    numero_documento: str | None = None
    #: Texto do campo OBSERVAÇÕES quando você reescreveu no painel. Vale só
    #: para este mês — o molde da conta continua no registro.
    descricao: str | None = None
    #: Escolhas feitas na etapa da autorização, também válidas só para o mês.
    pagante: str | None = None
    beneficiario: str | None = None
    meio_pgto: str | None = None
    forma_pgto: str | None = None
    pdf_final: Path | None = None
    xlsx: Path | None = None
    email_enviado: bool = False
    enviado_em: date | None = None
    enviado_por_fora: bool = False

    def registrar(self, resultado: ResultadoEtapa) -> ResultadoEtapa:
        self.resultados[resultado.etapa] = resultado
        return resultado

    def situacao_de(self, etapa: Etapa) -> Situacao:
        r = self.resultados.get(etapa)
        return r.situacao if r else Situacao.PENDENTE

    @property
    def concluido(self) -> bool:
        return all(
            self.situacao_de(e) in (Situacao.OK, Situacao.ATENCAO, Situacao.PULADO)
            for e in Etapa
        )

    @property
    def tem_erro(self) -> bool:
        return any(r.situacao is Situacao.ERRO for r in self.resultados.values())

    def documentos_por_tipo(self, tipo: TipoDocumento) -> list[Documento]:
        return [d for d in self.documentos if d.tipo is tipo]
