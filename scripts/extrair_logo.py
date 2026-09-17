"""
Extrai a logomarca da empresa do template de autorização e gera os arquivos
que o painel usa: `painel/static/logo.png`, `painel/static/logo-escura.png`
e `painel/static/favicon.ico`.

A logo já vive dentro do xlsx de autorização (`xl/media/`), em PNG com fundo
transparente. Tirar de lá evita depender de alguém achar o arquivo original —
e garante que o painel use exatamente a mesma arte do documento que ele gera.

    python scripts/extrair_logo.py

Os arquivos gerados ficam FORA do git (são marca da empresa). Em uma máquina
nova, rode este script uma vez; sem ele o painel mostra o nome em texto.
"""

from __future__ import annotations

import io
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from PIL import Image, ImageDraw  # noqa: E402

from automacao.nucleo.config import ambiente, contas  # noqa: E402

DESTINO = Path(__file__).resolve().parents[1] / "painel" / "static"

#: Menor imagem com transparência dentro do xlsx é a logo atual; as maiores
#: em RGB são a marca antiga, que ainda sobrevive em alguns modelos.
def _melhor_logo(pacote: zipfile.ZipFile) -> tuple[str, Image.Image] | None:
    candidatas: list[tuple[int, str, Image.Image]] = []
    for nome in pacote.namelist():
        if not nome.startswith("xl/media/"):
            continue
        try:
            imagem = Image.open(io.BytesIO(pacote.read(nome)))
        except Exception:
            continue
        if imagem.mode not in ("RGBA", "LA", "P"):
            continue
        imagem = imagem.convert("RGBA")
        if imagem.getbbox() is None:
            continue
        largura, altura = imagem.size
        # A logo é larga e baixa; ícone solto e carimbo não são.
        if largura < altura * 1.5:
            continue
        candidatas.append((largura * altura, nome, imagem))
    if not candidatas:
        return None
    _, nome, imagem = max(candidatas)
    return nome, imagem


def _fim_da_primeira_letra(logo: Image.Image) -> int:
    """
    Coluna onde acaba o primeiro glifo, achada pelos vazios do canal alfa.

    O PNG não traz metadado de letra; o que dá para medir é onde a tinta
    some. O primeiro intervalo em branco depois do começo é o fim do "g".
    """
    largura, altura = logo.size
    alfa = logo.getchannel("A")
    tem_tinta = [
        alfa.crop((x, 0, x + 1, altura)).getbbox() is not None
        for x in range(largura)
    ]
    comecou = False
    for x, tinta in enumerate(tem_tinta):
        if tinta:
            comecou = True
        elif comecou:
            return x
    return largura


#: A marca tem duas tintas: o azul-marinho do "gol" e o laranja do "plus".
AZUL = (16, 42, 77)
LARANJA = (243, 110, 33)


def _versao_escura(logo: Image.Image) -> Image.Image:
    """
    A mesma arte para fundo escuro: o azul vira branco, o laranja fica.

    Inverter a imagem inteira (o caminho curto) apagaria o laranja junto e a
    marca sairia num branco chapado. Aqui cada pixel é classificado pela tinta
    de que está mais perto — e só a metade azul é trocada.

    O alfa é preservado pixel a pixel, então a borda suavizada continua
    suavizada: o que era azul a 40% vira branco a 40%.
    """
    escura = logo.convert("RGBA").copy()
    tela = escura.load()
    for y in range(escura.height):
        for x in range(escura.width):
            r, g, b, a = tela[x, y]
            if a == 0:
                continue
            do_azul = sum((c - t) ** 2 for c, t in zip((r, g, b), AZUL))
            do_laranja = sum((c - t) ** 2 for c, t in zip((r, g, b), LARANJA))
            if do_azul <= do_laranja:
                tela[x, y] = (255, 255, 255, a)
    return escura


def _modelo_disponivel() -> Path | None:
    """Primeiro `modelo_base` que exista no disco."""
    raiz = ambiente().caminhos.autorizacoes
    for conta in contas():
        if not conta.modelo_base:
            continue
        caminho = raiz / conta.modelo_base
        if caminho.is_file():
            return caminho
    return None


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    modelo = _modelo_disponivel()
    if modelo is None:
        print("Nenhum modelo de autorização encontrado — nada a extrair.")
        return 1

    # A planilha costuma estar aberta no Excel; lemos uma cópia.
    copia = Path(tempfile.gettempdir()) / "extrair-logo.xlsx"
    shutil.copy2(modelo, copia)

    with zipfile.ZipFile(copia) as pacote:
        achado = _melhor_logo(pacote)
    copia.unlink(missing_ok=True)

    if achado is None:
        print(f"Não achei logo com transparência em {modelo.name}.")
        return 1

    nome_interno, logo = achado
    logo = logo.crop(logo.getbbox())          # tira a margem vazia do PNG
    DESTINO.mkdir(parents=True, exist_ok=True)

    arquivo_logo = DESTINO / "logo.png"
    logo.save(arquivo_logo, optimize=True)

    arquivo_escura = DESTINO / "logo-escura.png"
    _versao_escura(logo).save(arquivo_escura, optimize=True)

    # Ícone da aba: só a primeira letra. Em 16px a palavra inteira vira um
    # borrão, e o "g" com o traço embaixo já identifica a marca. O corte sai
    # das colunas com tinta, não de uma fração chutada da largura — assim
    # nenhuma letra fica cortada ao meio.
    marca = logo.crop((0, 0, _fim_da_primeira_letra(logo), logo.height))
    lado = int(max(marca.size) * 1.34)       # respiro em volta da letra

    # Fundo branco arredondado embutido: o "g" é azul-marinho e desapareceria
    # na barra de abas do navegador em tema escuro. Com o fundo, o ícone lê em
    # qualquer tema — e é assim que a marca costuma ser aplicada.
    quadrado = Image.new("RGBA", (lado, lado), (0, 0, 0, 0))
    canto = ImageDraw.Draw(quadrado)
    canto.rounded_rectangle((0, 0, lado - 1, lado - 1),
                            radius=int(lado * 0.22), fill=(255, 255, 255, 255))
    quadrado.alpha_composite(marca, ((lado - marca.width) // 2,
                                     (lado - marca.height) // 2))
    arquivo_icone = DESTINO / "favicon.ico"
    quadrado.save(arquivo_icone, sizes=[(16, 16), (32, 32), (48, 48), (64, 64)])

    print(f"origem : {modelo.name} ({nome_interno})")
    print(f"logo   : {arquivo_logo}  {logo.size}")
    print(f"escura : {arquivo_escura}  (azul → branco, laranja mantido)")
    print(f"ícone  : {arquivo_icone}  {quadrado.size} em 4 tamanhos")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
