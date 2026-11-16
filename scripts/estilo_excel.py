"""Aplica o estilo padrão "planilhas-design" a uma aba do openpyxl.

Layout esperado: título em A1, subtítulo em A2, linha 3 vazia, cabeçalho na
linha 4 e dados logo abaixo. Se o título/subtítulo forem passados, o script
os escreve.

Uso:
    from estilo_excel import aplicar_estilo
    aplicar_estilo(ws, titulo="Despesas", subtitulo="2026",
                   formatos={"Valor": "moeda", "Data": "data"},
                   linha_total=True, colunas_entrada=["Meta"])
"""
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

COR = {
    "primaria": "12424D",
    "destaque": "FFB745",
    "primaria_clara": "E7EFF0",
    "zebra": "F5F8F8",
    "borda": "D5E0E2",
    "texto": "1E2B2E",
    "texto_suave": "5F7378",
    "branco": "FFFFFF",
    "positivo": "2E7D32",
    "negativo": "C62828",
    "alerta": "B26A00",
    "entrada": "FFF1D6",
    "entrada_texto": "12424D",
    "texto_faixa": "B7CDD2",  # subtítulo sobre a faixa petróleo do topo
}
CORES_GRAFICO = ["12424D", "FFB745", "3E8A96", "8CC2C9", "D9822B", "A7B4B7"]
FONTE = "Poppins"        # texto simples (corpo, notas, total)
FONTE_TITULO = "Sora"    # títulos: A1 e cabeçalho da tabela
# As duas precisam estar instaladas no computador para o Excel exibi-las.

FORMATOS = {
    "moeda": '"R$" #,##0.00;-"R$" #,##0.00',
    "inteiro": "#,##0",
    "decimal": "#,##0.00",
    "percentual": "0.0%",
    "variacao": "+0.0%;-0.0%;0.0%",
    # negativos ficam vermelho da paleta (#C62828) via formatação condicional,
    # porque o [Red] do formato de número é o vermelho puro do Excel
    "data": "dd/mm/yyyy",
    "mes": "mmm/yyyy",
}
NUMERICOS = {"moeda", "inteiro", "decimal", "percentual", "variacao"}
CENTRO = {"data", "mes", "status", "centro"}  # "centro": qualquer texto curto centralizado

# Cor da fonte por status (negrito). Estenda com o parâmetro `status_extra`.
STATUS = {
    "positivo": ["OK", "Concluído", "Concluido", "Pago", "Ativo", "Em dia", "Aprovado", "Novo"],
    "alerta": ["Pendente", "Atenção", "Atencao", "Em andamento", "Em análise", "Em renovação"],
    "negativo": ["Atrasado", "Erro", "Vencido", "Reprovado"],
    "texto_suave": ["Cancelado", "Inativo", "Encerrado", "Arquivado"],  # neutro: não é "ruim"
}


def _fill(hex_):
    return PatternFill("solid", start_color=hex_, end_color=hex_)


def _fonte(tamanho=10, cor="texto", titulo=False, **kw):
    return Font(name=FONTE_TITULO if titulo else FONTE, size=tamanho, color=COR[cor], **kw)


def aplicar_estilo(ws, titulo=None, subtitulo=None, linha_cabecalho=4,
                   formatos=None, linha_total=False, colunas_total=None,
                   colunas_entrada=None, coluna_status=None, nota=None,
                   congelar_primeira_coluna=False, cor_aba=False, logo="auto",
                   filtro=None, status_extra=None, linha_nota=None, cartoes=None,
                   linha_cartoes=4, congelar=None):
    """Formata `ws` no padrão.

    formatos: {"Nome do cabeçalho": "moeda"|"inteiro"|"decimal"|"percentual"|
               "variacao"|"data"|"mes"|"status"|"texto"}
    linha_total: adiciona linha "Total" somando as colunas numéricas
                 (ou só `colunas_total`, se informado).
    colunas_entrada: cabeçalhos cujas células são inputs do usuário.
    coluna_status: cabeçalho da coluna de status (cor condicional).
    logo: caminho de imagem; "auto" usa ../assets/logo.png se existir; None desliga.
          Use a logo só na aba principal (logo=None nas demais).
    filtro: True/False; None = automático (liga só em listas com 8+ linhas).
    status_extra: {"positivo"|"alerta"|"negativo"|"texto_suave": ["Termo", ...]}.
    congelar: True/False; None = automático (só em listas com 8+ linhas, como o filtro).
              Em abas de resumo/painel curtas não congele: corta a primeira linha da tabela.
    linha_nota: linha da nota (padrão: 2 linhas abaixo do fim da tabela).
    cartoes: cartões de resultado (ver cartoes_resultado) — divididos em blocos de
             largura IGUAL na largura da tabela. Use linha_cabecalho=8 com eles.
    Nenhum texto passa da borda direita da tabela: título/subtítulo são presos à largura
    da faixa e a nota quebra linha dentro da largura da tabela (altura ajustada).
    Colunas de entrada têm prioridade sobre a zebra (o âmbar-claro aparece inteiro).
    """
    formatos = dict(formatos or {})
    if coluna_status:
        formatos.setdefault(coluna_status, "status")
    colunas_entrada = set(colunas_entrada or [])
    h = linha_cabecalho

    # Título e subtítulo (a faixa petróleo é pintada depois de achar a última coluna)
    if titulo is not None:
        ws.cell(1, 1, titulo)
    if subtitulo is not None:
        ws.cell(2, 1, subtitulo)

    # Descobrir colunas e última linha de dados
    ultima_col = 0
    for c in range(1, ws.max_column + 1):
        if ws.cell(h, c).value not in (None, ""):
            ultima_col = c
    if ultima_col == 0:
        raise ValueError(f"Nenhum cabeçalho encontrado na linha {h}.")
    ultima_lin = h
    for r in range(h + 1, ws.max_row + 1):
        if any(ws.cell(r, c).value not in (None, "") for c in range(1, ultima_col + 1)):
            ultima_lin = r

    # Faixa do topo: linhas 1–2 em petróleo na largura da tabela
    for r in (1, 2):
        for c in range(1, ultima_col + 1):
            ws.cell(r, c).fill = _fill(COR["primaria"])
    ws.cell(1, 1).font = _fonte(16, "branco", titulo=True, bold=True)
    ws.cell(1, 1).alignment = Alignment(horizontal="left", vertical="bottom", indent=1)
    ws.cell(2, 1).font = _fonte(10, "texto_faixa", italic=True)
    ws.cell(2, 1).alignment = Alignment(horizontal="left", vertical="top", indent=1)
    ws.row_dimensions[1].height = 34
    ws.row_dimensions[2].height = 24

    cabecalhos = {c: str(ws.cell(h, c).value) for c in range(1, ultima_col + 1)}
    tipo = {c: formatos.get(nome, "texto") for c, nome in cabecalhos.items()}

    borda_fina = Side(style="thin", color=COR["borda"])
    borda_total = Side(style="medium", color=COR["primaria"])
    borda_destaque = Side(style="medium", color=COR["destaque"])

    def alinhamento(c, wrap=False):
        t = tipo[c]
        horiz = "right" if t in NUMERICOS else "center" if t in CENTRO else "left"
        # recuo 1 para o conteúdo não encostar na borda da célula
        recuo = 0 if horiz == "center" else 1
        return Alignment(horizontal=horiz, vertical="center", wrap_text=wrap, indent=recuo)

    usar_filtro = filtro if filtro is not None else (ultima_lin - h) >= 8

    # Cabeçalho
    for c in cabecalhos:
        cel = ws.cell(h, c)
        cel.font = _fonte(10, "branco", titulo=True, bold=True)
        cel.fill = _fill(COR["primaria"])
        al = alinhamento(c, wrap=True)
        if al.horizontal == "right" and usar_filtro:
            # recuo maior para o texto não ficar embaixo do botão do filtro
            al = Alignment(horizontal="right", vertical="center", wrap_text=True, indent=3)
        cel.alignment = al
        cel.border = Border(bottom=borda_destaque)
    ws.row_dimensions[h].height = 24

    # Corpo
    for r in range(h + 1, ultima_lin + 1):
        zebra = (r - h) % 2 == 0
        for c in cabecalhos:
            cel = ws.cell(r, c)
            entrada = cabecalhos[c] in colunas_entrada
            cel.font = _fonte(10, "entrada_texto" if entrada else "texto")
            if entrada:
                cel.fill = _fill(COR["entrada"])
            elif zebra:
                cel.fill = _fill(COR["zebra"])
            else:
                cel.fill = PatternFill(fill_type=None)
            cel.border = Border(bottom=borda_fina)
            cel.alignment = alinhamento(c)
            if tipo[c] in FORMATOS:
                cel.number_format = FORMATOS[tipo[c]]
        ws.row_dimensions[r].height = 20

    # Linha de total
    if linha_total and ultima_lin > h:
        rt = ultima_lin + 1
        alvo = set(colunas_total) if colunas_total else None
        for c in cabecalhos:
            cel = ws.cell(rt, c)
            if c == 1:
                cel.value = "Total"
            elif (alvo is None and tipo[c] in {"moeda", "inteiro", "decimal"}) or \
                    (alvo is not None and cabecalhos[c] in alvo):
                L = get_column_letter(c)
                cel.value = f"=SUM({L}{h + 1}:{L}{ultima_lin})"
                if tipo[c] in FORMATOS:
                    cel.number_format = FORMATOS[tipo[c]]
            cel.font = _fonte(10, "primaria", bold=True)
            cel.fill = _fill(COR["primaria_clara"])
            cel.border = Border(top=borda_total, bottom=borda_fina)
            cel.alignment = alinhamento(c)
        ws.row_dimensions[rt].height = 22
        fim_tabela = rt
    else:
        fim_tabela = ultima_lin

    # Nota de rodapé (mesclada e com quebra de linha depois de calcular as larguras)
    linha_nota = linha_nota or fim_tabela + 2
    if nota:
        cel = ws.cell(linha_nota, 1, nota)
        cel.font = _fonte(9, "texto_suave", italic=True)

    # Negativos em vermelho da paleta nas colunas numéricas (inclui a linha de total)
    for c in cabecalhos:
        if tipo[c] in NUMERICOS:
            L = get_column_letter(c)
            ws.conditional_formatting.add(
                f"{L}{h + 1}:{L}{fim_tabela}",
                CellIsRule(operator="lessThan", formula=["0"],
                           font=Font(name=FONTE, color=COR["negativo"])))

    # Status: cor condicional na fonte
    if coluna_status and coluna_status in cabecalhos.values():
        c = next(k for k, v in cabecalhos.items() if v == coluna_status)
        L = get_column_letter(c)
        faixa = f"{L}{h + 1}:{L}{ultima_lin}"
        cores_status = {k: list(v) for k, v in STATUS.items()}
        for k, v in (status_extra or {}).items():
            cores_status.setdefault(k, []).extend(v)
        for cor, palavras in cores_status.items():
            cond = ",".join(f'{L}{h + 1}="{p}"' for p in palavras)
            ws.conditional_formatting.add(
                faixa, FormulaRule(formula=[f"OR({cond})"],
                                   font=Font(name=FONTE, color=COR[cor], bold=True)))

    # Larguras de coluna
    for c in cabecalhos:
        maior = 0
        for r in range(h + 1, fim_tabela + 1):
            v = ws.cell(r, c).value
            if v is None or (isinstance(v, str) and v.startswith("=")):
                continue
            n = 12 if tipo[c] in NUMERICOS else len(str(v))
            maior = max(maior, n)
        if tipo[c] in NUMERICOS and maior == 0:
            maior = 12  # coluna só de fórmulas
        # Poppins é ~15% mais larga que Arial; +4 cobre o recuo
        conteudo = round(maior * 1.15) + 4
        # cabeçalho (Sora negrito) + espaço do botão do filtro
        # texto centralizado precisa de espaço para o botão dos DOIS lados
        folga = (10 if tipo[c] in CENTRO else 6) if usar_filtro else 3
        cab = round(len(cabecalhos[c]) * 1.1) + folga
        ws.column_dimensions[get_column_letter(c)].width = max(10, min(50, max(conteudo, cab)))

    # Texto maior que a coluna: quebra linha e a linha cresce para caber inteiro
    # (nada cortado e nada vazando para a coluna ao lado ou para fora da tabela)
    import math
    for r in range(h + 1, fim_tabela + 1):
        linhas_max = 1
        for c in cabecalhos:
            v = ws.cell(r, c).value
            if not isinstance(v, str) or v.startswith("="):
                continue
            largura = ws.column_dimensions[get_column_letter(c)].width or 8.43
            # Poppins 10: ~1 caractere por unidade de largura; -2 desconta o recuo
            n = math.ceil(len(v) / max(1, largura - 2))
            if n > 1:
                ws.cell(r, c).alignment = alinhamento(c, wrap=True)
                linhas_max = max(linhas_max, n)
        if linhas_max > 1:
            ws.row_dimensions[r].height = max(20, linhas_max * 15 + 5)

    # Cartões de resultado: blocos de largura igual (pode alargar colunas um pouco)
    if cartoes:
        cartoes_resultado(ws, cartoes, linha=linha_cartoes, ultima_col=ultima_col)

    # Textos fora da tabela presos à largura dela (nada passa da borda direita)
    for r in (1, 2):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ultima_col)
    if nota:
        quebrar_na_largura(ws, linha_nota, 1, ultima_col, tamanho=9)

    # Logo no canto superior direito da área do título (linhas 1–2)
    if logo == "auto":
        import os
        pasta = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "assets")
        # o topo é petróleo → usa a versão clara da logo
        opcoes = [os.path.join(pasta, n) for n in ("logo.png", "logo-escura.png")]
        logo = next((o for o in opcoes if os.path.exists(o)), None)
    if logo:
        from openpyxl.drawing.image import Image
        img = Image(logo)
        altura = 40  # px — centralizada na faixa das linhas 1–2 (~77 px)
        img.width, img.height = round(img.width * altura / img.height), altura
        # Alinha a borda direita da logo com a borda direita da tabela
        from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
        from openpyxl.drawing.xdr import XDRPositiveSize2D
        from openpyxl.utils.units import pixels_to_EMU

        def px(col):
            w = ws.column_dimensions[get_column_letter(col)].width or 8.43
            return int(w * 7 + 5)
        margem = 12  # px de respiro até a borda direita da faixa
        col, acumulado = ultima_col, px(ultima_col) - margem
        while acumulado < img.width and col > 1:
            col -= 1
            acumulado += px(col)
        deslocamento = max(0, acumulado - img.width)
        img.anchor = OneCellAnchor(
            _from=AnchorMarker(col=col - 1, colOff=pixels_to_EMU(deslocamento), row=0,
                               rowOff=pixels_to_EMU(18)),
            ext=XDRPositiveSize2D(pixels_to_EMU(img.width), pixels_to_EMU(img.height)))
        ws.add_image(img)

    # Ajustes de aba
    ws.sheet_view.showGridLines = False
    if congelar if congelar is not None else usar_filtro:
        primeira = ws.cell(h + 1, 2 if congelar_primeira_coluna else 1)
        ws.freeze_panes = primeira
        # seleção inicial na 1ª célula da área rolável: com a seleção em A1 (fora dela)
        # o Excel do Mac abre a aba rolada e esconde a primeira linha de dados
        for sel in ws.sheet_view.selection:
            sel.activeCell = sel.sqref = primeira.coordinate
    if usar_filtro:
        ws.auto_filter.ref = f"A{h}:{get_column_letter(ultima_col)}{ultima_lin}"
    if cor_aba:  # desligado por padrão: o Excel do Mac desenha a cor da aba com degradê
        ws.sheet_properties.tabColor = COR["primaria"]
    return ws


def quebrar_na_largura(ws, linha, col_ini, col_fim, tamanho=9, horizontal="left"):
    """Mescla linha[col_ini:col_fim], liga a quebra de texto e ajusta a altura da linha
    para o texto caber inteiro DENTRO da largura — use para qualquer texto solto
    (notas, observações, explicações) que poderia passar da borda da tabela."""
    import math
    cel = ws.cell(linha, col_ini)
    texto = str(cel.value or "")
    if col_fim > col_ini:
        ws.merge_cells(start_row=linha, start_column=col_ini, end_row=linha, end_column=col_fim)
    largura = sum((ws.column_dimensions[get_column_letter(c)].width or 8.43)
                  for c in range(col_ini, col_fim + 1))
    # Poppins 9pt cabe ~1,05 caractere por unidade de largura (margem de segurança)
    por_linha = max(1, largura * 1.05 * 9 / tamanho)
    linhas = max(1, math.ceil(len(texto) / por_linha))
    cel.alignment = Alignment(horizontal=horizontal, vertical="top", wrap_text=True)
    ws.row_dimensions[linha].height = linhas * (tamanho * 1.6) + 4
    return linhas


def _dividir_em_blocos(ws, n, ultima_col):
    """Escolhe n grupos de colunas contíguas (A..ultima_col) com larguras o mais
    parecidas possível."""
    from itertools import combinations
    larg = [ws.column_dimensions[get_column_letter(c)].width or 8.43 for c in range(1, ultima_col + 1)]
    alvo = sum(larg) / n
    melhor, melhor_erro = None, None
    for cortes in combinations(range(1, ultima_col), n - 1):
        lim = (0,) + cortes + (ultima_col,)
        grupos = [(lim[i] + 1, lim[i + 1]) for i in range(n)]
        erro = max(abs(sum(larg[a - 1:b]) - alvo) for a, b in grupos)
        if melhor_erro is None or erro < melhor_erro:
            melhor, melhor_erro = grupos, erro
    return melhor


def _igualar_blocos(ws, grupos):
    """Alarga as colunas dos blocos mais estreitos até todos terem a mesma largura
    (nunca estreita colunas, para não cortar conteúdo)."""
    def w(c):
        return ws.column_dimensions[get_column_letter(c)].width or 8.43
    larguras = [sum(w(c) for c in range(a, b + 1)) for a, b in grupos]
    maior = max(larguras)
    for (a, b), atual in zip(grupos, larguras):
        extra = (maior - atual) / (b - a + 1)
        for c in range(a, b + 1):
            ws.column_dimensions[get_column_letter(c)].width = round(w(c) + extra, 2)


def cartoes_resultado(ws, cartoes, linha=4, ultima_col=None, igualar=True):
    """Faixa de cartões com o RESULTADO FINAL no topo da aba (abaixo da faixa petróleo),
    para quem abrir a planilha entender o número principal sem ler a tabela.

    cartoes: lista de dicts, um por cartão, da esquerda para a direita:
        {"colunas": (1, 2),            # coluna inicial e final do cartão
         "rotulo": "Custo mensal atual",
         "valor": "=C12",              # fórmula ou número
         "formato": "moeda",           # chave de FORMATOS (opcional)
         "detalhe": "6 itens ativos",  # texto ou fórmula pequena embaixo (opcional)
         "destaque": True}             # só o número principal: fundo âmbar
    Ocupa 3 linhas (rótulo, valor, detalhe); deixe a tabela começar 2 linhas depois
    (com linha=4, use linha_cabecalho=8 em aplicar_estilo).
    Sem "colunas" (e com ultima_col), os cartões são distribuídos sozinhos em blocos de
    largura igual. igualar=True alarga colunas para os blocos ficarem idênticos.
    Prefira passar os cartões para aplicar_estilo(cartoes=...), que chama esta função
    na hora certa (antes de posicionar a logo).
    """
    if any("colunas" not in k for k in cartoes):
        grupos = _dividir_em_blocos(ws, len(cartoes), ultima_col)
        cartoes = [dict(k, colunas=g) for k, g in zip(cartoes, grupos)]
    if igualar:
        _igualar_blocos(ws, [k["colunas"] for k in cartoes])
    branco = Side(style="thick", color=COR["branco"])
    for k in cartoes:
        c0, c1 = k["colunas"]
        fundo = COR["destaque"] if k.get("destaque") else COR["primaria_clara"]
        cor_rot = "primaria" if k.get("destaque") else "texto_suave"
        for r in range(linha, linha + 3):
            for c in range(c0, c1 + 1):
                cel = ws.cell(r, c)
                cel.fill = _fill(fundo)
                # bordas brancas grossas nas laterais = espaço entre os cartões
                cel.border = Border(left=branco if c == c0 else None, right=branco if c == c1 else None)
            if c1 > c0:  # cada linha do cartão mesclada: o texto nunca invade o cartão vizinho
                ws.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c1)
        rot = ws.cell(linha, c0, k["rotulo"].upper())
        rot.font = _fonte(9, cor_rot, bold=True)
        rot.alignment = Alignment(horizontal="center", vertical="bottom", wrap_text=True)
        val = ws.cell(linha + 1, c0, k["valor"])
        val.font = _fonte(20, "primaria", titulo=True, bold=True)
        val.alignment = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
        if k.get("formato"):
            val.number_format = FORMATOS[k["formato"]]
        if k.get("detalhe") is not None:
            det = ws.cell(linha + 2, c0, k["detalhe"])
            det.font = _fonte(9, cor_rot)
            det.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
    ws.row_dimensions[linha].height = 22
    ws.row_dimensions[linha + 1].height = 34
    ws.row_dimensions[linha + 2].height = 32  # cabe até 2 linhas de detalhe


def _texto_rico(tamanho, cor, negrito=False, fonte=FONTE):
    from openpyxl.chart.text import RichText
    from openpyxl.drawing.text import CharacterProperties, Font as DFont, Paragraph, ParagraphProperties
    cp = CharacterProperties(sz=int(tamanho * 100), b=negrito, solidFill=cor, latin=DFont(typeface=fonte))
    return RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)])


def estilizar_grafico(grafico, titulo=None, formato_valor=None, rotulos=False):
    """Aplica o padrão a um gráfico openpyxl: paleta na ordem, título em Sora 14
    petróleo, eixos em Poppins 9, sem linhas de grade e legenda embaixo
    (ou nenhuma, se houver uma série só).

    Depois, posicione com posicionar_grafico() (mesma largura da tabela).
    """
    from openpyxl.chart.label import DataLabelList
    from openpyxl.chart.title import Title
    from openpyxl.chart.text import Text

    for i, serie in enumerate(grafico.series):
        cor = CORES_GRAFICO[i % len(CORES_GRAFICO)]
        serie.graphicalProperties.solidFill = cor
        serie.graphicalProperties.line.solidFill = cor
    if titulo:
        rico = _texto_rico(14, COR["primaria"], True, FONTE_TITULO)
        from openpyxl.drawing.text import RegularTextRun
        rico.p[0].r = [RegularTextRun(rPr=rico.p[0].pPr.defRPr, t=titulo)]
        grafico.title = Title(tx=Text(rich=rico), overlay=False)
    for eixo, cor in ((getattr(grafico, "x_axis", None), "texto"), (getattr(grafico, "y_axis", None), "texto_suave")):
        if eixo is not None:
            eixo.txPr = _texto_rico(9, COR[cor])
            eixo.majorGridlines = None
            eixo.delete = False
    if formato_valor and getattr(grafico, "y_axis", None) is not None:
        grafico.y_axis.numFmt = formato_valor
    if len(grafico.series) <= 1:
        grafico.legend = None
    elif grafico.legend is not None:
        grafico.legend.position = "b"
        grafico.legend.txPr = _texto_rico(9, COR["texto_suave"])
    if rotulos:
        # só o valor: sem nome da série, da categoria nem o quadradinho da legenda
        grafico.dataLabels = DataLabelList(showVal=True, showSerName=False, showCatName=False,
                                           showLegendKey=False, showPercent=False,
                                           showLeaderLines=False)
        if getattr(grafico, "tagname", "") in ("barChart", "bar3DChart"):
            grafico.dataLabels.position = "outEnd"
        if formato_valor:
            grafico.dataLabels.numFmt = formato_valor  # ver salvar(): marca sourceLinked=0
        grafico.dataLabels.txPr = _texto_rico(9, COR["texto"])
        # com rótulos, o eixo de valores é redundante
        if getattr(grafico, "y_axis", None) is not None:
            grafico.y_axis.delete = True
    # área interna ocupando ~100% da largura (o Excel reserva espaço vazio por padrão)
    from openpyxl.chart.layout import Layout, ManualLayout
    topo = 0.12 if titulo else 0.03
    grafico.plot_area.layout = Layout(manualLayout=ManualLayout(
        layoutTarget="outer", xMode="edge", yMode="edge",
        x=0.01, y=topo, w=0.98, h=0.97 - topo - (0.08 if len(grafico.series) > 1 else 0)))
    # sem moldura arredondada/cinza em volta do gráfico
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.drawing.line import LineProperties
    if getattr(grafico, "x_axis", None) is not None:
        grafico.x_axis.graphicalProperties = GraphicalProperties(
            ln=LineProperties(solidFill=COR["borda"]))
    grafico.roundedCorners = False
    grafico.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))
    return grafico


def _px_coluna(ws, col):
    w = ws.column_dimensions[get_column_letter(col)].width or 8.43
    return int(w * 7 + 5)


def posicionar_grafico(ws, grafico, linha, coluna_final=None, altura_px=340,
                       largura_max_px=1400, parte=(1, 1), espaco_px=24):
    """Coloca o gráfico abaixo da tabela com a MESMA largura dela (coluna A até
    `coluna_final`), para a aba parecer uma tela só.

    - Tabela muito larga (> largura_max_px): o gráfico fica com a largura máxima,
      centralizado embaixo da tabela.
    - parte=(i, n): divide a largura em n gráficos lado a lado (i começa em 1).
    `linha`: linha onde o gráfico começa (recomendado: linha da nota + 2).
    """
    # Âncora presa às células (TwoCellAnchor): o gráfico estica junto com as colunas,
    # então bate exatamente com a borda da tabela em qualquer Excel/zoom.
    from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
    from openpyxl.utils.units import pixels_to_EMU

    if coluna_final is None:
        coluna_final = max(c for c in range(1, ws.max_column + 1)
                           if ws.column_dimensions[get_column_letter(c)].width) if ws.max_column else 1
    total = sum(_px_coluna(ws, c) for c in range(1, coluna_final + 1))
    largura_area = min(total, largura_max_px)
    inicio_area = (total - largura_area) // 2  # centraliza se a tabela for mais larga
    i, n = parte
    largura = (largura_area - espaco_px * (n - 1)) // n
    x0 = inicio_area + (i - 1) * (largura + espaco_px)
    x1 = x0 + largura

    def marcador_x(x):
        """px desde a coluna A -> (coluna 0-based, deslocamento em px)."""
        col, acumulado = 1, 0
        while col <= coluna_final and acumulado + _px_coluna(ws, col) <= x:
            acumulado += _px_coluna(ws, col)
            col += 1
        return col - 1, x - acumulado

    c0, off0 = marcador_x(x0)
    if x1 >= total:  # termina exatamente na borda direita da última coluna
        c1, off1 = coluna_final, 0
    else:
        c1, off1 = marcador_x(x1)
    linhas = max(8, round(altura_px / 20))  # linha padrão ≈ 20 px
    grafico.anchor = TwoCellAnchor(
        _from=AnchorMarker(col=c0, colOff=pixels_to_EMU(off0), row=linha - 1, rowOff=0),
        to=AnchorMarker(col=c1, colOff=pixels_to_EMU(off1), row=linha - 1 + linhas, rowOff=0))
    ws.add_chart(grafico)
    return grafico


def salvar(wb, caminho):
    """Salva o arquivo e corrige um detalhe do openpyxl: o formato de número dos
    rótulos de gráfico sai sem sourceLinked="0" e o Excel ignora o formato pedido.
    Use sempre no lugar de wb.save() quando houver gráficos."""
    import re
    import shutil
    import tempfile
    import zipfile
    # abre sempre na primeira aba (Resumo)
    wb.active = 0
    for i, aba in enumerate(wb.worksheets):
        aba.sheet_view.tabSelected = (i == 0)
    wb.save(caminho)
    tmp = tempfile.mktemp(suffix=".xlsx")
    with zipfile.ZipFile(caminho) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            dados = zin.read(item.filename)
            if item.filename.startswith("xl/charts/chart"):
                xml = dados.decode("utf-8")
                xml = re.sub(r'<((?:c:)?numFmt) formatCode="([^"]*)"\s*/>',
                             r'<\1 formatCode="\2" sourceLinked="0"/>', xml)
                dados = xml.encode("utf-8")
            zout.writestr(item, dados)
    shutil.move(tmp, caminho)
    return caminho


if __name__ == "__main__":
    # Exemplo: gera exemplo.xlsx para conferir o estilo
    import datetime as dt
    import sys

    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Resumo"
    cab = ["Item", "Data", "Categoria", "Valor", "Meta", "Variação", "Status"]
    linhas = [
        ["Aluguel", dt.date(2026, 1, 5), "Moradia", 1800, 1800, 0.0, "Pago"],
        ["Mercado", dt.date(2026, 1, 9), "Alimentação", 950.4, 900, 0.056, "Pago"],
        ["Internet", dt.date(2026, 1, 12), "Contas", 119.9, 120, -0.001, "Pendente"],
        ["Academia", dt.date(2026, 1, 15), "Saúde", 99.9, 100, -0.001, "Atrasado"],
        ["Transporte", dt.date(2026, 1, 20), "Mobilidade", 320, 280, 0.143, "Pago"],
        ["Reembolso", dt.date(2026, 1, 25), "Outros", -150, 0, 0.0, "OK"],
    ]
    for i, nome in enumerate(cab, 1):
        ws.cell(4, i, nome)
    for r, linha in enumerate(linhas, 5):
        for c, v in enumerate(linha, 1):
            ws.cell(r, c, v)
    aplicar_estilo(
        ws, titulo="Controle de Despesas", subtitulo="Janeiro 2026 · atualizado em 16/11/2026",
        formatos={"Data": "data", "Valor": "moeda", "Meta": "moeda", "Variação": "variacao"},
        linha_total=True, colunas_total=["Valor", "Meta"], colunas_entrada=["Meta"],
        coluna_status="Status", nota="Fonte: extrato bancário. Células amarelas são editáveis.",
    )
    saida = sys.argv[1] if len(sys.argv) > 1 else "exemplo.xlsx"
    wb.save(saida)
    print(f"Gerado: {saida}")
