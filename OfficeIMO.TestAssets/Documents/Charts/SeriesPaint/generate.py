"""Reproduce the independent series fill/outline cohort using the existing Office chart package writer."""
import importlib.util
from pathlib import Path

HERE = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location('automatic_colors', HERE.parent / 'AutomaticColors' / 'generate.py')
g = importlib.util.module_from_spec(spec)
spec.loader.exec_module(g)
g.SLIDES = []
for kind in ('area', 'bubble', 'column', 'line'):
    for fill in (True, False):
        for outline in (True, False):
            title = f'{kind}: fill {int(fill)}, outline {int(outline)}'
            paint = '<c:spPr>' + ('<a:solidFill><a:srgbClr val="156082"/></a:solidFill>' if fill else '<a:noFill/>')
            paint += '<a:ln w="38100">' + ('<a:solidFill><a:srgbClr val="E97132"/></a:solidFill>' if outline else '<a:noFill/>') + '</a:ln></c:spPr>'
            body = g.series(0, g.VALUES[0], g.CATEGORIES, scatter=kind == 'bubble', extra=paint)
            if kind == 'bubble':
                body = body.replace('</c:ser>', '<c:bubbleSize><c:numLit>' + g.num_cache([4, 8, 6, 10]) + '</c:numLit></c:bubbleSize></c:ser>')
                group = '<c:bubbleChart><c:varyColors val="0"/>' + body + '<c:bubbleScale val="100"/><c:sizeRepresents val="area"/><c:axId val="5001"/><c:axId val="5002"/></c:bubbleChart>'
            elif kind == 'column':
                group = '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>' + body + '<c:gapWidth val="150"/><c:axId val="5001"/><c:axId val="5002"/></c:barChart>'
            else:
                if kind == 'line':
                    body = body.replace('</c:spPr>', '</c:spPr><c:marker><c:symbol val="none"/></c:marker>').replace('</c:ser>', '<c:smooth val="0"/></c:ser>')
                group = f'<c:{kind}Chart><c:grouping val="standard"/><c:varyColors val="0"/>' + body + f'<c:axId val="5001"/><c:axId val="5002"/></c:{kind}Chart>'
            g.SLIDES.append((title, g.chart_xml(title, group, g.axes(scatter=kind == 'bubble').replace('<c:majorGridlines/>', '<c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="E2E8F0"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>'), legend=False), None, None))
for extension, writer in [('xlsx', g.build_workbook), ('docx', g.build_document), ('pptx', g.build_presentation)]:
    writer(HERE / ('series-paint.' + extension))
