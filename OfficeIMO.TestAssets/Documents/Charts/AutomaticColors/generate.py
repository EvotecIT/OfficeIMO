"""Generates the automatic-colour chart fixtures.

Every chart series is left without c:spPr, so Office chooses its colours itself.
Some charts carry modern chart-style and colour-style parts and others only the
classic c:style number; the Office exports show which of them Office uses. Run
from the repository root:

    python3 OfficeIMO.TestAssets/Documents/Charts/AutomaticColors/generate.py

The script uses only the Python standard library and writes fixed timestamps,
so running it again reproduces the same package bytes.
"""

import os
import re
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
STAMP = (2026, 10, 16, 0, 0, 0)

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_CS = "http://schemas.microsoft.com/office/drawing/2012/chartStyle"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
CT_CHART_STYLE = "application/vnd.ms-office.chartstyle+xml"
CT_CHART_COLORS = "application/vnd.ms-office.chartcolorstyle+xml"
REL_CHART_STYLE = "http://schemas.microsoft.com/office/2011/relationships/chartStyle"
REL_CHART_COLORS = "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle"
XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

THEME = XML_HEADER + (
    '<a:theme xmlns:a="' + NS_A + '" name="Office Theme"><a:themeElements>'
    '<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
    '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="0E2841"/></a:dk2>'
    '<a:lt2><a:srgbClr val="E8E8E8"/></a:lt2><a:accent1><a:srgbClr val="156082"/></a:accent1>'
    '<a:accent2><a:srgbClr val="E97132"/></a:accent2><a:accent3><a:srgbClr val="196B24"/></a:accent3>'
    '<a:accent4><a:srgbClr val="0F9ED5"/></a:accent4><a:accent5><a:srgbClr val="A02B93"/></a:accent5>'
    '<a:accent6><a:srgbClr val="4EA72E"/></a:accent6><a:hlink><a:srgbClr val="467886"/></a:hlink>'
    '<a:folHlink><a:srgbClr val="96607D"/></a:folHlink></a:clrScheme>'
    '<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
    '<a:minorFont><a:latin typeface="Aptos"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
    '<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
    '<a:lnStyleLst><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
    '<a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
    '<a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
    '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle>'
    '<a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
    '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>'
    '</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>')

# Office's "Colorful palette 1" colour style, as PowerPoint writes it.
COLORFUL1 = XML_HEADER + (
    '<cs:colorStyle xmlns:cs="' + NS_CS + '" xmlns:a="' + NS_A + '" meth="cycle" id="10">'
    + ''.join('<a:schemeClr val="accent%d"/>' % i for i in range(1, 7)) +
    '<cs:variation/><cs:variation><a:lumMod val="60000"/></cs:variation>'
    '<cs:variation><a:lumMod val="80000"/><a:lumOff val="20000"/></cs:variation>'
    '<cs:variation><a:lumMod val="80000"/></cs:variation>'
    '<cs:variation><a:lumMod val="60000"/><a:lumOff val="40000"/></cs:variation>'
    '<cs:variation><a:lumMod val="50000"/></cs:variation>'
    '<cs:variation><a:lumMod val="70000"/><a:lumOff val="30000"/></cs:variation>'
    '<cs:variation><a:lumMod val="70000"/></cs:variation>'
    '<cs:variation><a:lumMod val="50000"/><a:lumOff val="50000"/></cs:variation></cs:colorStyle>')

# Office's "Monochromatic palette 1" colour style.
MONOCHROMATIC1 = XML_HEADER + (
    '<cs:colorStyle xmlns:cs="' + NS_CS + '" xmlns:a="' + NS_A + '" meth="withinLinear" id="14">'
    '<a:schemeClr val="accent1"/></cs:colorStyle>')


# PowerPoint's chart style 201, as PowerPoint writes it for new charts. Its
# data-point entries apply the automatic colour to fills, lines and markers.
with open(os.path.join(HERE, 'chart-style-201.xml'), encoding='utf-8') as style_file:
    STYLE_201 = style_file.read()
if not STYLE_201.startswith('<?xml'):
    STYLE_201 = XML_HEADER + STYLE_201

CATEGORIES = ["North", "East", "South", "West"]
VALUES = [
    [4.3, 2.5, 3.5, 4.5], [2.4, 4.4, 1.8, 2.8], [2.0, 2.0, 3.0, 5.0], [3.1, 1.6, 2.6, 3.9],
    [1.2, 3.3, 4.1, 2.2], [3.8, 2.9, 1.4, 3.3], [2.7, 3.6, 2.2, 1.7], [1.9, 1.1, 3.8, 2.6],
    [3.4, 2.1, 2.9, 1.5], [2.2, 4.0, 1.6, 3.1], [1.4, 2.6, 3.3, 4.2], [4.1, 1.8, 2.5, 2.9]]


def str_cache(values):
    return '<c:ptCount val="%d"/>' % len(values) + ''.join(
        '<c:pt idx="%d"><c:v>%s</c:v></c:pt>' % (i, v) for i, v in enumerate(values))


def num_cache(values):
    return '<c:formatCode>General</c:formatCode><c:ptCount val="%d"/>' % len(values) + ''.join(
        '<c:pt idx="%d"><c:v>%s</c:v></c:pt>' % (i, repr(float(v))) for i, v in enumerate(values))


def column(index):
    return chr(ord('B') + index)


def series(index, values, categories, scatter=False, extra=''):
    tx = ('<c:tx><c:strRef><c:f>Sheet1!$%s$1</c:f><c:strCache>%s</c:strCache></c:strRef></c:tx>'
          % (column(index), str_cache(["Series %d" % (index + 1)])))
    count = len(values) + 1
    if scatter:
        data = ('<c:xVal><c:numRef><c:f>Sheet1!$A$2:$A$%d</c:f><c:numCache>%s</c:numCache></c:numRef></c:xVal>'
                '<c:yVal><c:numRef><c:f>Sheet1!$%s$2:$%s$%d</c:f><c:numCache>%s</c:numCache></c:numRef></c:yVal>'
                % (count, num_cache(range(1, len(values) + 1)), column(index), column(index), count, num_cache(values)))
    else:
        data = ('<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$%d</c:f><c:strCache>%s</c:strCache></c:strRef></c:cat>'
                '<c:val><c:numRef><c:f>Sheet1!$%s$2:$%s$%d</c:f><c:numCache>%s</c:numCache></c:numRef></c:val>'
                % (count, str_cache(categories), column(index), column(index), count, num_cache(values)))
    return '<c:ser><c:idx val="%d"/><c:order val="%d"/>%s%s%s</c:ser>' % (index, index, tx, extra, data)


def axes(category_position='b', value_position='l', scatter=False):
    category = ('<c:valAx><c:axId val="5001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
                '<c:axPos val="%s"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/>'
                '<c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="5002"/>'
                '<c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>' % category_position) if scatter else (
                '<c:catAx><c:axId val="5001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
                '<c:axPos val="%s"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/>'
                '<c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="5002"/>'
                '<c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/>'
                '<c:noMultiLvlLbl val="0"/></c:catAx>' % category_position)
    value = ('<c:valAx><c:axId val="5002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
             '<c:axPos val="%s"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/>'
             '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
             '<c:crossAx val="5001"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>'
             % value_position)
    return category + value


def chart_xml(title, group, axes_xml, style=2, legend=True):
    legend_xml = '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>' if legend else ''
    return XML_HEADER + (
        '<c:chartSpace xmlns:c="' + NS_C + '" xmlns:a="' + NS_A + '" xmlns:r="' + NS_R + '">'
        '<c:date1904 val="0"/><c:roundedCorners val="0"/><c:style val="%d"/><c:chart>' % style +
        '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>%s</a:t></a:r></a:p>'
        '</c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>' % title +
        '<c:plotArea><c:layout/>' + group + axes_xml + '</c:plotArea>' + legend_xml +
        '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>')


def bar_group(direction, count, grouping='clustered'):
    body = ''.join(series(i, VALUES[i], CATEGORIES, extra='<c:invertIfNegative val="0"/>') for i in range(count))
    overlap = '' if grouping == 'clustered' else '<c:overlap val="100"/>'
    return ('<c:barChart><c:barDir val="%s"/><c:grouping val="%s"/><c:varyColors val="0"/>%s'
            '<c:gapWidth val="150"/>%s<c:axId val="5001"/><c:axId val="5002"/></c:barChart>'
            % (direction, grouping, body, overlap))


def line_group(count):
    body = ''.join(series(i, VALUES[i], CATEGORIES) for i in range(count))
    body = body.replace('</c:ser>', '<c:smooth val="0"/></c:ser>')
    return ('<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>%s<c:marker val="1"/>'
            '<c:axId val="5001"/><c:axId val="5002"/></c:lineChart>' % body)


def area_group(count):
    body = ''.join(series(i, VALUES[i], CATEGORIES) for i in range(count))
    return ('<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>%s'
            '<c:axId val="5001"/><c:axId val="5002"/></c:areaChart>' % body)


def scatter_group(count):
    body = ''.join(series(i, VALUES[i], CATEGORIES, scatter=True) for i in range(count))
    body = body.replace('</c:ser>', '<c:smooth val="0"/></c:ser>')
    return ('<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>%s'
            '<c:axId val="5001"/><c:axId val="5002"/></c:scatterChart>' % body)


def pie_group(points):
    categories = ["Slice %d" % (i + 1) for i in range(points)]
    values = [8 - (i % 5) for i in range(points)]
    return ('<c:pieChart><c:varyColors val="1"/>%s<c:firstSliceAng val="0"/></c:pieChart>'
            % series(0, values, categories))


def doughnut_group(rings):
    categories = ["Slice %d" % (i + 1) for i in range(4)]
    body = ''.join(series(i, VALUES[i], categories) for i in range(rings))
    return ('<c:doughnutChart><c:varyColors val="1"/>%s<c:firstSliceAng val="0"/><c:holeSize val="50"/></c:doughnutChart>'
            % body)


def stock_group():
    names = ["High", "Low", "Close"]
    high, low, close = [44, 25, 38, 47], [24, 18, 21, 31], [38, 22, 33, 41]
    body = ''
    for i, values in enumerate([high, low, close]):
        marker = '<c:marker><c:symbol val="%s"/>%s</c:marker>' % (
            'dot' if i == 2 else 'none', '<c:size val="5"/>' if i == 2 else '')
        item = series(i, values, CATEGORIES, extra='<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr>' + marker)
        body += item.replace('Series %d' % (i + 1), names[i]).replace('</c:ser>', '<c:smooth val="0"/></c:ser>')
    return ('<c:stockChart>%s<c:hiLowLines/><c:axId val="5001"/><c:axId val="5002"/></c:stockChart>' % body)


# (slide title, chart xml, optional modern style part, optional colour style part)
SLIDES = [
    ("Classic style 2 clustered column", chart_xml("Classic column", bar_group('col', 3), axes()), None, None),
    ("Classic style 2 with eight series", chart_xml("Classic eight series", bar_group('col', 8), axes()), None, None),
    ("Colourful palette 1 with eight series", chart_xml("Colourful eight series", bar_group('col', 8), axes()),
     STYLE_201, COLORFUL1),
    ("Clustered bar", chart_xml("Clustered bar", bar_group('bar', 3), axes('l', 'b')), STYLE_201, COLORFUL1),
    ("Line with markers", chart_xml("Line", line_group(3), axes()), STYLE_201, COLORFUL1),
    ("Classic line", chart_xml("Classic line", line_group(3), axes()), None, None),
    ("Area", chart_xml("Area", area_group(2), axes()), STYLE_201, COLORFUL1),
    ("Scatter", chart_xml("Scatter", scatter_group(3), axes(scatter=True)), STYLE_201, COLORFUL1),
    ("Classic pie", chart_xml("Classic pie", pie_group(4), ''), None, None),
    ("Pie with eight slices", chart_xml("Pie eight slices", pie_group(8), ''), STYLE_201, COLORFUL1),
    ("Stock high-low-close", chart_xml("Stock", stock_group(), axes()), None, None),
    ("Monochromatic palette 1", chart_xml("Monochromatic", bar_group('col', 4), axes()), STYLE_201, MONOCHROMATIC1),
    ("Classic style 4 monochrome", chart_xml("Classic style 4", bar_group('col', 4), axes(), style=4), None, None),
    ("Classic style 1 greyscale", chart_xml("Classic style 1", bar_group('col', 6), axes(), style=1), None, None),
    ("Pie with monochromatic palette", chart_xml("Pie monochromatic", pie_group(4), ''), STYLE_201, MONOCHROMATIC1),
    ("Doughnut with two rings", chart_xml("Two-ring doughnut", doughnut_group(2), ''), None, None),
    ("Classic style 2 with six series", chart_xml("Classic six series", bar_group('col', 6), axes()), None, None),
    ("Classic style 2 with twelve series", chart_xml("Classic twelve series", bar_group('col', 12), axes()), None, None),
    ("Classic style 1 with four series", chart_xml("Classic style 1 four", bar_group('col', 4), axes(), style=1), None, None),
]


def text_box(title):
    return ('<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="457200" y="228600"/><a:ext cx="8229600" cy="457200"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            '<p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="2000" b="1"/>'
            '<a:t>%s</a:t></a:r></a:p></p:txBody></p:sp>' % title)


def chart_frame():
    return ('<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="Chart"/><p:cNvGraphicFramePr/><p:nvPr/>'
            '</p:nvGraphicFramePr><p:xfrm><a:off x="914400" y="800100"/><a:ext cx="7315200" cy="4114800"/></p:xfrm>'
            '<a:graphic><a:graphicData uri="' + NS_C + '"><c:chart xmlns:c="' + NS_C + '" r:id="rId2"/>'
            '</a:graphicData></a:graphic></p:graphicFrame>')


def relationships(items):
    return XML_HEADER + '<Relationships xmlns="' + REL_PKG + '">' + ''.join(
        '<Relationship Id="%s" Type="%s" Target="%s"/>' % item for item in items) + '</Relationships>'


def write(zip_file, name, text):
    info = zipfile.ZipInfo(name, STAMP)
    info.compress_type = zipfile.ZIP_DEFLATED
    zip_file.writestr(info, text.encode('utf-8'))


def add_chart_style_parts(package, folder, number, style, colors, overrides):
    chart_rels = []
    if style:
        write(package, '%s/style%d.xml' % (folder, number), style)
        overrides.append(('/%s/style%d.xml' % (folder, number), CT_CHART_STYLE))
        chart_rels.append(('rId1', REL_CHART_STYLE, 'style%d.xml' % number))
    if colors:
        write(package, '%s/colors%d.xml' % (folder, number), colors)
        overrides.append(('/%s/colors%d.xml' % (folder, number), CT_CHART_COLORS))
        chart_rels.append(('rId2', REL_CHART_COLORS, 'colors%d.xml' % number))
    if chart_rels:
        write(package, '%s/_rels/chart%d.xml.rels' % (folder, number), relationships(chart_rels))


def content_types(overrides):
    return XML_HEADER + (
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>' + ''.join(
            '<Override PartName="%s" ContentType="%s"/>' % item for item in overrides) + '</Types>')


def build_presentation(path):
    overrides = [
        ('/ppt/presentation.xml', 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'),
        ('/ppt/slideMasters/slideMaster1.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'),
        ('/ppt/slideLayouts/slideLayout1.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'),
        ('/ppt/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml')]
    with zipfile.ZipFile(path, 'w') as package:
        slide_ids, presentation_rels = '', [('rId1', REL + '/slideMaster', 'slideMasters/slideMaster1.xml'),
                                            ('rId2', REL + '/theme', 'theme/theme1.xml')]
        for number, (title, chart, style, colors) in enumerate(SLIDES, 1):
            overrides.append(('/ppt/slides/slide%d.xml' % number,
                              'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'))
            overrides.append(('/ppt/charts/chart%d.xml' % number, CT_CHART))
            slide_ids += '<p:sldId id="%d" r:id="rId%d"/>' % (255 + number, number + 2)
            presentation_rels.append(('rId%d' % (number + 2), REL + '/slide', 'slides/slide%d.xml' % number))
            write(package, 'ppt/slides/slide%d.xml' % number, XML_HEADER + (
                '<p:sld xmlns:a="' + NS_A + '" xmlns:r="' + NS_R + '" xmlns:p="' + NS_P + '"><p:cSld><p:spTree>'
                '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>'
                + text_box(title) + chart_frame() + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'))
            write(package, 'ppt/slides/_rels/slide%d.xml.rels' % number, relationships([
                ('rId1', REL + '/slideLayout', '../slideLayouts/slideLayout1.xml'),
                ('rId2', REL + '/chart', '../charts/chart%d.xml' % number)]))
            write(package, 'ppt/charts/chart%d.xml' % number, chart)
            add_chart_style_parts(package, 'ppt/charts', number, style, colors, overrides)
        write(package, '[Content_Types].xml', content_types(overrides))
        write(package, '_rels/.rels', relationships([
            ('rId1', REL + '/officeDocument', 'ppt/presentation.xml')]))
        write(package, 'ppt/presentation.xml', XML_HEADER + (
            '<p:presentation xmlns:a="' + NS_A + '" xmlns:r="' + NS_R + '" xmlns:p="' + NS_P + '">'
            '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
            '<p:sldIdLst>' + slide_ids + '</p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/>'
            '<p:notesSz cx="6858000" cy="9144000"/></p:presentation>'))
        write(package, 'ppt/_rels/presentation.xml.rels', relationships(presentation_rels))
        write(package, 'ppt/slideMasters/slideMaster1.xml', XML_HEADER + (
            '<p:sldMaster xmlns:a="' + NS_A + '" xmlns:r="' + NS_R + '" xmlns:p="' + NS_P + '"><p:cSld>'
            '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr>'
            '<p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>'
            '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" '
            'accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
            '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>'
            '<p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>'))
        write(package, 'ppt/slideMasters/_rels/slideMaster1.xml.rels', relationships([
            ('rId1', REL + '/slideLayout', '../slideLayouts/slideLayout1.xml'),
            ('rId2', REL + '/theme', '../theme/theme1.xml')]))
        write(package, 'ppt/slideLayouts/slideLayout1.xml', XML_HEADER + (
            '<p:sldLayout xmlns:a="' + NS_A + '" xmlns:r="' + NS_R + '" xmlns:p="' + NS_P + '" type="blank">'
            '<p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/>'
            '</p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>'))
        write(package, 'ppt/slideLayouts/_rels/slideLayout1.xml.rels', relationships([
            ('rId1', REL + '/slideMaster', '../slideMasters/slideMaster1.xml')]))
        write(package, 'ppt/theme/theme1.xml', THEME)


CELL_REFERENCE = re.compile(
    r"<c:f>Sheet1!\$([A-Z])\$(\d+)(?::\$([A-Z])\$(\d+))?</c:f><c:(str|num)Cache>(.*?)</c:(?:str|num)Cache>")
POINT = re.compile(r'<c:pt idx="(\d+)"><c:v>(.*?)</c:v></c:pt>')


def worksheet_cells(chart):
    """Rebuilds the worksheet cells that each cached chart reference points to."""
    cells = {}
    for match in CELL_REFERENCE.finditer(chart):
        first_column, first_row, last_column, _, kind, cache = match.groups()
        vertical = last_column is None or last_column == first_column
        for index, value in POINT.findall(cache):
            offset = int(index)
            column_name = first_column if vertical else chr(ord(first_column) + offset)
            row = int(first_row) + (offset if vertical else 0)
            cells[(row, column_name)] = (kind, value)
    return cells


def sheet_xml(cells, strings):
    rows = {}
    for (row, column_name), value in cells.items():
        rows.setdefault(row, []).append((column_name, value))
    body = ''
    for row in sorted(rows):
        body += '<row r="%d">' % row
        for column_name, (kind, value) in sorted(rows[row]):
            if kind == 'str':
                if value not in strings:
                    strings.append(value)
                body += '<c r="%s%d" t="s"><v>%d</v></c>' % (column_name, row, strings.index(value))
            else:
                body += '<c r="%s%d"><v>%s</v></c>' % (column_name, row, value)
        body += '</row>'
    return XML_HEADER + (
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="' + NS_R + '">'
        '<sheetData>' + body + '</sheetData><pageMargins left="0.5" right="0.5" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>'
        '<pageSetup orientation="landscape"/><drawing r:id="rId1"/></worksheet>')


def build_workbook(path):
    """One worksheet per chart; each sheet holds the chart's own cached data."""
    overrides = [
        ('/xl/workbook.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'),
        ('/xl/styles.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'),
        ('/xl/sharedStrings.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'),
        ('/xl/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml')]
    strings, sheets = [], ''
    workbook_rels = [('rId1', REL + '/styles', 'styles.xml'), ('rId2', REL + '/theme', 'theme/theme1.xml'),
                     ('rId3', REL + '/sharedStrings', 'sharedStrings.xml')]
    with zipfile.ZipFile(path, 'w') as package:
        for number, (title, chart, style, colors) in enumerate(SLIDES, 1):
            sheet_name = 'Chart%d' % number
            chart = chart.replace('Sheet1!', sheet_name + '!')
            overrides += [('/xl/worksheets/sheet%d.xml' % number,
                           'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'),
                          ('/xl/drawings/drawing%d.xml' % number, 'application/vnd.openxmlformats-officedocument.drawing+xml'),
                          ('/xl/charts/chart%d.xml' % number, CT_CHART)]
            sheets += '<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (sheet_name, number, number + 3)
            workbook_rels.append(('rId%d' % (number + 3), REL + '/worksheet', 'worksheets/sheet%d.xml' % number))
            write(package, 'xl/worksheets/sheet%d.xml' % number,
                  sheet_xml(worksheet_cells(chart.replace(sheet_name + '!', 'Sheet1!')), strings))
            write(package, 'xl/worksheets/_rels/sheet%d.xml.rels' % number, relationships([
                ('rId1', REL + '/drawing', '../drawings/drawing%d.xml' % number)]))
            write(package, 'xl/drawings/drawing%d.xml' % number, XML_HEADER + (
                '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
                'xmlns:a="' + NS_A + '" xmlns:r="' + NS_R + '"><xdr:twoCellAnchor>'
                '<xdr:from><xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
                '<xdr:to><xdr:col>20</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>22</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
                '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="%s"/><xdr:cNvGraphicFramePr/>'
                '</xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
                '<a:graphic><a:graphicData uri="' + NS_C + '"><c:chart xmlns:c="' + NS_C + '" r:id="rId1"/>'
                '</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>') % title)
            write(package, 'xl/drawings/_rels/drawing%d.xml.rels' % number, relationships([
                ('rId1', REL + '/chart', '../charts/chart%d.xml' % number)]))
            write(package, 'xl/charts/chart%d.xml' % number, chart)
            add_chart_style_parts(package, 'xl/charts', number, style, colors, overrides)
        write(package, '[Content_Types].xml', content_types(overrides))
        write(package, '_rels/.rels', relationships([('rId1', REL + '/officeDocument', 'xl/workbook.xml')]))
        write(package, 'xl/workbook.xml', XML_HEADER + (
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="' + NS_R + '">'
            '<sheets>' + sheets + '</sheets></workbook>'))
        write(package, 'xl/_rels/workbook.xml.rels', relationships(workbook_rels))
        write(package, 'xl/styles.xml', XML_HEADER + (
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<fonts count="1"><font><sz val="11"/><name val="Aptos Narrow"/><family val="2"/><scheme val="minor"/></font></fonts>'
            '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
            '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'))
        write(package, 'xl/sharedStrings.xml', XML_HEADER + (
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="%d" uniqueCount="%d">'
            % (len(strings), len(strings)) + ''.join('<si><t>%s</t></si>' % value for value in strings) + '</sst>'))
        write(package, 'xl/theme/theme1.xml', THEME)


def build_document(path):
    """One chart per page, each inline in its own paragraph."""
    overrides = [
        ('/word/document.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'),
        ('/word/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml')]
    document_rels = [('rId1', REL + '/theme', 'theme/theme1.xml')]
    body = ''
    with zipfile.ZipFile(path, 'w') as package:
        for number, (title, chart, style, colors) in enumerate(SLIDES, 1):
            overrides.append(('/word/charts/chart%d.xml' % number, CT_CHART))
            document_rels.append(('rId%d' % (number + 1), REL + '/chart', 'charts/chart%d.xml' % number))
            page_break = '' if number == 1 else '<w:pageBreakBefore/>'
            body += ('<w:p><w:pPr>%s</w:pPr><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>%s</w:t></w:r></w:p>'
                     '<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
                     '<wp:extent cx="5486400" cy="3200400"/><wp:docPr id="%d" name="Chart %d"/><wp:cNvGraphicFramePr/>'
                     '<a:graphic><a:graphicData uri="' + NS_C + '"><c:chart xmlns:c="' + NS_C + '" r:id="rId%d"/>'
                     '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>') % (
                page_break, title, number, number, number + 1)
            write(package, 'word/charts/chart%d.xml' % number, chart)
            add_chart_style_parts(package, 'word/charts', number, style, colors, overrides)
        write(package, '[Content_Types].xml', content_types(overrides))
        write(package, '_rels/.rels', relationships([('rId1', REL + '/officeDocument', 'word/document.xml')]))
        write(package, 'word/document.xml', XML_HEADER + (
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="' + NS_R + '" '
            'xmlns:a="' + NS_A + '" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
            '<w:body>' + body + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
            '</w:sectPr></w:body></w:document>'))
        write(package, 'word/_rels/document.xml.rels', relationships(document_rels))
        write(package, 'word/theme/theme1.xml', THEME)


def line_without_marker_flag(count, flag):
    body = ''.join(series(i, VALUES[i], CATEGORIES) for i in range(count))
    body = body.replace('</c:ser>', '<c:smooth val="0"/></c:ser>')
    marker = '' if flag is None else '<c:marker val="%d"/>' % flag
    return ('<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>%s%s'
            '<c:axId val="5001"/><c:axId val="5002"/></c:lineChart>' % (body, marker))


# Line series without a stored marker, under a group with no c:marker flag and with it off.
LINE_MARKER_SLIDES = [
    ("Line without the group marker flag", chart_xml("No flag", line_without_marker_flag(3, None), axes()), None, None),
    ("Line with markers switched off", chart_xml("Flag off", line_without_marker_flag(3, 0), axes()), None, None),
]


if __name__ == '__main__':
    build_presentation(os.path.join(HERE, 'automatic-colors.pptx'))
    automatic_slides, SLIDES = SLIDES, LINE_MARKER_SLIDES
    build_presentation(os.path.join(HERE, 'line-markers.pptx'))
    SLIDES = automatic_slides
    build_workbook(os.path.join(HERE, 'automatic-colors.xlsx'))
    build_document(os.path.join(HERE, 'automatic-colors.docx'))
