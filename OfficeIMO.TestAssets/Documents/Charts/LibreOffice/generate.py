import pathlib, subprocess, time, uno
from com.sun.star.beans import PropertyValue
root = pathlib.Path(__file__).resolve().parent
profile = root / 'lo-profile'
pipe = 'officeimo_chart_producer_20260928'
process = subprocess.Popen(['/usr/bin/libreoffice', '-env:UserInstallation=' + profile.as_uri(), '--headless', '--nologo', '--nodefault', '--norestore', '--accept=pipe,name=' + pipe + ';urp;StarOffice.ComponentContext'], stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)
def prop(name, value):
    item = PropertyValue(); item.Name = name; item.Value = value; return item
context = None
document = None
desktop = None
try:
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
    for attempt in range(100):
        try:
            context = resolver.resolve('uno:pipe,name=' + pipe + ';urp;StarOffice.ComponentContext'); break
        except Exception:
            if process.poll() is not None: raise RuntimeError('Isolated LibreOffice exited during startup')
            time.sleep(0.1)
    if context is None: raise RuntimeError('Isolated LibreOffice did not start within ten seconds')
    desktop = context.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', context)
    document = desktop.loadComponentFromURL('private:factory/swriter', '_blank', 0, (prop('Hidden', True),))
    text = document.getText()
    text.setString('Independent status chart fixture')
    cursor = text.createTextCursor(); cursor.gotoEnd(False)
    text.insertControlCharacter(cursor, uno.getConstantByName('com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK'), False)
    embedded = document.createInstance('com.sun.star.text.TextEmbeddedObject')
    embedded.CLSID = '12DCAE26-281F-416F-A234-C3086127382E'
    embedded.Width = 12000; embedded.Height = 7500
    text.insertTextContent(cursor, embedded, False)
    chart = embedded.Model
    diagram = chart.createInstance('com.sun.star.chart.PieDiagram')
    chart.setDiagram(diagram)
    chart.getData().setData(((3.0,), (4.0,), (5.0,)))
    chart.getData().setRowDescriptions(('Pass', 'Fail', 'Unknown'))
    chart.getData().setColumnDescriptions(('Status',))
    diagram.DataRowSource = uno.Enum('com.sun.star.chart.ChartDataRowSource', 'COLUMNS')
    for index, colour in enumerate((0x228844, 0xD97706, 0x445566)):
        diagram.getDataPointProperties(index, 0).FillColor = colour
    unknown = diagram.getDataPointProperties(2, 0)
    unknown.FillStyle = uno.Enum('com.sun.star.drawing.FillStyle', 'NONE')
    unknown.LineStyle = uno.Enum('com.sun.star.drawing.LineStyle', 'SOLID')
    unknown.LineColor = 0x445566; unknown.LineWidth = 70
    print('Point paint properties:', [p.Name for p in unknown.getPropertySetInfo().getProperties() if 'Border' in p.Name or 'Line' in p.Name])
    series = chart.getFirstDiagram().getCoordinateSystems()[0].getChartTypes()[0].getDataSeries()[0]
    point = series.getDataPointByIndex(2)
    point.BorderStyle = uno.Enum('com.sun.star.drawing.LineStyle', 'SOLID')
    point.BorderColor = 0x445566; point.BorderWidth = 70
    hatch = uno.createUnoStruct('com.sun.star.drawing.Hatch')
    hatch.Style = uno.Enum('com.sun.star.drawing.HatchStyle', 'SINGLE')
    hatch.Color = 0xD97706; hatch.Distance = 150; hatch.Angle = 450
    hatch_table = chart.createInstance('com.sun.star.drawing.HatchTable')
    hatch_table.insertByName('StatusFail', hatch)
    failure = series.getDataPointByIndex(1)
    failure.FillStyle = uno.Enum('com.sun.star.drawing.FillStyle', 'HATCH')
    failure.FillHatchName = 'StatusFail'; failure.FillBackground = True; failure.FillColor = 0xFFF0DD
    chart.HasLegend = True
    chart.HasMainTitle = True; chart.Title.String = 'Status'
    document.storeToURL((root/'status-pie.docx').as_uri(), (prop('FilterName', 'Office Open XML Text'), prop('Overwrite', True)))
    document.storeToURL((root/'status-pie.odt').as_uri(), (prop('FilterName', 'writer8'), prop('Overwrite', True)))
    document.storeToURL((root/'status-pie-source-reference.pdf').as_uri(), (prop('FilterName', 'writer_pdf_Export'), prop('Overwrite', True)))
    print('Generated independent Writer pie DOCX/ODT/reference PDF')
finally:
    if document is not None: document.close(True)
    if desktop is not None: desktop.terminate()
    try: process.wait(timeout=10)
    except subprocess.TimeoutExpired: process.terminate(); process.wait(timeout=5)
