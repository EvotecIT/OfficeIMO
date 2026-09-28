import pathlib
import subprocess
import time
import uno
from com.sun.star.beans import PropertyValue

root = pathlib.Path(__file__).resolve().parent
pipe = 'officeimo_area_probe_20260928'
profile = root / 'lo-profile'
process = subprocess.Popen(['/usr/bin/libreoffice', '-env:UserInstallation=' + profile.as_uri(),
    '--headless', '--nologo', '--nodefault', '--norestore',
    '--accept=pipe,name=' + pipe + ';urp;StarOffice.ComponentContext'])

def prop(name, value):
    item = PropertyValue(); item.Name = name; item.Value = value; return item

document = None
desktop = None
try:
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
    context = None
    for attempt in range(400):
        try:
            context = resolver.resolve('uno:pipe,name=' + pipe + ';urp;StarOffice.ComponentContext')
            break
        except Exception:
            if process.poll() is not None: raise RuntimeError('LibreOffice exited')
            time.sleep(0.1)
    if context is None: raise RuntimeError('LibreOffice did not start')
    desktop = context.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', context)
    document = desktop.loadComponentFromURL('private:factory/swriter', '_blank', 0, (prop('Hidden', True),))
    text = document.getText()
    text.setString('Area point paint probe')
    cursor = text.createTextCursor(); cursor.gotoEnd(False)
    text.insertControlCharacter(cursor, uno.getConstantByName('com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK'), False)
    embedded = document.createInstance('com.sun.star.text.TextEmbeddedObject')
    embedded.CLSID = '12DCAE26-281F-416F-A234-C3086127382E'
    embedded.Width = 14000; embedded.Height = 8000
    text.insertTextContent(cursor, embedded, False)
    chart = embedded.Model
    chart.setDiagram(chart.createInstance('com.sun.star.chart.AreaDiagram'))
    chart.getData().setData(((3.0,), (6.0,), (4.0,), (7.0,)))
    chart.getData().setRowDescriptions(('A', 'B', 'C', 'D'))
    chart.getData().setColumnDescriptions(('Probe',))
    diagram = chart.getFirstDiagram()
    series = diagram.getCoordinateSystems()[0].getChartTypes()[0].getDataSeries()[0]
    for index, colour in enumerate((0x228844, 0xD97706, 0x445566, 0xC04060)):
        series.getDataPointByIndex(index).FillColor = colour
    chart.HasLegend = True
    chart.HasMainTitle = True; chart.Title.String = 'Area point paint'
    document.storeToURL((root/'area-point.odt').as_uri(), (prop('FilterName', 'writer8'), prop('Overwrite', True)))
    document.storeToURL((root/'area-point.docx').as_uri(), (prop('FilterName', 'Office Open XML Text'), prop('Overwrite', True)))
    document.storeToURL((root/'area-point-source-reference.pdf').as_uri(), (prop('FilterName', 'writer_pdf_Export'), prop('Overwrite', True)))
finally:
    if document is not None: document.close(True)
    if desktop is not None: desktop.terminate()
    try: process.wait(timeout=10)
    except subprocess.TimeoutExpired: process.terminate(); process.wait(timeout=5)
