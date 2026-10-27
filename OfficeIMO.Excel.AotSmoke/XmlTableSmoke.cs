using System.Xml.Linq;
using OfficeIMO.Excel;

internal static class XmlTableSmoke {
    internal static void Run() {
        var options = new ExcelXmlTableOptions();
        options.Columns.Add(new ExcelXmlTableColumn { Name = "Code", XmlName = "code" });
        options.Columns.Add(new ExcelXmlTableColumn { Name = "Total", XmlName = "total", ValueType = ExcelXmlTableValueType.Number });
        options.Columns.Add(new ExcelXmlTableColumn { Name = "Paid", XmlName = "paid", ValueType = ExcelXmlTableValueType.Boolean });
        options.Columns.Add(new ExcelXmlTableColumn { Name = "Date", XmlName = "date", ValueType = ExcelXmlTableValueType.DateTime });
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("XML");
        sheet.FromXml("<Table><Row><code>00127</code><total>12.5</total><paid>true</paid><date>2026-10-19T12:30:00</date></Row></Table>", options: options);
        if (sheet.CellAt(1, 1).GetValue<string>() != "00127" || sheet.CellAt(1, 2).GetValue<double>() != 12.5 || !sheet.CellAt(1, 3).GetValue<bool>())
            throw new InvalidOperationException("XML table scalar types did not survive NativeAOT import.");
        XElement row = XElement.Parse(sheet.ToXml("A1:D1", options: options,
            readOptions: new ExcelReaderOptions { TreatDatesUsingNumberFormat = true })).Element("Row")!;
        if (row.Element("date")?.Value != "2026-10-19T12:30:00")
            throw new InvalidOperationException("XML table date did not survive NativeAOT export.");
    }
}
