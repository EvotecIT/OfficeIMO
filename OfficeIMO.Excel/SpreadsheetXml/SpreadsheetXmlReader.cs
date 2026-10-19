using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Core.Internal;
using DocumentFormat.OpenXml.Spreadsheet;
using static OfficeIMO.Excel.SpreadsheetXml.SpreadsheetXmlContext;

namespace OfficeIMO.Excel.SpreadsheetXml;

internal static class SpreadsheetXmlReader {
    internal static OfficeConversionResult<ExcelDocument, ExcelSpreadsheetXmlReport> Read(Stream input, ExcelSpreadsheetXmlOptions options) {
        if (input == null) throw new ArgumentNullException(nameof(input));
        if (!input.CanRead) throw new ArgumentException("Input must be readable.", nameof(input));
        using var bytes = new MemoryStream();
        var buffer = new byte[8192];
        int count;
        while ((count = input.Read(buffer, 0, (int)Math.Min(buffer.Length, (long)options.MaxInputBytes - bytes.Length + 1))) != 0) {
            options.CancellationToken.ThrowIfCancellationRequested();
            if (bytes.Length + count > options.MaxInputBytes) throw new InvalidDataException("XML Spreadsheet exceeds MaxInputBytes.");
            bytes.Write(buffer, 0, count);
        }
        bytes.Position = 0;
        using var xml = XmlReader.Create(bytes, new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = options.MaxInputBytes,
            IgnoreComments = true, IgnoreProcessingInstructions = true
        });
        using var limited = new OfficeXmlLimitingReader(xml, "XML Spreadsheet", options.MaxXmlDepth, options.MaxXmlElements,
            (int)Math.Min(int.MaxValue, 8L * options.MaxXmlElements), options.CancellationToken,
            "Cell", Ss.NamespaceName, options.MaxCells, nameof(options.MaxCells));
        XElement root = XDocument.Load(limited, LoadOptions.PreserveWhitespace).Root
            ?? throw new InvalidDataException("XML Spreadsheet has no root element.");
        if (root.Name != Ss + "Workbook") throw new InvalidDataException("Input is not an XML Spreadsheet 2003 Workbook.");
        var context = new SpreadsheetXmlContext(options);
        context.Inspect(root, new[] { "Styles", "Worksheet", (X + "ExcelWorkbook").ToString() }, Array.Empty<string>());
        if (root.Element(X + "ExcelWorkbook") is XElement excelWorkbook)
            context.Inspect(excelWorkbook, new[] { (X + "Date1904").ToString() }, Array.Empty<string>());
        bool date1904 = root.Element(X + "ExcelWorkbook")?.Element(X + "Date1904") != null;
        XElement[] sheets = root.Elements(Ss + "Worksheet").ToArray();
        if (sheets.Length == 0 || sheets.Length > options.MaxWorksheets) throw new InvalidDataException("XML Spreadsheet requires worksheets within MaxWorksheets.");
        var document = ExcelDocument.Create();
        try {
            var workbookPart = document._spreadSheetDocument.WorkbookPart!;
            var workbook = workbookPart.Workbook!;
            workbook.WorkbookProperties ??= new WorkbookProperties();
            workbook.WorkbookProperties.Date1904 = date1904;
            var styles = SpreadsheetXmlStyles.Read(root.Element(Ss + "Styles"), workbookPart, context);
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (XElement source in sheets) {
                options.CancellationToken.ThrowIfCancellationRequested();
                string name = (string?)source.Attribute(Ss + "Name") ?? throw new InvalidDataException("Worksheet has no name.");
                if (!names.Add(name) || name.Length == 0 || name.Length > 31 || name.IndexOfAny(new[] { '[', ']', ':', '*', '?', '/', '\\' }) >= 0)
                    throw new InvalidDataException("Invalid or duplicate worksheet name.");
                var sheet = document.AddWorksheet(name);
                context.Inspect(source, new[] { "Table", (X + "WorksheetOptions").ToString() }, new[] { "Name" });
                if (source.Element(X + "WorksheetOptions") is XElement worksheetOptions)
                    context.Inspect(worksheetOptions, new[] { (X + "Visible").ToString() }, Array.Empty<string>());
                XElement? visible = source.Element(X + "WorksheetOptions")?.Element(X + "Visible");
                if (visible?.Value is "SheetHidden" or "SheetVeryHidden") sheet.SheetElement.State = visible.Value == "SheetHidden" ? SheetStateValues.Hidden : SheetStateValues.VeryHidden;
                XElement[] tables = source.Elements(Ss + "Table").ToArray();
                if (tables.Length > 1) throw new InvalidDataException("A worksheet cannot contain multiple tables.");
                if (tables.Length == 1) ReadTable(tables[0], sheet, styles, context, date1904);
            }
            return new(document, context.Report());
        } catch { document.Dispose(); throw; }
    }

    private static void ReadTable(XElement table, ExcelSheet sheet, IReadOnlyDictionary<string, uint> styles,
        SpreadsheetXmlContext context, bool date1904) {
        context.Inspect(table, new[] { "Row" }, new[] { "ExpandedRowCount", "ExpandedColumnCount", "StyleID" });
        uint tableStyle = Style(table, styles, styles.TryGetValue("Default", out uint defaultStyle) ? defaultStyle : 0);
        int rowIndex = 0;
        var data = new SheetData();
        foreach (XElement sourceRow in table.Elements(Ss + "Row")) {
            context.Options.CancellationToken.ThrowIfCancellationRequested();
            int nextRow = Index(sourceRow, "Index", rowIndex + 1, 65536);
            Coordinate(nextRow, 1);
            if (nextRow <= rowIndex) throw new InvalidDataException("Worksheet row indexes must increase.");
            rowIndex = nextRow;
            context.Inspect(sourceRow, new[] { "Cell" }, new[] { "Index", "StyleID", "Hidden", "Height", "AutoFitHeight" });
            var row = new Row { RowIndex = (uint)rowIndex };
            if ((string?)sourceRow.Attribute(Ss + "Hidden") == "1") row.Hidden = true;
            string? height = (string?)sourceRow.Attribute(Ss + "Height");
            if (height != null) {
                if (!double.TryParse(height, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) || double.IsNaN(value) || double.IsInfinity(value) || value < 0 || value > 409)
                    throw new InvalidDataException("Invalid XML Spreadsheet row height.");
                row.Height = value; row.CustomHeight = true;
            }
            uint rowStyle = Style(sourceRow, styles, tableStyle);
            int columnIndex = 0;
            foreach (XElement sourceCell in sourceRow.Elements(Ss + "Cell")) {
                context.Cell();
                int nextColumn = Index(sourceCell, "Index", columnIndex + 1, 256);
                Coordinate(rowIndex, nextColumn);
                if (nextColumn <= columnIndex) throw new InvalidDataException("Worksheet cell indexes must increase.");
                columnIndex = nextColumn;
                context.Inspect(sourceCell, new[] { "Data" }, new[] { "Index", "StyleID", "Formula", "MergeAcross", "MergeDown" });
                var cell = new Cell { CellReference = A1.CellReference(rowIndex, columnIndex), StyleIndex = Style(sourceCell, styles, rowStyle) };
                XElement[] values = sourceCell.Elements(Ss + "Data").ToArray();
                if (values.Length > 1) throw new InvalidDataException("A cell cannot contain multiple values.");
                if (values.Length == 1) ReadValue(values[0], cell, date1904, sheet, context);
                string? formula = (string?)sourceCell.Attribute(Ss + "Formula");
                if (formula != null) {
                    if (formula.Length > 8192) throw new InvalidDataException("Invalid XML Spreadsheet formula length.");
                    if (!formula.StartsWith("=", StringComparison.Ordinal)) {
                        context.Omitted("non-R1C1 formula dialect (cached values retained)");
                    } else {
                        string converted = ExcelFormulaSyntaxTree.Parse(formula.Substring(1), ExcelReferenceStyle.R1C1, rowIndex, columnIndex).ConvertReferences(ExcelReferenceStyle.A1);
                        if (cell.DataType?.Value == CellValues.InlineString) {
                            cell.CellValue = new CellValue(cell.InlineString!.InnerText); cell.InlineString = null; cell.DataType = CellValues.String;
                        }
                        cell.CellFormula = new CellFormula(converted);
                    }
                }
                row.Append(cell);
                if (sourceCell.Attribute(Ss + "MergeAcross") != null || sourceCell.Attribute(Ss + "MergeDown") != null) {
                    context.Omitted("merged-cell geometry");
                    int across = Nonnegative(sourceCell, "MergeAcross", 255);
                    Coordinate(rowIndex + Nonnegative(sourceCell, "MergeDown", 65535), columnIndex + across);
                    columnIndex += across;
                }
            }
            data.Append(row);
        }
        var worksheet = sheet.WorksheetPart.Worksheet!;
        worksheet.GetFirstChild<SheetData>()?.Remove();
        worksheet.Append(data);
        Utilities.ExcelXmlPartWriter.SavePreservingLineEndings(sheet.WorksheetPart, worksheet);
    }

    private static int Nonnegative(XElement element, string name, int maximum) {
        string? text = (string?)element.Attribute(Ss + name);
        if (text == null) return 0;
        if (!int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out int value) || value > maximum)
            throw new InvalidDataException("Invalid XML Spreadsheet " + name + ".");
        return value;
    }

    private static uint Style(XElement element, IReadOnlyDictionary<string, uint> styles, uint fallback) {
        string? name = (string?)element.Attribute(Ss + "StyleID");
        if (name == null) return fallback;
        return styles.TryGetValue(name, out uint value) ? value : throw new InvalidDataException("Unknown XML Spreadsheet style: " + name);
    }

    private static void ReadValue(XElement source, Cell cell, bool date1904, ExcelSheet sheet, SpreadsheetXmlContext context) {
        context.Inspect(source, Array.Empty<string>(), new[] { "Type" });
        string text = source.Value;
        if (text.Length > 32767) throw new InvalidDataException("XML Spreadsheet cell text exceeds Excel's cell limit.");
        switch ((string?)source.Attribute(Ss + "Type")) {
            case "String": cell.DataType = CellValues.InlineString; cell.InlineString = new InlineString(new Text(text) { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }); break;
            case "Number":
                if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double number) || double.IsNaN(number) || double.IsInfinity(number)) throw new InvalidDataException("Invalid XML Spreadsheet number.");
                cell.DataType = CellValues.Number; cell.CellValue = new CellValue(number.ToString("R", CultureInfo.InvariantCulture)); break;
            case "Boolean":
                if (text != "0" && text != "1") throw new InvalidDataException("Invalid XML Spreadsheet Boolean.");
                cell.DataType = CellValues.Boolean; cell.CellValue = new CellValue(text); break;
            case "DateTime":
                DateTime date;
                try { date = XmlConvert.ToDateTime(text, XmlDateTimeSerializationMode.Unspecified); }
                catch (FormatException exception) { throw new InvalidDataException("Invalid XML Spreadsheet date.", exception); }
                double serial = ExcelDateSystemConverter.ToSerial(date, date1904 ? ExcelDateSystem.NineteenFour : ExcelDateSystem.NineteenHundred);
                cell.DataType = CellValues.Number; cell.CellValue = new CellValue(serial.ToString("R", CultureInfo.InvariantCulture));
                cell.StyleIndex = SpreadsheetXmlStyles.EnsureDateStyle(sheet.WorksheetPart.GetParentParts().OfType<DocumentFormat.OpenXml.Packaging.WorkbookPart>().Single(), cell.StyleIndex?.Value ?? 0);
                break;
            case "Error":
                if (!new[] { "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A" }.Contains(text)) throw new InvalidDataException("Invalid XML Spreadsheet error.");
                cell.DataType = CellValues.Error; cell.CellValue = new CellValue(text); break;
            default: throw new InvalidDataException("Unsupported XML Spreadsheet cell type.");
        }
    }
}
