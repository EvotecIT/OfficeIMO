using System.Globalization;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Compound;
using OfficeIMO.Excel.LegacyXls.Model;
using OfficeIMO.Excel.Utilities;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    private const int LegacySpreadsheet = OfficeXmlNamespaces.SpreadsheetML;

    /// <summary>Projects a plain BIFF sheet using the workbook shell, owned XML trees and shared-string writer.</summary>
    private static bool TryProjectPlainLegacyXls(LegacyXlsWorkbook source, CancellationToken token,
        out byte[]? bytes, out ExcelDeferredWorkbook? workbook) {
        bytes = null;
        workbook = null;
        if (source.WasEncryptedSource || source.Worksheets.Count != 1 || source.ChartSheets.Count != 0
            || source.SourceCompoundFile?.Streams is not { Count: 1 } streams
            || LegacyWorkbookStreamLocator.FindWorkbookStream(streams) is not { } nativeStream
            || !HasPlainLegacyXlsRecords(nativeStream, source.Worksheets[0].StreamOffset, token)) {
            return false;
        }

        LegacyXlsWorksheet sheet = source.Worksheets[0];
        string name = sheet.Name;
        if (sheet.Visibility != 0 || name.Length is 0 or > 31 || name != TrimSanitizedSheetNameEdges(name)
            || name.Contains("__", StringComparison.Ordinal) || name.Any(character => char.IsSurrogate(character)
                || char.IsControl(character) || character is ':' or '\\' or '/' or '?' or '*' or '[' or ']')
            || sheet.Cells.Count == 0) {
            return false;
        }

        using OfficeOpcEditablePackage package = OfficeOpcEditablePackage.Create();
        PopulateWorkbookShell(package, CreatedMainContentType(ExcelDocumentType.Workbook)!, theme: null);
        const string worksheetPart = "/xl/worksheets/sheet1.xml";
        const string worksheetTarget = "worksheets/sheet1.xml";
        string relationship = package.AddRelationship(CreatedWorkbookPartName, WorksheetRelationshipType, worksheetTarget);
        OfficeXmlTree workbookTree = package.GetXml(CreatedWorkbookPartName);
        var root = new OfficeXmlNode(workbookTree, workbookTree.Root);
        OfficeXmlNode properties = AddLegacySpreadsheetElement(root, "workbookPr");
        if (source.Uses1904DateSystem) properties.SetBooleanAttribute("date1904", true);
        AddLegacySpreadsheetElement(AddLegacySpreadsheetElement(root, "bookViews"), "workbookView");
        OfficeXmlNode entry = AddLegacySpreadsheetElement(AddLegacySpreadsheetElement(root, "sheets"), "sheet");
        entry.SetAttribute("name", name);
        entry.SetUInt32Attribute("sheetId", 1);
        entry.SetAttribute(OfficeXmlNamespaces.Relationships, "id", relationship);

        OfficeXmlTree sheetTree = package.CreateXmlPart(worksheetPart, WorksheetContentType, LegacySpreadsheet, "worksheet");
        var worksheet = new OfficeXmlNode(sheetTree, sheetTree.Root);
        OfficeXmlNode dimension = AddLegacySpreadsheetElement(worksheet, "dimension");
        OfficeXmlNode view = AddLegacySpreadsheetElement(AddLegacySpreadsheetElement(worksheet, "sheetViews"), "sheetView");
        view.SetUInt32Attribute("workbookViewId", 0);
        OfficeXmlNode selection = AddLegacySpreadsheetElement(view, "selection");
        selection.SetAttribute("activeCell", "A1");
        selection.SetAttribute("sqref", "A1");
        OfficeXmlNode grid = AddLegacySpreadsheetElement(worksheet, "sheetData");
        ExcelDeferredSharedStrings strings = ExcelDeferredSharedStrings.Load(package, partName: null)!;

        int firstColumn = int.MaxValue, lastColumn = 0, row = 0, lastColumnInRow = 0;
        OfficeXmlNode rowNode = default;
        foreach (LegacyXlsCell cell in sheet.Cells) {
            token.ThrowIfCancellationRequested();
            if (cell.Row < 1 || cell.Column < 1 || cell.Row > 65536 || cell.Column > 256 || cell.Row < row
                || cell.Row == row && cell.Column <= lastColumnInRow || cell.StyleIndex != 0
                || cell.IsFormula || cell.TextFormattingRuns.Count != 0) {
                return false;
            }

            string type, value;
            switch (cell.Kind) {
                case LegacyXlsCellValueKind.Text when cell.Value is string text:
                    if (!strings.TryGetOrAdd(ExcelSanitizer.SanitizeString(text), ExcelSheet.CellValuePlainStringPromotionSharedStringCount,
                        out int index)) return false;
                    strings.AddedReferences++;
                    type = "s";
                    value = index.ToString(CultureInfo.InvariantCulture);
                    break;
                case LegacyXlsCellValueKind.Number when cell.Value is double number && !double.IsNaN(number) && !double.IsInfinity(number):
                    type = "n";
                    value = OfficeRoundTripText.Format(number, CultureInfo.InvariantCulture);
                    break;
                case LegacyXlsCellValueKind.Boolean when cell.Value is bool flag:
                    type = "b";
                    value = flag ? "1" : "0";
                    break;
                default:
                    return false;
            }

            if (row != cell.Row) {
                row = cell.Row;
                rowNode = AddLegacySpreadsheetElement(grid, "row");
                rowNode.SetUInt32Attribute("r", checked((uint)row));
            }
            lastColumnInRow = cell.Column;
            firstColumn = Math.Min(firstColumn, cell.Column);
            lastColumn = Math.Max(lastColumn, cell.Column);
            OfficeXmlNode cellNode = AddLegacySpreadsheetElement(rowNode, "c");
            cellNode.SetAttribute("r", ExcelA1.CellReference(row, cell.Column));
            cellNode.SetAttribute("t", type);
            AddLegacySpreadsheetElement(cellNode, "v").SetInnerText(value);
        }

        string first = ExcelA1.CellReference(sheet.Cells[0].Row, firstColumn);
        string last = ExcelA1.CellReference(row, lastColumn);
        dimension.SetAttribute("ref", first == last ? first : first + ":" + last);
        if (!strings.TryWrite(package, CreatedWorkbookPartName)) return false;

        token.ThrowIfCancellationRequested();
        using var output = new MemoryStream(16 * 1024);
        package.Save(output);
        bytes = output.ToArray();
        workbook = new ExcelDeferredWorkbook(CreatedWorkbookPartName,
            new List<ExcelDeferredSheet> { new(name, 1, relationship, worksheetPart, isWorksheet: true) },
            source.Uses1904DateSystem);
        return true;
    }

    private static OfficeXmlNode AddLegacySpreadsheetElement(OfficeXmlNode parent, string name) =>
        parent.Append(parent.CreateElement(LegacySpreadsheet, name));

    // Only records whose complete worksheet projection is implemented above can take the owned path.
    // The parsed BIFF model remains authoritative for values and native data retention.
    private static bool HasPlainLegacyXlsRecords(byte[] stream, int sheetOffset, CancellationToken token) {
        int offset = 0;
        int completedSubstreams = 0;
        while (offset <= stream.Length - 4) {
            token.ThrowIfCancellationRequested();
            var type = (BiffRecordType)BiffRecordReader.ReadUInt16(stream, offset);
            ushort length = BiffRecordReader.ReadUInt16(stream, offset + 2);
            if (length > stream.Length - offset - 4) return false;

            if (offset == 0 || offset == sheetOffset) {
                ushort substream = offset == 0 ? (ushort)0x0005 : (ushort)0x0010;
                if (type != BiffRecordType.Bof || length < 4
                    || BiffRecordReader.ReadUInt16(stream, offset + 4) != LegacyBiffVersionValidator.Biff8Version
                    || BiffRecordReader.ReadUInt16(stream, offset + 6) != substream) return false;
            } else if (type is not (BiffRecordType.Eof or BiffRecordType.BoundSheet8 or BiffRecordType.Sst
                or BiffRecordType.Continue or BiffRecordType.CodePage or BiffRecordType.Date1904 or BiffRecordType.Dimensions
                or BiffRecordType.Number or BiffRecordType.Label or BiffRecordType.BoolErr or BiffRecordType.LabelSst
                or BiffRecordType.Rk or BiffRecordType.MulRk)) {
                return false;
            }

            offset += 4 + length;
            if (type != BiffRecordType.Eof) continue;
            if (length != 0) return false;
            if (++completedSubstreams == 1) {
                if (offset != sheetOffset) return false;
            } else {
                // Some compound files include sector padding in the recorded Workbook stream size.
                for (; offset < stream.Length; offset++) {
                    if ((offset & 4095) == 0) token.ThrowIfCancellationRequested();
                    if (stream[offset] != 0) return false;
                }
                return true;
            }
        }
        return false;
    }
}
