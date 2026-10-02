using System.Globalization;
using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.IWork.Benchmarks;

public sealed partial class IWorkRuntimeWorkload {
    private readonly bool _native;
    private static readonly string[] NativePagesText = ["hello pages", "", "second paragraph with some words\""];
    private static readonly object[,] NativeNumbersCells = { { "a", "b", "C" }, { 1d, 2d, 3d }, { "x", "y", "Z" } };

    /// <summary>Loads a pinned independent corpus fixture outside measurement.</summary>
    public IWorkRuntimeWorkload(string kind, string corpusRoot) {
        _kind = Enum.Parse<IWorkDocumentKind>(kind, ignoreCase: false);
        var (name, hash, units) = _kind switch {
            IWorkDocumentKind.Pages => ("simple.pages", "5AEE6D03277D2DB2104F593E64AFE081DEC539F0117B97124B6F99158124C93E", 3),
            IWorkDocumentKind.Numbers => ("simple.numbers", "D0B00D9CAE5985CCCAA3B2FB251FAE92EB0E38360FB4B5DF8B4350EB658F752B", 9),
            _ => throw new ArgumentException("Native runtime fixtures currently support Pages and Numbers.", nameof(kind))
        };
        _input = File.ReadAllBytes(Path.Combine(corpusRoot, "nim-iwork", name));
        InputSha256 = Convert.ToHexString(SHA256.HashData(_input));
        if (InputSha256 != hash) throw new InvalidDataException("Native runtime fixture differs from its qualified hash.");
        _units = units;
        _native = true;
    }

    private void ValidateNative() {
        VerifiedUnits = 0;
        if (_operation == "LoadProject") {
            if (_projection is IWorkPagesProjection pages) {
                if (!pages.HasEditableContent) throw new InvalidDataException("Native Pages projection is incomplete.");
                ValidateNativeText(pages.Body.Paragraphs.Select(paragraph => paragraph.Text));
            } else if (_projection is IWorkNumbersProjection numbers) {
                if (!numbers.HasEditableContent) throw new InvalidDataException("Native Numbers projection is incomplete.");
                IWorkTable table = numbers.Sheets.Single().Tables.Single();
                if (table.RowCount != 3 || table.ColumnCount != 3 || table.Cells.Count != 9)
                    throw new InvalidDataException("Native Numbers dimensions or cell count mismatch.");
                for (int row = 1; row <= 3; row++) {
                    for (int column = 1; column <= 3; column++) {
                        IWorkTableCell? cell = table.GetCell(row, column);
                        if (cell is null || cell.Row != row || cell.Column != column || cell.Formula is not null)
                            throw new InvalidDataException("Missing or unexpected native Numbers cell.");
                        ValidateNativeCell(row, column, cell.Value);
                        VerifiedUnits++;
                    }
                }
            } else throw new InvalidDataException("Missing native projection.");
        } else if (_operation == "ConvertSave" && _output is { Length: > 0 }) {
            using var saved = new MemoryStream(_output, writable: false);
            if (_kind == IWorkDocumentKind.Pages) {
                using var document = WordprocessingDocument.Open(saved, false);
                ValidateNativeText(document.MainDocumentPart!.Document!.Body!
                    .Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Select(paragraph => paragraph.InnerText));
            } else {
                using var document = SpreadsheetDocument.Open(saved, false);
                var workbook = document.WorkbookPart!;
                var rows = workbook.WorksheetParts.Single().Worksheet!.GetFirstChild<SheetData>()!.Elements<Row>().ToArray();
                if (rows.Length != 3) throw new InvalidDataException("Native XLSX row count mismatch.");
                for (int row = 1; row <= 3; row++) {
                    if (rows[row - 1].RowIndex?.Value != row) throw new InvalidDataException("Native XLSX row coordinate mismatch.");
                    Cell[] cells = rows[row - 1].Elements<Cell>().ToArray();
                    if (cells.Length != 3) throw new InvalidDataException("Native XLSX column count mismatch.");
                    for (int column = 1; column <= 3; column++) {
                        Cell cell = cells[column - 1];
                        string address = ((char)('A' + column - 1)).ToString() + row.ToString(CultureInfo.InvariantCulture);
                        if (cell.CellReference?.Value != address || cell.CellFormula is not null)
                            throw new InvalidDataException("Unexpected native XLSX coordinate or formula.");
                        object? value;
                        if (cell.DataType?.Value == CellValues.SharedString)
                            value = workbook.SharedStringTablePart!.SharedStringTable!.Elements<SharedStringItem>()
                                .ElementAt(int.Parse(cell.CellValue!.Text, CultureInfo.InvariantCulture)).InnerText;
                        else if (cell.DataType?.Value == CellValues.InlineString) value = cell.InlineString!.InnerText;
                        else if (cell.DataType is null || cell.DataType.Value == CellValues.Number)
                            value = double.Parse(cell.CellValue!.Text, CultureInfo.InvariantCulture);
                        else throw new InvalidDataException("Unexpected native XLSX cell type.");
                        ValidateNativeCell(row, column, value);
                        VerifiedUnits++;
                    }
                }
            }
        } else throw new InvalidDataException("Missing native saved result.");
        if (VerifiedUnits != _units) throw new InvalidDataException("Native workload unit count mismatch.");
    }

    private void ValidateNativeText(IEnumerable<string> text) {
        if (!text.SequenceEqual(NativePagesText)) throw new InvalidDataException("Native Pages text or paragraph order mismatch.");
        VerifiedUnits = NativePagesText.Length;
    }

    private static void ValidateNativeCell(int row, int column, object? value) {
        if (!Equals(NativeNumbersCells[row - 1, column - 1], value))
            throw new InvalidDataException("Native Numbers cell type or value mismatch.");
    }
}
