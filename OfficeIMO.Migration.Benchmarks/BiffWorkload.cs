using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;

namespace OfficeIMO.Migration.Benchmarks;

/// <summary>Exercises ordered cell findings followed by repeated unresolved number-format declarations in BIFF4W.</summary>
internal sealed class BiffWorkload : MigrationWorkload {
    private ExcelDocument? _document;
    private ExcelWorkbookInfo? _inspection;
    private OfficeConversionReport? _report;

    internal BiffWorkload(int units) : base(CreateInput(units), units) { }

    public override void Execute(string operation) {
        using var source = new MemoryStream(Input, writable: false);
        _document = ExcelDocument.Load(source, null, out OfficeConversionReport report);
        _report = report;
        if (operation == "LoadInspect") _inspection = _document.Inspect();
        else if (operation == "ConvertSave") {
            using var output = new MemoryStream();
            _document.Save(output, new ExcelSaveOptions { AllowLoss = true });
            Output = output.ToArray();
        } else throw new ArgumentOutOfRangeException(nameof(operation));
    }

    public override void Validate() {
        if (_document == null || _report == null) throw new InvalidOperationException("The BIFF4 operation did not load a workbook.");
        Require(_document.SourceFormat == ExcelFileFormat.Xls && _report.HasLoss, "source format and explicit loss");
        Require(_report.Diagnostics.Count == Units + 1, "diagnostic count");
        var digest = new StringBuilder();
        for (int index = 0; index <= Units; index++) {
            OfficeDiagnostic finding = _report.Diagnostics[index];
            string expectedCode = index == Units ? "early-biff-number-format" : "early-biff-cell-style";
            string? expectedLocation = index == Units ? null : "'Cells'!A" + (index + 1).ToString(CultureInfo.InvariantCulture);
            string expectedMessage = index == Units ? "A referenced early BIFF number format could not be resolved."
                : "A cell references a missing XF record.";
            Require(finding.Code == expectedCode && finding.Location == expectedLocation && finding.Message == expectedMessage
                && finding.Severity == OfficeDiagnosticSeverity.Warning && finding.LossKind == OfficeConversionLossKind.Approximation
                && finding.Format == "OfficeIMO.LegacyImport", "ordered diagnostic " + index);
            digest.Append(finding.Code).Append('|').Append(finding.Location).Append('|').Append(finding.Severity)
                .Append('|').Append(finding.LossKind).Append('|').Append(finding.Format).Append('|').Append(finding.Message).Append('\n');
        }
        DiagnosticDigest = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(digest.ToString())));
        ValidateInspection(_inspection ?? _document.Inspect());
        if (Output != null) {
            using (var bytes = new MemoryStream(Output, writable: false))
            using (var package = SpreadsheetDocument.Open(bytes, isEditable: false)) {
                var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(package).Take(5).ToArray();
                Require(errors.Length == 0, "saved Open XML schema: " + string.Join("; ", errors.Select(error => error.Description)));
                WorkbookPart workbook = package.WorkbookPart!;
                Sheet[] sheets = (workbook.Workbook ?? throw new InvalidDataException("Missing saved workbook.")).Sheets!.Elements<Sheet>().ToArray();
                Require(sheets.Length == 2 && sheets[0].Name?.Value == "Cells" && sheets[1].Name?.Value == "Formats", "saved sheet identity and order");
                for (int sheetIndex = 0; sheetIndex < sheets.Length; sheetIndex++) {
                    var part = (WorksheetPart)workbook.GetPartById(sheets[sheetIndex].Id!.Value!);
                    Cell[] cells = (part.Worksheet ?? throw new InvalidDataException("Missing saved worksheet.")).Descendants<Cell>().ToArray();
                    Require(cells.Length == (sheetIndex == 0 ? Units : 0), "saved cell count " + sheetIndex);
                    for (int index = 0; index < cells.Length; index++) {
                        string value = (index + 1).ToString(CultureInfo.InvariantCulture);
                        Require(cells[index].CellReference?.Value == "A" + value && cells[index].CellFormula == null
                            && (cells[index].DataType == null || cells[index].DataType!.Value == CellValues.Number)
                            && cells[index].CellValue?.Text == value, "saved ordered numeric cell " + index);
                    }
                }
            }
            using var source = new MemoryStream(Output, writable: false);
            using ExcelDocument reopened = ExcelDocument.Load(source, null, out OfficeConversionReport reopenedReport);
            Require(reopened.SourceFormat == ExcelFileFormat.Xlsx && reopenedReport.Diagnostics.Count == 0 && !reopenedReport.HasLoss, "saved reopen contract");
            ValidateInspection(reopened.Inspect());
        }
        VerifiedUnits = Units;
    }

    public override void ReleaseResults() {
        _inspection = null;
        _report = null;
        _document?.Dispose();
        _document = null;
        base.ReleaseResults();
    }

    private void ValidateInspection(ExcelWorkbookInfo inspection) {
        Require(inspection.Worksheets.Count == 2 && inspection.Worksheets[0].Name == "Cells"
            && inspection.Worksheets[1].Name == "Formats" && inspection.Worksheets[0].Index == 0
            && inspection.Worksheets[1].Index == 1, "inspection sheet identity and order");
        Require(inspection.Worksheets[0].Cells.Count == Units && inspection.Worksheets[1].Cells.Count == 0, "inspection cell count");
        for (int index = 0; index < Units; index++) {
            ExcelCellInfo cell = inspection.Worksheets[0].Cells[index];
            Require(cell.Row == index + 1 && cell.Column == 1 && cell.Formula == null
                && cell.Value is double value && BitConverter.DoubleToInt64Bits(value) == BitConverter.DoubleToInt64Bits(index + 1d),
                "inspection ordered typed numeric cell " + index);
        }
    }

    private static void Require(bool condition, string boundary) {
        if (!condition) throw new InvalidDataException("The finding-heavy BIFF4 workload changed " + boundary + ".");
    }

    private static byte[] CreateInput(int units) {
        if (units < 1 || units > 16_384) throw new ArgumentOutOfRangeException(nameof(units));
        using var bytes = new MemoryStream();
        using var writer = new BinaryWriter(bytes, Encoding.ASCII, leaveOpen: true);
        WriteRecord(writer, 0x0409, record => WriteBof(record, 0x0100));
        WriteRecord(writer, 0x008E, record => record.Write(0));
        foreach (string name in new[] { "Cells", "Formats" }) WriteRecord(writer, 0x0085, record => WriteName(record, name));
        long firstHeader = bytes.Position;
        bytes.Position = 14;
        writer.Write(checked((int)firstHeader));
        bytes.Position = firstHeader;
        foreach (string name in new[] { "Cells", "Formats" }) {
            using var sheetBytes = new MemoryStream();
            using (var sheet = new BinaryWriter(sheetBytes, Encoding.ASCII, leaveOpen: true)) {
                WriteRecord(sheet, 0x0409, record => WriteBof(record, 0x0010));
                for (int index = 0; index < units; index++) {
                    if (name == "Cells") WriteRecord(sheet, 0x0203, record => {
                        record.Write(checked((ushort)index));
                        record.Write((ushort)0);
                        record.Write((ushort)65534);
                        record.Write(index + 1d);
                    });
                    else WriteRecord(sheet, 0x0443, record => record.Write(new byte[] { 0, 1, 0xF1, 0xFF, 0, 0xFC, 0, 0, 0, 0, 0, 0 }));
                }
                WriteRecord(sheet, 0x000A, _ => { });
            }
            WriteRecord(writer, 0x008F, record => {
                record.Write(checked((int)sheetBytes.Length));
                WriteName(record, name);
            });
            writer.Write(sheetBytes.ToArray());
        }
        WriteRecord(writer, 0x000A, _ => { });
        return bytes.ToArray();
    }

    private static void WriteBof(BinaryWriter writer, ushort kind) {
        writer.Write((ushort)0);
        writer.Write(kind);
        writer.Write((ushort)0);
    }

    private static void WriteName(BinaryWriter writer, string name) {
        byte[] text = Encoding.ASCII.GetBytes(name);
        writer.Write(checked((byte)text.Length));
        writer.Write(text);
    }

    private static void WriteRecord(BinaryWriter writer, ushort kind, Action<BinaryWriter> write) {
        using var bytes = new MemoryStream();
        using (var payload = new BinaryWriter(bytes, Encoding.ASCII, leaveOpen: true)) write(payload);
        writer.Write(kind);
        writer.Write(checked((ushort)bytes.Length));
        writer.Write(bytes.ToArray());
    }
}

