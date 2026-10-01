using System.Globalization;
using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.IWork;
using OfficeIMO.Word.IWork;

namespace OfficeIMO.IWork.Benchmarks;

/// <summary>Supported iWork workloads and semantic validation; PowerForge owns measurement.</summary>
public sealed class IWorkRuntimeWorkload {
    private readonly IWorkDocumentKind _kind;
    private readonly int _units;
    private readonly byte[] _input;
    private object? _projection;
    private byte[]? _output;
    private string? _operation;

    /// <summary>Creates deterministic input outside the measured operation.</summary>
    public IWorkRuntimeWorkload(string kind, int units) {
        _kind = Enum.Parse<IWorkDocumentKind>(kind, ignoreCase: false);
        _units = units; _input = IWorkScaleInput.Create(_kind, units);
        InputSha256 = Convert.ToHexString(SHA256.HashData(_input));
    }
    /// <summary>Gets the packaged ZIP input size.</summary>
    public long InputBytes => _input.LongLength;
    /// <summary>Gets the deterministic source hash.</summary>
    public string InputSha256 { get; }
    /// <summary>Gets the last saved Office package size; zero for projection-only work.</summary>
    public long OutputBytes => _output?.LongLength ?? 0;
    /// <summary>Gets the verified number of paragraphs, cells or slides.</summary>
    public int VerifiedUnits { get; private set; }

    /// <summary>Loads/projects or converts/saves once; excludes input generation and readback validation.</summary>
    public void Execute(string operation) {
        _operation = operation; _projection = null; _output = null; VerifiedUnits = 0;
        using var input = new MemoryStream(_input, writable: false);
        IWorkSourceDocument source = IWorkSourceDocument.Open(input, _kind);
        if (operation == "LoadProject") {
            _projection = _kind switch {
                IWorkDocumentKind.Pages => source.ReadPages(),
                IWorkDocumentKind.Numbers => source.ReadNumbers(),
                IWorkDocumentKind.Keynote => source.ReadKeynote(),
                _ => throw new InvalidOperationException()
            };
            return;
        }
        if (operation != "ConvertSave") throw new ArgumentException("Unknown iWork workload operation.", nameof(operation));
        using var saved = new MemoryStream();
        if (_kind == IWorkDocumentKind.Pages) {
            using var result = source.ToWordDocumentResult();
            result.Report.RequireCompleteEditableReconstruction(); result.Value.Save(saved);
        } else if (_kind == IWorkDocumentKind.Numbers) {
            using var result = source.ToExcelDocumentResult();
            result.Report.RequireCompleteEditableReconstruction(); result.Value.Save(saved);
        } else {
            using var result = source.ToPowerPointPresentationResult();
            result.Report.RequireCompleteEditableReconstruction(); result.Value.Save(saved);
        }
        _output = saved.ToArray();
    }

    /// <summary>Checks every source or saved destination unit outside timing; failures fail the lane.</summary>
    public void Validate() {
        if (_operation == "LoadProject") {
            if (_projection is IWorkPagesProjection pages) {
                if (!pages.HasEditableContent) throw new InvalidDataException("Pages input is incomplete.");
                ValidateText(pages.Body.Paragraphs.Select(p => p.Text));
            } else if (_projection is IWorkNumbersProjection numbers) {
                if (!numbers.HasEditableContent) throw new InvalidDataException("Numbers input is incomplete.");
                IWorkTable table = numbers.Sheets.Single().Tables.Single();
                if (table.RowCount != _units || table.ColumnCount != 1 || table.Cells.Count != _units)
                    throw new InvalidDataException("Numbers dimensions or sparse cell count mismatch.");
                int index = 0;
                foreach (IWorkTableCell cell in table.Cells.OrderBy(cell => cell.Row)) {
                    index++;
                    if (cell.Row != index || cell.Column != 1 || cell.Kind != IWorkCellKind.Number
                        || Convert.ToDouble(cell.Value, CultureInfo.InvariantCulture) != index)
                        throw new InvalidDataException("Numbers source coordinate/value mismatch.");
                }
                VerifiedUnits = index;
            } else if (_projection is IWorkKeynoteProjection keynote) {
                if (!keynote.HasEditableContent) throw new InvalidDataException("Keynote input is incomplete: "
                    + string.Join("; ", keynote.Diagnostics.Select(d => d.Code + ": " + d.Message)));
                ValidateText(keynote.Slides.Select(slide => slide.Title));
            } else throw new InvalidDataException("Missing projection result.");
        } else if (_operation == "ConvertSave" && _output is { Length: > 0 }) {
            using var saved = new MemoryStream(_output, writable: false);
            if (_kind == IWorkDocumentKind.Pages) {
                using var document = WordprocessingDocument.Open(saved, false);
                ValidateText(document.MainDocumentPart!.Document!.Body!
                    .Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Select(p => p.InnerText));
            } else if (_kind == IWorkDocumentKind.Numbers) {
                using var document = SpreadsheetDocument.Open(saved, false);
                WorksheetPart sheet = document.WorkbookPart!.WorksheetParts.Single();
                var rows = sheet.Worksheet!.GetFirstChild<SheetData>()!.Elements<Row>().ToArray();
                if (rows.Length != _units) throw new InvalidDataException("XLSX row count mismatch.");
                for (int i = 0; i < rows.Length; i++) {
                    Cell cell = rows[i].Elements<Cell>().Single();
                    if (rows[i].RowIndex!.Value != i + 1 || cell.CellReference!.Value != "A" + (i + 1).ToString(CultureInfo.InvariantCulture)
                        || (cell.DataType is not null && cell.DataType.Value != CellValues.Number)
                        || cell.CellFormula is not null
                        || double.Parse(cell.CellValue!.Text, CultureInfo.InvariantCulture) != i + 1)
                        throw new InvalidDataException("XLSX coordinate/value mismatch.");
                }
                VerifiedUnits = rows.Length;
            } else {
                using PowerPointPresentation document = PowerPointPresentation.Load(saved);
                ValidateText(document.Slides.Select(slide => slide.TextBoxes.Single().Text));
            }
        } else throw new InvalidDataException("Missing saved result.");
        if (VerifiedUnits != _units) throw new InvalidDataException("iWork workload unit count mismatch.");
    }

    private void ValidateText(IEnumerable<string> values) {
        int index = 0;
        foreach (string text in values) {
            index++;
            if (text != IWorkScaleInput.Text(index)) throw new InvalidDataException("iWork text or order mismatch.");
        }
        VerifiedUnits = index;
    }
}
