using System.Management.Automation;
using System.Security;
using System.Text;
using OfficeIMO;
using OfficeIMO.Core.Internal;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Models;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Comparison;

namespace PSWriteOffice.Cmdlets.Conversion;

/// <summary>Compares native document content or rendered pages and optionally publishes a review artifact.</summary>
[Cmdlet(VerbsData.Compare, "OfficeDocument", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(OfficeComparisonReport))]
public sealed class CompareOfficeDocumentCommand : OfficeCmdlet, IDisposable {
    private readonly CancellationTokenSource _cancel = new();
    /// <summary>Original document path.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Path")]
    public string ReferencePath { get; set; } = string.Empty;
    /// <summary>Modified document path.</summary>
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "Path")]
    public string DifferencePath { get; set; } = string.Empty;
    /// <summary>Caller-owned original native model.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Document"), ValidateNotNull]
    public object? ReferenceDocument { get; set; }
    /// <summary>Caller-owned modified native model.</summary>
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "Document"), ValidateNotNull]
    public object? DifferenceDocument { get; set; }
    /// <summary>Compare native rendered pages instead of semantic content.</summary>
    [Parameter]
    public SwitchParameter Visual { get; set; }
    /// <summary>Ordered one-based pages used by visual comparison.</summary>
    [Parameter, OfficePageRangeTransformation]
    public OfficePageRange? Page { get; set; }
    /// <summary>Optional Word redline, Excel findings workbook or visual HTML gallery.</summary>
    [Parameter]
    public string? OutputPath { get; set; }
    /// <summary>Typed native comparison settings or native property overrides.</summary>
    [Parameter]
    public object[]? Options { get; set; }
    /// <summary>Maximum native findings returned; Excel also enforces this during its comparison.</summary>
    [Parameter, ValidateRange(1, int.MaxValue)]
    public int? MaximumDifferences { get; set; }
    /// <summary>Original file's password.</summary>
    [Parameter(ParameterSetName = "Path"), OfficePasswordTransformation]
    public SecureString? Password { get; set; }
    /// <summary>Modified file's password; omission reuses Password.</summary>
    [Parameter(ParameterSetName = "Path"), OfficePasswordTransformation]
    public SecureString? DifferencePassword { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (!Visual && Page != null) throw new PSArgumentException("Page selection applies to visual comparison.");
        var engine = (EngineIntrinsics)SessionState.PSVariable.GetValue("ExecutionContext");
        var pdfSettings = Visual ? OfficeComparisonOptionsBinding.CapturePdf(engine, Options) : null;
        if (ParameterSetName == "Path" && pdfSettings != null) {
            if (MyInvocation.BoundParameters.ContainsKey(nameof(Password))) pdfSettings.ReferenceLoadOptions = (pdfSettings.ReferenceLoadOptions ?? new PdfLoadOptions()).Clone(null, Password.ToPlainText());
            if (MyInvocation.BoundParameters.ContainsKey(nameof(DifferencePassword))) pdfSettings.DifferenceLoadOptions = (pdfSettings.DifferenceLoadOptions ?? new PdfLoadOptions()).Clone(null, DifferencePassword.ToPlainText());
            else if (MyInvocation.BoundParameters.ContainsKey(nameof(Password))) pdfSettings.DifferenceLoadOptions = (pdfSettings.DifferenceLoadOptions ?? new PdfLoadOptions()).Clone(null, Password.ToPlainText());
        }
        if (ParameterSetName == "Document" && pdfSettings != null &&
            (pdfSettings.ReferenceLoadOptions != null || pdfSettings.DifferenceLoadOptions != null))
            throw new PSArgumentException("PDF load policies apply to file sources; a live document already owns its parsed source policy.");
        var owned = new List<IDisposable>();
        try {
            object left, right;
            string? leftPath = null, rightPath = null;
            if (ParameterSetName == "Path") {
                leftPath = OfficePath.ResolveInput(SessionState, ReferencePath, literal: true)[0];
                rightPath = OfficePath.ResolveInput(SessionState, DifferencePath, literal: true)[0];
                left = Load(leftPath, Password.ToPlainText(), pdfSettings?.ReferenceLoadOptions); owned.Add((IDisposable)left);
                right = Load(rightPath, (MyInvocation.BoundParameters.ContainsKey(nameof(DifferencePassword)) ? DifferencePassword : Password).ToPlainText(), pdfSettings?.DifferenceLoadOptions); owned.Add((IDisposable)right);
            } else {
                left = OfficeArgumentValue.Unwrap(ReferenceDocument) ?? throw new PSArgumentException("ReferenceDocument is required.");
                right = OfficeArgumentValue.Unwrap(DifferenceDocument) ?? throw new PSArgumentException("DifferenceDocument is required.");
            }
            _cancel.Token.ThrowIfCancellationRequested();
            string? output = OutputPath == null ? null : ResolveOutputPath(OutputPath);
            OfficeComparisonReport report;
            if (Visual) {
                var conversions = new List<OfficeConversionReport>();
                var expected = Pdf(left, owned, conversions);
                var actual = Pdf(right, owned, conversions);
                var native = expected.Proof.CompareVisual(actual, _cancel.Token, Page?.Selection, pdfSettings!.Comparison);
                var findings = native.StructuralDifferences.Cast<object>().Concat(native.Pages.Where(page => !page.IsMatch).Cast<object>());
                report = new OfficeComparisonReport(native, native.IsMatch, findings, MaximumDifferences, conversions);
                if (output != null && Approve(output, "Write visual comparison gallery")) {
                    if (!new[] { ".html", ".htm" }.Contains(System.IO.Path.GetExtension(output), StringComparer.OrdinalIgnoreCase)) throw new PSArgumentException("Visual OutputPath must name an HTML gallery.");
                    byte[] bytes = Encoding.UTF8.GetBytes(native.ToHtmlGallery(null, pdfSettings.Comparison.MaxTotalOutputBytes, _cancel.Token));
                    WriteArtifact(output, temporary => File.WriteAllBytes(temporary, bytes)); report.OutputPath = output;
                }
            } else if (left is WordDocument referenceWord && right is WordDocument differenceWord) {
                var settings = OfficeComparisonOptionsBinding.Capture<WordComparisonOptions>(engine, Options);
                var native = WordDocumentComparer.CompareStructure(referenceWord, differenceWord, settings, leftPath ?? "reference", rightPath ?? "difference");
                report = new OfficeComparisonReport(native, !native.HasChanges, native.Findings.Cast<object>(), MaximumDifferences);
                if (output != null && Approve(output, "Write Word comparison redline")) {
                    if (!System.IO.Path.GetExtension(output).Equals(".docx", StringComparison.OrdinalIgnoreCase)) throw new PSArgumentException("A Word redline OutputPath must end in .docx.");
                    if ((leftPath ?? referenceWord.FilePath) is string referencePath && OfficePath.AreEquivalent(output, referencePath) ||
                        (rightPath ?? differenceWord.FilePath) is string differencePath && OfficePath.AreEquivalent(output, differencePath)) throw new PSArgumentException("A redline must be written to a distinct destination.");
                    WriteArtifact(output, temporary => WordDocumentComparer.CreateRedlineDocument(referenceWord, differenceWord, temporary,
                        new WordComparisonRedlineOptions { ComparisonOptions = settings }));
                    report.OutputPath = output;
                }
            } else if (left is ExcelDocument referenceExcel && right is ExcelDocument differenceExcel) {
                var settings = OfficeComparisonOptionsBinding.Capture<ExcelWorkbookDiffOptions>(engine, Options);
                if (MaximumDifferences.HasValue) settings.MaxDifferences = MaximumDifferences.Value;
                var native = referenceExcel.CompareWorkbook(differenceExcel, settings);
                report = new OfficeComparisonReport(native, native.AreEqual, native.Differences.Cast<object>(), MaximumDifferences);
                if (output != null && Approve(output, "Write Excel comparison workbook")) {
                    if (!System.IO.Path.GetExtension(output).Equals(".xlsx", StringComparison.OrdinalIgnoreCase)) throw new PSArgumentException("An Excel findings workbook OutputPath must end in .xlsx.");
                    if (leftPath != null && OfficePath.AreEquivalent(output, leftPath) || rightPath != null && OfficePath.AreEquivalent(output, rightPath)) throw new PSArgumentException("A comparison artifact must use a distinct destination.");
                    WriteExcel(native, output); report.OutputPath = output;
                }
            } else throw new PSNotSupportedException("Semantic comparison supports Word/Word and Excel/Excel. Use Visual for native PDF page comparison and format-rendered proofs.");
            _cancel.Token.ThrowIfCancellationRequested();
            WriteObject(report);
        } finally { for (int index = owned.Count - 1; index >= 0; index--) owned[index].Dispose(); }
    }

    private static object Load(string path, string? password, PdfLoadOptions? pdf) {
        var family = OfficeProtectionDispatch.Family(System.IO.Path.GetExtension(path));
        return family switch {
            OfficeProtectionFamily.Word => WordDocument.Load(path, new WordLoadOptions { Password = password, AccessMode = DocumentAccessMode.ReadOnly, PersistenceMode = DocumentPersistenceMode.Explicit, AllowLegacyProjection = false }),
            OfficeProtectionFamily.Excel => ExcelDocument.Load(path, new ExcelLoadOptions { Password = password, AccessMode = DocumentAccessMode.ReadOnly, PersistenceMode = DocumentPersistenceMode.Explicit, AllowBinaryProjection = false }),
            OfficeProtectionFamily.PowerPoint => PowerPointDocument.Load(path, new PowerPointLoadOptions { Password = password, AccessMode = DocumentAccessMode.ReadOnly, PersistenceMode = DocumentPersistenceMode.Explicit, AllowLegacyProjection = false }),
            _ => PdfDocument.Load(path, pdf ?? new PdfLoadOptions { Password = password })
        };
    }
    private PdfDocument Pdf(object document, List<IDisposable> owned, List<OfficeConversionReport> conversions) {
        OfficeConversionReport evidence;
        PdfDocument result;
        switch (document) {
            case PdfDocument pdf: return pdf;
            case WordDocument word: result = word.ToPdf(new WordPdfExportOptions { FailOnLoss = true }, out evidence, _cancel.Token); break;
            case ExcelDocument excel: result = excel.ToPdf(new ExcelPdfExportOptions { FailOnLoss = true }, out evidence, _cancel.Token); break;
            case PowerPointDocument ppt: result = ppt.ToPdf(new PowerPointPdfExportOptions { FailOnLoss = true }, out evidence, _cancel.Token); break;
            default: throw new PSNotSupportedException("Visual comparison requires a supported native document model.");
        }
        conversions.Add(evidence); owned.Add(result); return result;
    }
    private void WriteExcel(ExcelWorkbookDiffReport report, string output) {
        using var workbook = ExcelDocument.Create();
        var sheet = workbook.AddSheet("Differences");
        string[] headers = { "Category", "Sheet", "Address", "Reference", "Difference", "Message" };
        for (int column = 0; column < headers.Length; column++) sheet.Cell(1, column + 1).SetValue(headers[column]);
        int row = 2;
        foreach (var finding in report.Differences) {
            string?[] values = { finding.Category, finding.SheetName, finding.Address, finding.LeftValue, finding.RightValue, finding.Message };
            for (int column = 0; column < values.Length; column++) sheet.Cell(row, column + 1).SetValue(values[column] ?? string.Empty);
            row++;
        }
        WriteArtifact(output, temporary => workbook.Save(temporary));
    }
    private bool Approve(string output, string action) {
        _cancel.Token.ThrowIfCancellationRequested();
        return ShouldProcess(output, action);
    }
    private void WriteArtifact(string output, Action<string> write) => OfficeFileOutput.Write(output, overwrite: true, temporary => {
        _cancel.Token.ThrowIfCancellationRequested();
        write(temporary);
        _cancel.Token.ThrowIfCancellationRequested();
    });
    /// <inheritdoc />
    protected override void StopProcessing() => _cancel.Cancel();
    /// <inheritdoc />
    protected override void EndProcessing() => _cancel.Dispose();
    /// <summary>Releases this command-owned cancellation source.</summary>
    public void Dispose() => _cancel.Dispose();
}
