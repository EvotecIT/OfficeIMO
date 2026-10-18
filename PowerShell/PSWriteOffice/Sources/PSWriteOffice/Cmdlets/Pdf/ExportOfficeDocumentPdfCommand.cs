using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;
using PSWriteOffice.Services.Pdf;
using PSWriteOffice.Services.PowerPoint;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Exports Word, Excel, PowerPoint, HTML, Markdown, RTF, and literal text documents to PDF.</summary>
/// <para>Accepts live OfficeIMO documents, individual files, directories, and selected file pipelines. Directory and file batches require PowerShell 7.4 or newer.</para>
/// <example>
///   <summary>Export a live Word document.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$document | Export-OfficeDocumentPdf -Path .\Report.pdf</code>
/// </example>
/// <example>
///   <summary>Export a supported file without opening it explicitly.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Export-OfficeDocumentPdf -InputPath .\Report.docx -Path .\Report.pdf -PassThru</code>
/// </example>
/// <example>
///   <summary>Configure Markdown PDF export with ordinary PowerShell parameters.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeMarkdownPdfOptions -Title 'Service report' -IncludeLocalImages -BaseDirectory .\Assets
/// Export-OfficeDocumentPdf -InputPath .\Report.md -Path .\Report.pdf -MarkdownOptions $options</code>
///   <para>Typed conversion options also apply to the matching formats in a mixed batch.</para>
/// </example>
/// <example>
///   <summary>Export a directory with optional checkpoints.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Export-OfficeDocumentPdf -InputDirectory .\Documents -OutputDirectory .\PDF -CheckpointDirectory .\PDF-State</code>
/// </example>
/// <example>
///   <summary>Stream selected file outcomes and retain the final summary.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-ChildItem .\Documents -File | Export-OfficeDocumentPdf -OutputDirectory .\PDF -ItemResults -SummaryVariable summary</code>
/// </example>
[Cmdlet(VerbsData.Export, "OfficeDocumentPdf", DefaultParameterSetName = ParameterSetDocument, SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
#if !FRAMEWORK
[OutputType(typeof(OfficeIMO.Workflows.OfficeConversionBatchResult), typeof(OfficeIMO.Workflows.OfficeConversionBatchItemResult))]
#endif
public sealed partial class ExportOfficeDocumentPdfCommand : AsyncPSCmdlet {
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Live Word, Excel, PowerPoint, HTML, Markdown, or RTF document to export. Saved FileInfo and path strings from the pipeline are opened automatically.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0, ParameterSetName = ParameterSetDocument)]
    public object Document { get; set; } = null!;

    /// <summary>Source .doc, .docx, .txt, .xlsx, .pptx, .html, .htm, .xhtml, .md, .markdown, or .rtf file.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("SourcePath", "FullName")]
    public string InputPath { get; set; } = string.Empty;

    /// <summary>Destination PDF path.</summary>
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = ParameterSetDocument)]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = ParameterSetPath)]
    [Alias("OutputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Password used to open an encrypted Word, Excel, or PowerPoint source file.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>Word-specific PDF options.</summary>
    [Parameter]
    public WordToPdfOptions? WordOptions { get; set; }

    /// <summary>Excel-specific PDF options.</summary>
    [Parameter]
    public ExcelToPdfOptions? ExcelOptions { get; set; }

    /// <summary>PowerPoint-specific PDF options.</summary>
    [Parameter]
    public PowerPointToPdfOptions? PowerPointOptions { get; set; }

    /// <summary>Markdown-specific PDF options.</summary>
    [Parameter]
    public MarkdownToPdfOptions? MarkdownOptions { get; set; }

    /// <summary>RTF-specific PDF options.</summary>
    [Parameter]
    public RtfToPdfOptions? RtfOptions { get; set; }

    /// <summary>HTML-specific PDF rendering options.</summary>
    [Parameter]
    public HtmlToPdfOptions? HtmlOptions { get; set; }

    /// <summary>Literal text layout and strict decoding options. Applies only to TXT sources.</summary>
    [Parameter]
    public PdfPlainTextOptions? TextOptions { get; set; }

    /// <summary>Accept reported legacy DOC import loss. Known loss otherwise blocks conversion.</summary>
    [Parameter]
    public SwitchParameter AllowLegacyImportLoss { get; set; }

    /// <summary>Maximum source bytes for DOC and TXT import.</summary>
    [Parameter]
    [ValidateRange(1, long.MaxValue)]
    public long MaximumInputBytes { get; set; } = 64L * 1024 * 1024;

    /// <summary>Variable receiving source-stage reports separately from PDF rendering diagnostics.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    [Parameter(ParameterSetName = ParameterSetPath)]
    public string? SourceConversionReportVariable { get; set; }

    /// <summary>Variable name that receives structured PDF conversion warnings.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    [Parameter(ParameterSetName = ParameterSetPath)]
    public string? PdfWarningVariable { get; set; }

    /// <summary>Variable name that receives the structured PDF conversion report.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    [Parameter(ParameterSetName = ParameterSetPath)]
    public string? PdfConversionReportVariable { get; set; }

    /// <summary>Open the PDF after exporting it.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    [Parameter(ParameterSetName = ParameterSetPath)]
    [Alias("Show")]
    public SwitchParameter Open { get; set; }

    /// <summary>Emit the saved PDF file.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    [Parameter(ParameterSetName = ParameterSetPath)]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync() {
        if (ParameterSetName == ParameterSetDirectory) { await ExportBatchAsync(null); return; }
        if (ParameterSetName == ParameterSetFiles) {
            foreach (object selected in InputPaths) {
                object value = UnwrapDocument(selected);
                string input = value is FileInfo file ? file.FullName : value is string path ? path
                    : throw new PSArgumentException("InputPaths accepts source path strings or FileInfo objects.");
                if (_batchInputs.Count >= MaximumFiles) throw new PSArgumentException("Selected files exceed MaximumFiles.");
                _batchInputs.Add(PdfCommandUtilities.ResolvePath(this, input));
            }
            return;
        }
        var outputPath = PdfCommandUtilities.ResolvePath(this, Path);
        if (!string.Equals(System.IO.Path.GetExtension(outputPath), ".pdf", StringComparison.OrdinalIgnoreCase))
            throw new PSArgumentException("The destination must use the .pdf extension.", nameof(Path));
        if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Export document to PDF")) {
            return;
        }

        PdfCommandUtilities.EnsureDirectory(outputPath);
        object document;
        Action? closeOwnedDocument = null;
        string? sourcePath = null;

        try {
            if (ParameterSetName == ParameterSetPath) {
                document = LoadDocument(InputPath, out closeOwnedDocument, out sourcePath);
            } else {
                document = UnwrapDocument(Document);
                if (document is FileInfo file) {
                    document = LoadDocument(file.FullName, out closeOwnedDocument, out sourcePath);
                } else if (document is string path) {
                    document = LoadDocument(path, out closeOwnedDocument, out sourcePath);
                }
            }

            PdfSaveResult result = SaveDocument(document, outputPath, sourcePath);
            PdfCommandUtilities.SetVariable(this, PdfWarningVariable, result.Warnings);
            PdfCommandUtilities.SetVariable(this, PdfConversionReportVariable, result.Report);
            PdfCommandUtilities.SetVariable(this, SourceConversionReportVariable, result.ConversionReports.Take(result.ConversionReports.Count - 1).ToArray());
            foreach (var report in result.ConversionReports.Take(result.ConversionReports.Count - 1))
                foreach (var finding in report.FidelityDiagnostics)
                    WriteWarning(finding.Code + " [" + finding.Source + "]: " + finding.Message);
            result.RequireSuccess();
        } finally {
            closeOwnedDocument?.Invoke();
        }

        if (Open.IsPresent) {
            FileOpenService.Open(outputPath);
        }

        if (PassThru.IsPresent) {
            WriteObject(new FileInfo(outputPath));
        }
    }

    private object LoadDocument(string inputPath, out Action? closeOwnedDocument, out string sourcePath) {
        sourcePath = PdfCommandUtilities.ResolveExistingFilePath(this, inputPath);
        string extension = System.IO.Path.GetExtension(sourcePath).ToLowerInvariant();
        if (TextOptions != null && extension != ".txt") throw new PSArgumentException("TextOptions requires a TXT source.");
        if (AllowLegacyImportLoss.IsPresent && extension != ".doc") throw new PSArgumentException("AllowLegacyImportLoss requires a DOC source.");
        switch (System.IO.Path.GetExtension(sourcePath).ToLowerInvariant()) {
            case ".doc":
            case ".txt": {
                closeOwnedDocument = null;
                using var source = File.OpenRead(sourcePath);
                return extension == ".doc"
                    ? LegacyDocPdfConverter.ToPdfDocumentResult(source, WordOptions,
                        new OfficeIMO.Word.LegacyDoc.LegacyDocImportOptions { MaxInputBytes = (int)Math.Min(int.MaxValue, MaximumInputBytes) },
                        AllowLegacyImportLoss.IsPresent ? OfficeIMO.OfficeConversionLossPolicy.Allow : OfficeIMO.OfficeConversionLossPolicy.Block, CancelToken)
                    : PdfPlainTextConverter.ToPdfDocumentResult(source, TextOptions, MaximumInputBytes, CancelToken);
            }
            case ".docx": {
                    var document = WordDocumentService.LoadDocument(sourcePath, readOnly: true, autoSave: false, Password);
                    closeOwnedDocument = () => WordDocumentService.CloseDocument(document);
                    return document;
                }
            case ".xlsx": {
                    var document = ExcelDocumentService.LoadDocument(sourcePath, readOnly: true, autoSave: false, Password);
                    closeOwnedDocument = () => ExcelDocumentService.CloseDocument(document);
                    return document;
                }
            case ".pptx": {
                    var document = PowerPointDocumentService.LoadPresentation(sourcePath, Password, readOnly: true);
                    closeOwnedDocument = () => PowerPointDocumentService.ClosePresentation(document, save: false, show: false);
                    return document;
                }
            case ".md":
            case ".markdown":
                closeOwnedDocument = null;
                return MarkdownDoc.Load(sourcePath);
            case ".rtf":
                closeOwnedDocument = null;
                return RtfDocument.Load(sourcePath);
            case ".html":
            case ".htm":
            case ".xhtml":
                closeOwnedDocument = null;
                return HtmlConversionDocument.Load(sourcePath);
            default:
                throw new PSArgumentException("Supported PDF source extensions are .doc, .docx, .txt, .xlsx, .pptx, .html, .htm, .xhtml, .md, .markdown, and .rtf.", nameof(InputPath));
        }
    }

    private PdfSaveResult SaveDocument(object document, string outputPath, string? sourcePath) {
        switch (document) {
            case PdfDocumentConversionResult converted:
                return converted.SaveResult(outputPath, CancelToken);
            case WordDocument word:
                return word.SaveAsPdf(outputPath, WordOptions ?? new WordToPdfOptions(), CancelToken);
            case ExcelDocument excel:
                return excel.SaveAsPdf(outputPath, ExcelOptions ?? new ExcelToPdfOptions(), CancelToken);
            case PowerPointPresentation powerPoint:
                return powerPoint.SaveAsPdf(outputPath, PowerPointOptions ?? new PowerPointToPdfOptions(), CancelToken);
            case MarkdownDoc markdown:
                return markdown.SaveAsPdf(outputPath, PrepareMarkdownOptions(sourcePath), CancelToken);
            case RtfDocument rtf:
                return rtf.SaveAsPdf(outputPath, RtfOptions ?? new RtfToPdfOptions(), CancelToken);
            case HtmlConversionDocument html:
                return html.ToPdfDocumentResult(HtmlOptions, CancelToken).SaveResult(outputPath, CancelToken);
            default:
                throw new PSArgumentException(
                    $"Document type '{document?.GetType().FullName ?? "<null>"}' cannot be exported to PDF. Use a WordDocument, ExcelDocument, PowerPointPresentation, HtmlConversionDocument, MarkdownDoc, or RtfDocument.",
                    nameof(Document));
        }
    }

    private MarkdownToPdfOptions PrepareMarkdownOptions(string? sourcePath) {
        var options = MarkdownOptions?.Clone() ?? new MarkdownToPdfOptions();
        if (options.ResourcePolicy.AllowLocalFileAccess &&
            string.IsNullOrWhiteSpace(options.BaseDirectory) &&
            !string.IsNullOrWhiteSpace(sourcePath)) {
            options.BaseDirectory = System.IO.Path.GetDirectoryName(sourcePath);
        }

        return options;
    }

    private static object UnwrapDocument(object document) {
        while (document is PSObject psObject && psObject.BaseObject != document) {
            document = psObject.BaseObject;
        }

        return document;
    }

    /// <inheritdoc />
    protected override async Task EndProcessingAsync() {
        if (ParameterSetName == ParameterSetFiles && _batchInputs.Count > 0) await ExportBatchAsync(_batchInputs.ToArray());
    }
}
