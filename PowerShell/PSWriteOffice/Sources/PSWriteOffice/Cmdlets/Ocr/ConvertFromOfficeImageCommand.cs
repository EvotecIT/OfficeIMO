using System.IO;
using System.Management.Automation;
using System.Threading.Tasks;
using OfficeIMO.CSV;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Word.Pdf;
using PSWriteOffice.Services.Ocr;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Ocr;

/// <summary>Exports a recognized image as editable Excel, Word, text, Markdown, HTML, JSON, CSV or searchable PDF.</summary>
/// <para>Recognize once with Get-OfficeImageDocument, inspect its evidence, then reuse that snapshot for each output.
/// Excel and CSV contain detected tables only. CSV requires exactly one table. Empty recognition or missing tables fail before output publication.</para>
/// <example>
/// <summary>Export an image as Excel and Word.</summary>
/// <prefix>PS&gt; </prefix>
/// <code>$image = Get-OfficeImageDocument -Path .\Ledger.png
/// $image | ConvertFrom-OfficeImage -OutputPath .\Ledger.xlsx -FirstRowIsHeader -PassThru
/// $image | ConvertFrom-OfficeImage -OutputPath .\Ledger.docx</code>
/// <para>Excel receives caller-confirmed headers and typed columns. PassThru returns recognition evidence and the format conversion report.</para>
/// </example>
/// <example>
/// <summary>Customize CSV formatting while retaining formula escaping.</summary>
/// <prefix>PS&gt; </prefix>
/// <code>$csv = @{ Delimiter = ';'; FormulaInjectionPolicy = 'Escape' }
/// $image | ConvertFrom-OfficeImage -OutputPath .\Ledger.csv -CsvOptions $csv</code>
/// <para>Set Escape explicitly when customizing CSV formatting for spreadsheet consumption. Every recognized table row is written as data without a synthesized header, regardless of IncludeHeader.</para>
/// </example>
[Cmdlet(VerbsData.ConvertFrom, "OfficeImage", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
[OutputType(typeof(OfficeImageConversionResult))]
public sealed class ConvertFromOfficeImageCommand : AsyncPSCmdlet
{
    /// <summary>Recognized image review returned by Get-OfficeImageDocument.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [ValidateNotNull]
    public PdfSearchableOcrReview InputObject { get; set; } = null!;

    /// <summary>Output path. The extension selects .xlsx, .docx, .pdf, .txt, .md, .html, .json or .csv.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    [ValidateNotNullOrEmpty]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Confirm the first row of each Excel table with unknown schema as column headers.</summary>
    [Parameter]
    public SwitchParameter FirstRowIsHeader { get; set; }

    /// <summary>Advanced Excel typing, culture, row limit and table formatting settings.</summary>
    [Parameter]
    public PdfTablesToExcelOptions? ExcelOptions { get; set; }

    /// <summary>Advanced Word projection settings.</summary>
    [Parameter]
    public PdfToWordOptions? WordOptions { get; set; }

    /// <summary>Advanced HTML projection settings. Word and HTML defaults exclude the original scan from editable content.</summary>
    [Parameter]
    public PdfToHtmlOptions? HtmlOptions { get; set; }

    /// <para>CSV field formatting and formula policy. When omitted, formula-like source text is escaped, including negative numeric text. Supplied options retain their FormulaInjectionPolicy; a new CsvSaveOptions object defaults to Preserve. Set Escape explicitly when customizing options for spreadsheet consumption.</para>
    /// <para>Every recognized table row is written as data without a synthesized header, regardless of IncludeHeader. This command controls UTF-8, uncompressed file output and replacement through Force; CsvOptions controls field formatting and formula handling.</para>
    [Parameter]
    public CsvSaveOptions? CsvOptions { get; set; }

    /// <summary>Replace an existing destination after conversion succeeds.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }

    /// <summary>Return output location, source review and format report instead of FileInfo.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override Task ProcessRecordAsync()
    {
        string extension = System.IO.Path.GetExtension(OutputPath).ToLowerInvariant();
        OfficeImageExporter.ValidateExtension(extension);
        if (extension != ".xlsx" && (FirstRowIsHeader.IsPresent || ExcelOptions != null))
            throw new PSArgumentException("Excel settings require an .xlsx output.");
        if (extension != ".docx" && WordOptions != null)
            throw new PSArgumentException("Word settings require a .docx output.");
        if (extension != ".html" && HtmlOptions != null)
            throw new PSArgumentException("HTML settings require an .html output.");
        if (extension != ".csv" && CsvOptions != null)
            throw new PSArgumentException("CSV settings require a .csv output.");
        string outputPath = PdfCommandUtilities.ResolveOutputFilePath(this, OutputPath, extension, Force.IsPresent);
        if (!ShouldProcess(outputPath, "Export recognized image")) return Task.CompletedTask;
        var excel = ExcelOptions?.Clone() ?? new PdfTablesToExcelOptions();
        if (MyInvocation.BoundParameters.ContainsKey(nameof(FirstRowIsHeader))) excel.UseFirstRowAsHeader = FirstRowIsHeader.IsPresent;
        var result = OfficeImageExporter.Export(InputObject, outputPath, Force.IsPresent, excel, WordOptions, CancelToken, HtmlOptions, CsvOptions);
        foreach (var page in InputObject.Ocr.Pages)
        {
            if (page.RejectedLowConfidenceCount > 0)
                WriteWarning($"Page {page.PageNumber}: excluded {page.RejectedLowConfidenceCount} low-confidence words. Inspect InputObject.Ocr.Pages before using the output.");
            foreach (string diagnostic in page.Diagnostics) WriteWarning(diagnostic);
        }
        if (result.ConversionReport != null)
            foreach (var diagnostic in result.ConversionReport.FidelityDiagnostics) WriteWarning(diagnostic.Message);
        if (extension == ".csv" && result.TableScope?.HasOmittedPageContent == true)
            WriteWarning("CSV contains the detected table. Content outside that table is recorded in the PassThru TableScope report.");
        WriteObject(PassThru.IsPresent ? result : (object)result.File);
        return Task.CompletedTask;
    }
}
