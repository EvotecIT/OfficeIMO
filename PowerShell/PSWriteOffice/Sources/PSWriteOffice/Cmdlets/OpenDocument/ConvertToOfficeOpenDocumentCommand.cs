using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.OpenDocument;
using OfficeIMO.Word;
using OfficeIMO.Word.OpenDocument;
using PSWriteOffice.Services.Excel;
using PSWriteOffice.Services.PowerPoint;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.OpenDocument;

/// <summary>Converts Word, Excel, or PowerPoint content to native OpenDocument with fidelity evidence.</summary>
/// <example>
///   <summary>Convert Word to ODT and reject lossy conversion.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeWordOpenDocumentOptions -IncludeImages -IncludeHeadersAndFooters
/// ConvertTo-OfficeOpenDocument -Path .\Report.docx -OutputPath .\Report.odt -WordOptions $options -FailOnLoss</code>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficeOpenDocument", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(OdfConversionResult<OdtDocument>), typeof(OdfConversionResult<OdsDocument>), typeof(OdfConversionResult<OdpPresentation>))]
public sealed class ConvertToOfficeOpenDocumentCommand : PSCmdlet
{
    /// <summary>Path to a DOCX, XLSX, or PPTX file.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Path")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Open Word document.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Word")]
    public WordDocument WordDocument { get; set; } = null!;

    /// <summary>Open Excel workbook.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Excel")]
    public ExcelDocument ExcelDocument { get; set; } = null!;

    /// <summary>Open PowerPoint presentation.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "PowerPoint")]
    public PowerPointPresentation PowerPointPresentation { get; set; } = null!;

    /// <summary>Destination ODT, ODS, or ODP path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Optional Word-to-ODT conversion settings.</summary>
    [Parameter]
    public WordOpenDocumentConversionOptions? WordOptions { get; set; }

    /// <summary>Optional Excel-to-ODS conversion settings.</summary>
    [Parameter]
    public ExcelOpenDocumentConversionOptions? ExcelOptions { get; set; }

    /// <summary>Optional PowerPoint-to-ODP conversion settings.</summary>
    [Parameter]
    public PowerPointOpenDocumentConversionOptions? PowerPointOptions { get; set; }

    /// <summary>Throw when the conversion approximates, skips, or cannot map a feature.</summary>
    [Parameter]
    public SwitchParameter FailOnLoss { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var output = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath);
        var kind = ResolveKind(out var sourcePath);
        OpenDocumentCommandUtilities.ValidateOpenDocumentExtension(output, kind, nameof(OutputPath));
        if (!ShouldProcess(output, "Convert Office document to OpenDocument")) return;
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(output) ?? SessionState.Path.CurrentFileSystemLocation.Path);
        WordDocument? ownedWord = null;
        ExcelDocument? ownedExcel = null;
        PowerPointPresentation? ownedPresentation = null;
        try
        {
            object result;
            switch (kind)
            {
                case OdfDocumentKind.Text:
                    var word = WordDocument;
                    if (sourcePath != null)
                    {
                        word = WordDocumentService.LoadDocument(sourcePath, readOnly: true, autoSave: false);
                        ownedWord = word;
                    }
                    var wordResult = word.ToOpenDocumentResult(WordOptions);
                    if (FailOnLoss.IsPresent) wordResult.RequireNoLoss();
                    wordResult.Value.Save(output);
                    result = wordResult;
                    break;
                case OdfDocumentKind.Spreadsheet:
                    var excel = ExcelDocument;
                    if (sourcePath != null)
                    {
                        excel = ExcelDocumentService.LoadDocument(sourcePath, readOnly: true, autoSave: false);
                        ownedExcel = excel;
                    }
                    var excelResult = excel.ToOpenDocumentResult(ExcelOptions);
                    if (FailOnLoss.IsPresent) excelResult.RequireNoLoss();
                    excelResult.Value.Save(output);
                    result = excelResult;
                    break;
                case OdfDocumentKind.Presentation:
                    var presentation = PowerPointPresentation;
                    if (sourcePath != null)
                    {
                        presentation = PowerPointDocumentService.LoadPresentation(sourcePath, readOnly: true);
                        ownedPresentation = presentation;
                    }
                    var presentationResult = presentation.ToOpenDocumentResult(PowerPointOptions);
                    if (FailOnLoss.IsPresent) presentationResult.RequireNoLoss();
                    presentationResult.Value.Save(output);
                    result = presentationResult;
                    break;
                default:
                    throw new InvalidOperationException("Unsupported OpenDocument conversion kind.");
            }
            WriteObject(result);
        }
        finally
        {
            if (ownedWord != null) WordDocumentService.CloseDocument(ownedWord);
            if (ownedExcel != null) ExcelDocumentService.CloseDocument(ownedExcel);
            if (ownedPresentation != null) PowerPointDocumentService.ClosePresentation(ownedPresentation, save: false, show: false);
        }
    }

    private OdfDocumentKind ResolveKind(out string? sourcePath)
    {
        sourcePath = null;
        if (ParameterSetName == "Word") return OdfDocumentKind.Text;
        if (ParameterSetName == "Excel") return OdfDocumentKind.Spreadsheet;
        if (ParameterSetName == "PowerPoint") return OdfDocumentKind.Presentation;
        sourcePath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        return System.IO.Path.GetExtension(sourcePath).ToLowerInvariant() switch
        {
            ".docx" => OdfDocumentKind.Text,
            ".xlsx" or ".xlsm" => OdfDocumentKind.Spreadsheet,
            ".pptx" or ".pptm" => OdfDocumentKind.Presentation,
            _ => throw new PSArgumentException("Path must identify a DOCX, XLSX, XLSM, PPTX, or PPTM file.")
        };
    }

}
