using System;
using System.IO;
using System.Management.Automation;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using PSWriteOffice.Services.Excel;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Converts an Excel workbook to an HTML review document.</summary>
/// <example>
///   <summary>Export a workbook as semantic HTML.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ConvertTo-OfficeExcelHtml -Path .\Report.xlsx -OutputPath .\Report.html -Title 'Workbook Review' -PassThru</code>
///   <para>Loads the workbook and writes an HTML file with tables, formulas, comments, charts, and image inventory where available.</para>
/// </example>
/// <example>
///   <summary>Export a visual review artifact.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ConvertTo-OfficeExcelHtml -Path .\Report.xlsx -Profile VisualReview -OutputPath .\Report.visual.html</code>
///   <para>Uses the OfficeIMO Excel visual review profile.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficeExcelHtml", DefaultParameterSetName = ParameterSetPath, SupportsShouldProcess = true)]
[Alias("ConvertTo-ExcelHtml")]
[OutputType(typeof(string), typeof(FileInfo))]
public sealed class ConvertToOfficeExcelHtmlCommand : PSCmdlet
{
    private const string ParameterSetPath = "Path";
    private const string ParameterSetWorkbook = "Workbook";

    /// <summary>Path to the workbook to convert.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = ParameterSetPath)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook instance to convert.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetWorkbook)]
    public ExcelDocument Workbook { get; set; } = null!;

    /// <summary>Password used to open an encrypted workbook package.</summary>
    [Parameter(ParameterSetName = ParameterSetPath)]
    public string? Password { get; set; }

    /// <summary>Optional output HTML path. When omitted, HTML text is returned.</summary>
    [Parameter]
    [Alias("OutPath")]
    public string? OutputPath { get; set; }

    /// <summary>HTML conversion profile.</summary>
    [Parameter]
    public OfficeExcelHtmlProfile Profile { get; set; } = OfficeExcelHtmlProfile.SemanticTables;

    /// <summary>Built-in HTML document theme.</summary>
    [Parameter]
    public OfficeVisualThemeKind Theme { get; set; } = OfficeVisualThemeKind.WordLike;

    /// <summary>Optional HTML document title.</summary>
    [Parameter]
    public string? Title { get; set; }

    /// <summary>Maximum number of used rows to emit per worksheet.</summary>
    [Parameter]
    public int? MaxRowsPerSheet { get; set; }

    /// <summary>Text used for empty cells in semantic table output.</summary>
    [Parameter]
    public string EmptyCellText { get; set; } = string.Empty;

    /// <summary>Do not include OfficeIMO default CSS styles.</summary>
    [Parameter]
    public SwitchParameter NoDefaultStyles { get; set; }

    /// <summary>Emit a FileInfo when saving to disk.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        ExcelDocument? document = null;
        var dispose = false;
        try
        {
            document = ResolveWorkbook(out dispose);
            var html = document.ToHtml(CreateOptions());
            WriteHtml(html);
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "ConvertToOfficeExcelHtmlFailed", ErrorCategory.InvalidOperation,
                ParameterSetName == ParameterSetPath ? Path : Workbook));
        }
        finally
        {
            if (dispose)
            {
                ExcelDocumentService.CloseDocument(document!);
            }
        }
    }

    private ExcelDocument ResolveWorkbook(out bool dispose)
    {
        dispose = false;
        if (ParameterSetName == ParameterSetWorkbook)
        {
            return Workbook ?? throw new PSArgumentNullException(nameof(Workbook));
        }

        dispose = true;
        return ExcelDocumentService.LoadDocument(PdfCommandUtilities.ResolvePath(this, Path), readOnly: true, autoSave: false, Password);
    }

    private ExcelHtmlSaveOptions CreateOptions()
    {
        var options = new ExcelHtmlSaveOptions
        {
            Profile = Profile == OfficeExcelHtmlProfile.VisualReview
                ? OfficeHtmlConversionProfile.ExcelVisualReview
                : OfficeHtmlConversionProfile.ExcelSemanticTables,
            Theme = Theme,
            IncludeDefaultStyles = !NoDefaultStyles.IsPresent,
            EmptyCellText = EmptyCellText ?? string.Empty
        };

        if (MaxRowsPerSheet.HasValue)
        {
            options.MaxRowsPerSheet = MaxRowsPerSheet.Value;
        }

        if (!string.IsNullOrWhiteSpace(Title))
        {
            options.Title = Title!;
        }

        return options;
    }

    private void WriteHtml(string html)
    {
        if (string.IsNullOrWhiteSpace(OutputPath))
        {
            WriteObject(html);
            return;
        }

        var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath!);
        if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Write Excel HTML"))
        {
            return;
        }

        PdfCommandUtilities.EnsureDirectory(outputPath);
        File.WriteAllText(outputPath, html, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        if (PassThru.IsPresent)
        {
            WriteObject(new FileInfo(outputPath));
        }
    }
}

/// <summary>Excel HTML conversion profiles exposed by PSWriteOffice.</summary>
public enum OfficeExcelHtmlProfile
{
    /// <summary>Emit semantic workbook tables and inventories.</summary>
    SemanticTables,

    /// <summary>Emit a visual review artifact.</summary>
    VisualReview
}
