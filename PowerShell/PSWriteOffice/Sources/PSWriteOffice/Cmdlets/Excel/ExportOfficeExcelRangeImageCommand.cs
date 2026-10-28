using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Exports one worksheet range as an image file.</summary>
/// <example>
///   <summary>Export a report range for use in an email.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Export-OfficeExcelRangeImage -Path .\Report.xlsx -WorksheetName Summary -Range A1:H20 -OutputPath .\summary.png</code>
/// </example>
[Cmdlet(VerbsData.Export, "OfficeExcelRangeImage", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(OfficeImageExportResult))]
public sealed class ExportOfficeExcelRangeImageCommand : PSCmdlet
{
    /// <summary>Path to the workbook.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Path")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Open workbook instance.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Document")]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Name of the worksheet containing the range.</summary>
    [Parameter(Mandatory = true)]
    public string WorksheetName { get; set; } = string.Empty;

    /// <summary>A1 range address to export.</summary>
    [Parameter(Mandatory = true)]
    public string Range { get; set; } = string.Empty;

    /// <summary>Optional destination image file. When omitted, returns the in-memory image result only.</summary>
    [Parameter(Position = 1)]
    public string? OutputPath { get; set; }

    /// <summary>Output image format.</summary>
    [Parameter]
    public OfficeImageExportFormat Format { get; set; } = OfficeImageExportFormat.Png;

    /// <summary>Optional rendering, size, font, and diagnostic policy settings.</summary>
    [Parameter]
    public ExcelImageExportOptions? Options { get; set; }

    /// <summary>Replace an existing destination file.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }

    /// <summary>Emit the structured image export result when a destination path is used.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        string? output = null;
        if (!string.IsNullOrWhiteSpace(OutputPath))
        {
            output = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath);
            if (!ShouldProcess(output, $"Save Excel range {WorksheetName}!{Range} as {Format}")) return;
        }

        ExcelDocument? owned = null;
        try
        {
            ExcelDocument document = Document;
            if (ParameterSetName == "Path")
            {
                string input = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
                owned = ExcelDocumentService.LoadDocument(input, readOnly: true, autoSave: false);
                document = owned;
            }

            OfficeImageExportResult result = document[WorksheetName]
                .Range(Range)
                .ExportImage(Format, Options);

            if (output != null)
            {
                string? directory = System.IO.Path.GetDirectoryName(output);
                if (!string.IsNullOrWhiteSpace(directory))
                {
                    System.IO.Directory.CreateDirectory(directory);
                }

                result = result.Save(output, Force.IsPresent
                    ? OfficeImageExportFileConflictPolicy.Replace
                    : OfficeImageExportFileConflictPolicy.FailIfExists);
            }
            if (output == null || PassThru.IsPresent) WriteObject(result);
        }
        finally
        {
            owned?.Dispose();
        }
    }
}
