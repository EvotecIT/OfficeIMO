using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Exports workbook sheets as PNG or SVG images with one result per sheet.</summary>
/// <example>
///   <summary>Export visible sheets as PNG files.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Export-OfficeExcelImage -Path .\Report.xlsx -OutputPath .\Images</code>
///   <para>Writes one image per selected sheet. Add <c>-PassThru</c> to receive the structured export results.</para>
/// </example>
[Cmdlet(VerbsData.Export, "OfficeExcelImage", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(OfficeImageExportResult))]
public sealed class ExportOfficeExcelImageCommand : PSCmdlet
{
    /// <summary>Path to the workbook.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Path")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Open workbook instance.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Document")]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Destination folder.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Output image format.</summary>
    [Parameter]
    public OfficeImageExportFormat Format { get; set; } = OfficeImageExportFormat.Png;

    /// <summary>Optional sheet selection, range, size, and rendering settings.</summary>
    [Parameter]
    public ExcelWorkbookImageExportOptions? Options { get; set; }

    /// <summary>Emit one structured image export result per saved sheet.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var output = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath);
        if (!ShouldProcess(output, $"Export Excel sheets as {Format}")) return;
        Directory.CreateDirectory(output);
        ExcelDocument? owned = null;
        try
        {
            var document = Document;
            if (ParameterSetName == "Path")
            {
                var input = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
                owned = ExcelDocumentService.LoadDocument(input, readOnly: true, autoSave: false);
                document = owned;
            }
            IReadOnlyList<OfficeImageExportResult> results = document.SaveAsImages(output, Format, Options);
            if (PassThru.IsPresent) WriteObject(results, enumerateCollection: true);
        }
        finally
        {
            owned?.Dispose();
        }
    }
}
