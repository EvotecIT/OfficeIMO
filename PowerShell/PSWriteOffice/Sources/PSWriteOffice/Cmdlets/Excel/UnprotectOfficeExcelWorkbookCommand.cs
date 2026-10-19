using System.IO;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Removes workbook structure/window protection metadata.</summary>
/// <example>
///   <summary>Remove workbook-level protection metadata.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Unprotect-OfficeExcelWorkbook -Path .\Report.xlsx
/// Test-OfficeExcelWorkbook -Path .\Report.xlsx -SkipOpenXmlValidation |
///     Select-Object Passed, ProtectionSummary</code>
///   <para>Removes workbook structure/window protection metadata and saves the workbook.</para>
/// </example>
[Cmdlet(VerbsSecurity.Unprotect, "OfficeExcelWorkbook", DefaultParameterSetName = ParameterSetContext, SupportsShouldProcess = true)]
[Alias("ExcelWorkbookUnprotect")]
[OutputType(typeof(ExcelDocument), typeof(FileInfo))]
public sealed class UnprotectOfficeExcelWorkbookCommand : PSCmdlet {
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Workbook path to update.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Emit the workbook after removing protection.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var pathPassThru = string.Equals(ParameterSetName, ParameterSetPath, System.StringComparison.OrdinalIgnoreCase);
        string? resolvedPath = pathPassThru
            ? SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path)
            : null;

        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: false);

        if (!ExcelShouldProcessService.ShouldProcessWorkbook(this, workbook.Document, Path, "Update Excel workbook")) {

            return;

        }

        var document = workbook.Document;
        document.UnprotectWorkbook();
        workbook.SaveIfOwned();

        if (PassThru.IsPresent) {
            WriteObject(pathPassThru ? new FileInfo(resolvedPath!) : document);
        }
    }
}