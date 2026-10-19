using System.IO;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Protects workbook structure or windows metadata. This is not file encryption.</summary>
/// <example>
///   <summary>Prevent worksheet add/delete/move/rename operations in Excel UI.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Protect-OfficeExcelWorkbook -Path .\Report.xlsx -Password secret
/// Test-OfficeExcelWorkbook -Path .\Report.xlsx -SkipOpenXmlValidation |
///     Select-Object Passed, ProtectionSummary</code>
///   <para>Writes workbook-level structure protection metadata and saves the workbook.</para>
/// </example>
[Cmdlet(VerbsSecurity.Protect, "OfficeExcelWorkbook", DefaultParameterSetName = ParameterSetContext, SupportsShouldProcess = true)]
[Alias("ExcelWorkbookProtect")]
[OutputType(typeof(ExcelDocument), typeof(FileInfo))]
public sealed class ProtectOfficeExcelWorkbookCommand : PSCmdlet {
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

    /// <summary>Do not protect workbook structure.</summary>
    [Parameter]
    public SwitchParameter NoStructure { get; set; }

    /// <summary>Protect workbook windows where supported by the consuming application.</summary>
    [Parameter]
    public SwitchParameter ProtectWindows { get; set; }

    /// <summary>Optional workbook protection password. This is UI protection, not package encryption.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>Optional precomputed legacy workbook protection hash to write as-is.</summary>
    [Parameter]
    public string? LegacyPasswordHash { get; set; }

    /// <summary>Emit the workbook after protection.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var protectStructure = !NoStructure.IsPresent;
        if (!protectStructure && !ProtectWindows.IsPresent) {
            throw new PSArgumentException("Use -ProtectWindows when -NoStructure is specified.");
        }

        var pathPassThru = string.Equals(ParameterSetName, ParameterSetPath, System.StringComparison.OrdinalIgnoreCase);
        string? resolvedPath = pathPassThru
            ? SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path)
            : null;

        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: false);

        if (!ExcelShouldProcessService.ShouldProcessWorkbook(this, workbook.Document, Path, "Update Excel workbook")) {

            return;

        }

        var document = workbook.Document;
        document.ProtectWorkbook(new ExcelWorkbookProtectionOptions {
            ProtectStructure = protectStructure,
            ProtectWindows = ProtectWindows.IsPresent,
            Password = Password,
            LegacyPasswordHash = LegacyPasswordHash
        });

        workbook.SaveIfOwned();

        if (PassThru.IsPresent) {
            WriteObject(pathPassThru ? new FileInfo(resolvedPath!) : document);
        }
    }
}