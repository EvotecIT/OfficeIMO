using System.Collections;
using System.Globalization;
using System.Management.Automation;
using System.Security;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Tests whether supplied values bind every template marker.</summary>
/// <example>
/// <summary>Check that values fill a template</summary>
/// <code>Test-OfficeExcelTemplate -Path .\Templates\Service-Report.xltx -Value @{ Client = 'Northwind Traders'; Month = 'October 2026'; Availability = 0.9994 } -Quiet</code>
/// <para>Returns $true when the values bind every marker in the template.</para>
/// </example>
/// <example>
/// <summary>Stop a report run when the template changed</summary>
/// <code>$values = @{ Client = 'Northwind Traders'; Month = 'October 2026'; Availability = 0.9994 }
/// $report = Test-OfficeExcelTemplate -Path .\Templates\Service-Report.xltx -Value $values
/// $report.ToMarkdown() | Set-Content -Path .\Logs\Template-Binding.md
/// if (-not (Test-OfficeExcelTemplate -Path .\Templates\Service-Report.xltx -Value $values -Quiet)) {
///     throw 'The service report template has markers this script does not fill. See Logs\Template-Binding.md.'
/// }</code>
/// <para>Writes a Markdown binding report for review and stops before generating the monthly reports if any marker is unbound.</para>
/// </example>
[Cmdlet(VerbsDiagnostic.Test, "OfficeExcelTemplate", DefaultParameterSetName = "Document")]
public sealed class TestOfficeExcelTemplateCommand : PSCmdlet {
    /// <summary>Workbook paths resolved with filesystem wildcard support.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Value or values consumed by the selected operation.</summary>
    [Parameter(Mandatory = true)]
    public Hashtable Value { get; set; } = null!;

    /// <summary>Returns a Boolean result instead of the native validation report.</summary>
    [Parameter()]
    public SwitchParameter Quiet { get; set; }

    /// <summary>Literal workbook paths, including FileInfo.FullName and provider PSPath.</summary>
    [Parameter(ParameterSetName = "LiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Alias("PSPath", "FullName")]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();

    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ParameterSetName = "Document", ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        foreach (var scope in ExcelCommandTargets.Workbooks(this, ParameterSetName, Path, LiteralPath, Document, readOnly: true)) {
            var report = scope.Document.ValidateTemplateBindings(ExcelTemplateValueService.ConvertValues(Value));
            WriteObject(Quiet ? (object)report.Passed : report);
        }
    }

}
