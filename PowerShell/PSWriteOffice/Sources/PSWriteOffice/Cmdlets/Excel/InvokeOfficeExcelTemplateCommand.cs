using System.Collections;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Applies Excel template markers such as {{Name}} to one or more worksheets.</summary>
/// <example>
///   <summary>Fill a workbook template from a hashtable.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Invoke-OfficeExcelTemplate -Path .\Invoice.xlsx -Sheet Invoice -Value @{ Number = 'INV-001'; Total = 123.45 } -CultureName en-US</code>
///   <para>Replaces matching template markers and saves the workbook.</para>
/// </example>
[Cmdlet(VerbsLifecycle.Invoke, "OfficeExcelTemplate", SupportsShouldProcess = true, ConfirmImpact = ConfirmImpact.Low, DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelTemplate", "ExcelTemplateApply")]
[OutputType(typeof(int))]
public sealed class InvokeOfficeExcelTemplateCommand : PSCmdlet {
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

    /// <summary>Worksheet name to update. Defaults to the current DSL sheet or all workbook sheets.</summary>
    [Parameter]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index (0-based) to update. Defaults to the current DSL sheet or all workbook sheets.</summary>
    [Parameter]
    public int? SheetIndex { get; set; }

    /// <summary>Template marker values keyed by marker name.</summary>
    [Parameter(Mandatory = true)]
    [Alias("Values")]
    public Hashtable Value { get; set; } = new();

    /// <summary>Culture name used for built-in marker format aliases such as currency and date.</summary>
    [Parameter]
    public string? CultureName { get; set; }

    /// <summary>Behavior used when a marker is not supplied by -Value.</summary>
    [Parameter]
    public ExcelTemplateMissingValueBehavior? MissingValueBehavior { get; set; }

    /// <summary>Throws when a marker is not supplied by -Value.</summary>
    [Parameter]
    public SwitchParameter ThrowOnMissing { get; set; }

    /// <summary>Returns the number of marker replacements.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var values = ExcelTemplateValueService.ConvertValues(Value);
        var options = ExcelTemplateValueService.CreateOptions(CultureName, MissingValueBehavior, ThrowOnMissing.IsPresent);
        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: false);
        var replacements = 0;
        var processedAnySheet = false;

        foreach (var sheet in ExcelWorkbookCommandService.ResolveSheets(this, workbook.Document, ParameterSetName, Sheet, SheetIndex)) {
            if (!ShouldProcess(sheet.Name, "Apply Excel template markers")) {
                continue;
            }

            replacements += sheet.ApplyTemplate(values, options);
            processedAnySheet = true;
        }

        if (processedAnySheet) {
            workbook.SaveIfOwned();
        }
        if (PassThru.IsPresent) {
            WriteObject(replacements);
        }
    }
}