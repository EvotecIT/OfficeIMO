#pragma warning disable CS1591
using System.Collections;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Validates Excel template markers against supplied bindings before applying a template.</summary>
/// <example>
///   <summary>Fail fast when a template workbook is missing required values.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$bindings = @{
///     CustomerName = 'Northwind'
///     InvoiceDate  = Get-Date
///     Total        = 1250.75
/// }
/// $result = Test-OfficeExcelTemplateBinding -Path .\InvoiceTemplate.xlsx -Binding $bindings -AsMarkdown
/// $result | Set-Content .\InvoiceTemplateBinding.md</code>
///   <para>Uses OfficeIMO template inspection and returns either structured missing-marker data or the reusable Markdown marker report.</para>
/// </example>
[Cmdlet(VerbsDiagnostic.Test, "OfficeExcelTemplateBinding", DefaultParameterSetName = ParameterSetPath)]
[Alias("ExcelTemplateBinding", "ExcelTemplateValidate")]
[OutputType(typeof(PSObject), typeof(string), typeof(bool))]
public sealed class TestOfficeExcelTemplateBindingCommand : PSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Workbook path.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook document.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Template bindings keyed by marker name.</summary>
    [Parameter(Mandatory = true)]
    public IDictionary Binding { get; set; } = null!;

    /// <summary>Return only a Boolean pass/fail value.</summary>
    [Parameter]
    public SwitchParameter Quiet { get; set; }

    /// <summary>Return OfficeIMO's Markdown marker report.</summary>
    [Parameter]
    public SwitchParameter AsMarkdown { get; set; }

    /// <summary>Throw when any marker is missing.</summary>
    [Parameter]
    public SwitchParameter ThrowOnMissing { get; set; }

    protected override void ProcessRecord() {
        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: true);
        var bindings = Binding.Cast<DictionaryEntry>().ToDictionary<DictionaryEntry, string, object?>(entry => entry.Key.ToString() ?? string.Empty, entry => entry.Value, System.StringComparer.OrdinalIgnoreCase);
        var report = workbook.Document.ValidateTemplateBindings(bindings);
        if (ThrowOnMissing.IsPresent) {
            report.Inspection.EnsureAllMarkersBound();
        }
        if (Quiet.IsPresent) {
            WriteObject(report.Passed);
            return;
        }
        if (AsMarkdown.IsPresent) {
            WriteObject(report.Markdown);
            return;
        }

        var output = new PSObject();
        output.Properties.Add(new PSNoteProperty("Path", workbook.Document.FilePath));
        output.Properties.Add(new PSNoteProperty("Passed", report.Passed));
        output.Properties.Add(new PSNoteProperty("TotalMarkers", report.TotalMarkers));
        output.Properties.Add(new PSNoteProperty("MissingMarkerNames", report.MissingMarkerNames.ToArray()));
        output.Properties.Add(new PSNoteProperty("Markers", report.Inspection.Markers.ToArray()));
        WriteObject(output);
    }
}
#pragma warning restore CS1591