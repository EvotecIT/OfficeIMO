#pragma warning disable CS1591
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Gets workbook formula references, functions, volatile formulas, and external links.</summary>
/// <example>
///   <summary>Find volatile and external-link formulas before publishing a workbook.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$analysis = Get-OfficeExcelFormulaAnalysis -Path .\Model.xlsx -IncludeFormulas
/// $analysis.Formulas |
///     Where-Object { $_.IsVolatile -or $_.HasExternalReference } |
///     Format-Table SheetName,Address,Formula,IsVolatile,HasExternalReference</code>
///   <para>Uses OfficeIMO formula analysis so scripts can review volatile functions and external workbook references without parsing package XML themselves.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelFormulaAnalysis", DefaultParameterSetName = ParameterSetPath)]
[Alias("ExcelFormulaAnalysis", "ExcelFormulaAudit")]
[OutputType(typeof(PSObject))]
public sealed class GetOfficeExcelFormulaAnalysisCommand : PSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Workbook path.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook document.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Include per-cell formula details.</summary>
    [Parameter]
    public SwitchParameter IncludeFormulas { get; set; }

    protected override void ProcessRecord() {
        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: true);
        var report = workbook.Document.AnalyzeFormulas();
        var output = new PSObject();
        output.Properties.Add(new PSNoteProperty("Path", workbook.Document.FilePath));
        output.Properties.Add(new PSNoteProperty("FormulaCount", report.FormulaCount));
        output.Properties.Add(new PSNoteProperty("VolatileFormulaCount", report.VolatileFormulaCount));
        output.Properties.Add(new PSNoteProperty("ExternalReferenceCount", report.ExternalReferenceCount));
        if (IncludeFormulas.IsPresent) {
            output.Properties.Add(new PSNoteProperty("Formulas", report.Formulas.Select(CreateFormula).ToArray()));
        }

        WriteObject(output);
    }

    private static PSObject CreateFormula(ExcelFormulaInfo formula) {
        var item = new PSObject();
        item.Properties.Add(new PSNoteProperty("SheetName", formula.SheetName));
        item.Properties.Add(new PSNoteProperty("Address", formula.Address));
        item.Properties.Add(new PSNoteProperty("Formula", formula.Formula));
        item.Properties.Add(new PSNoteProperty("References", formula.References.ToArray()));
        item.Properties.Add(new PSNoteProperty("Functions", formula.Functions.ToArray()));
        item.Properties.Add(new PSNoteProperty("HasExternalReference", formula.HasExternalReference));
        item.Properties.Add(new PSNoteProperty("IsVolatile", formula.IsVolatile));
        return item;
    }
}
#pragma warning restore CS1591