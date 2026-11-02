using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Lists native formula records, optionally filtered by worksheet, volatile functions or external references.</summary>
/// <example><summary>Find volatile formulas without opening Excel.</summary>
/// <code>Get-OfficeExcelFormula -Path .\Models\*.xlsx -Volatile | Format-Table SheetName, Address, Formula</code></example>
/// <example><summary>Inspect literal file names from the pipeline.</summary>
/// <code>Get-ChildItem -LiteralPath .\Models -Filter *.xlsx | Get-OfficeExcelFormula -External</code></example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelFormula", DefaultParameterSetName = "Document")]
[OutputType(typeof(ExcelFormulaInfo))]
public sealed class GetOfficeExcelFormulaCommand : ExcelQueryCmdlet {
    /// <summary>Return only formulas containing volatile functions.</summary>
    [Parameter]
    public SwitchParameter Volatile { get; set; }

    /// <summary>Return only formulas referring to external workbooks.</summary>
    [Parameter]
    public SwitchParameter External { get; set; }

    /// <inheritdoc />
    protected override IEnumerable<PSObject> Query(ExcelDocument document, IReadOnlyList<ExcelSheet> sheets, string? path) {
        var names = new HashSet<string>(sheets.Select(sheet => sheet.Name), StringComparer.OrdinalIgnoreCase);
        foreach (ExcelFormulaInfo formula in document.AnalyzeFormulas().Formulas) {
            if (names.Contains(formula.SheetName) && (!Volatile || formula.IsVolatile)
                && (!External || formula.HasExternalReference)) yield return CreateRecord(formula, path);
        }
    }
}
