using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Gets native data validation snapshots, with worksheet and source-path metadata.</summary>
/// <example><summary>Inspect rules for a named table column.</summary>
/// <code>Get-OfficeExcelValidation -Path .\Report.xlsx -Sheet Data -Table Sales -Header Status</code></example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelValidation", DefaultParameterSetName = "Document")]
[OutputType(typeof(ExcelDataValidationInfo))]
public sealed class GetOfficeExcelValidationCommand : ExcelQueryCmdlet {
    /// <summary>A1 range filter; Header takes precedence when both are supplied.</summary>
    [Parameter]
    public string? Range { get; set; }
    /// <summary>Header or table column name used to resolve the range filter.</summary>
    [Parameter]
    public string? Header { get; set; }
    /// <summary>Table name for header-based range resolution.</summary>
    [Parameter]
    public string? Table { get; set; }
    /// <summary>Worksheet header row; zero uses the first row of the used range.</summary>
    [Parameter, ValidateRange(0, 1048576)]
    public int HeaderRow { get; set; }
    /// <summary>Include the header cell when resolving a named column.</summary>
    [Parameter]
    public SwitchParameter IncludeHeader { get; set; }

    /// <inheritdoc />
    protected override IEnumerable<PSObject> Query(ExcelDocument document, IReadOnlyList<ExcelSheet> sheets, string? path) {
        foreach (ExcelSheet sheet in sheets) {
            string? range = ExcelTargetRangeResolver.ResolveOptional(sheet, string.IsNullOrWhiteSpace(Header) ? Range : null, Header, Table, HeaderRow, IncludeHeader);
            foreach (ExcelDataValidationInfo rule in sheet.GetDataValidations(range)) yield return CreateRecord(rule, path, sheet.Name);
        }
    }
}
