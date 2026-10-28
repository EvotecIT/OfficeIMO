using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Sets a formula in a worksheet cell.</summary>
/// <para>Supports A1 addresses or row/column coordinates in the Excel DSL.</para>
/// <example>
///   <summary>Write a SUM formula.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Set-OfficeExcelFormula -Address 'C2' -Formula 'SUM(A2:B2)' }</code>
///   <para>Stores the formula in cell C2.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelFormula")]
[Alias("ExcelFormula")]
public sealed class SetOfficeExcelFormulaCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>1-based row index.</summary>
    [Parameter(ParameterSetName = "Coordinates")]
    public int? Row { get; set; }

    /// <summary>1-based column index.</summary>
    [Parameter(ParameterSetName = "Coordinates")]
    public int? Column { get; set; }

    /// <summary>A1-style cell address (e.g., A1, C5).</summary>
    [Parameter(ParameterSetName = "Address")]
    public string? Address { get; set; }

    /// <summary>Formula text (without leading =).</summary>
    [Parameter(Mandatory = true)]
    public string Formula { get; set; } = string.Empty;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var context = ExcelDslContext.Require(this);
        var sheet = context.RequireSheet();
        var (row, column) = ExcelHostExtensions.ResolveCellAddress(Row, Column, Address);
        sheet.Cell(row, column, value: null, formula: Formula, numberFormat: null);
        WritePassThru(sheet);
    }
}