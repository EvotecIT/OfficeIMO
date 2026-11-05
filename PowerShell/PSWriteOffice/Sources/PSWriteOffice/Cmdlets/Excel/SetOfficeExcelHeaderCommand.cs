using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Updates worksheet header text, print flags or an image while retaining the footer.</summary>
/// <example><summary>Add a report title.</summary>
/// <code>ExcelSheet 'Data' { ExcelHeader -Center 'Quarterly review' }</code></example>
/// <example>
/// <summary>Add a page header</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Q3-Sales.xlsx {
///     ExcelSheet 'Orders' {
///         ExcelHeader -Center 'Q3 orders' -Right '&amp;D'
///     }
/// }</code>
/// <para>Prints the report name in the centre and the print date on the right of every page.</para>
/// </example>
/// <example>
/// <summary>Put the logo in the header</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Q3-Sales.xlsx {
///     ExcelSheet 'Summary' {
///         ExcelHeader -ImagePath .\Assets\Logo.png -ImagePosition Left -ImageHeight 1cm -Right '&amp;A'
///     }
/// }</code>
/// <para>Prints the company logo at the left of the header and the sheet name on the right.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelHeader", DefaultParameterSetName = "Text")]
[Alias("ExcelHeader")]
[OutputType(typeof(ExcelSheet))]
public sealed class SetOfficeExcelHeaderCommand : ExcelRunningTextCmdlet {
    /// <inheritdoc />
    protected override bool IsHeader => true;
}
