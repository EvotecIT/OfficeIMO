using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Updates worksheet footer text, print flags or an image while retaining the header.</summary>
/// <example><summary>Number printed pages without replacing the header.</summary>
/// <code>ExcelSheet 'Data' { ExcelFooter -Right 'Page &amp;P of &amp;N' }</code></example>
/// <example>
/// <summary>Number the printed pages</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Q3-Sales.xlsx {
///     ExcelSheet 'Orders' {
///         ExcelFooter -Right 'Page &amp;P of &amp;N'
///     }
/// }</code>
/// <para>Prints Page N of M on the right of the footer.</para>
/// </example>
/// <example>
/// <summary>Add a classification footer to every sheet</summary>
/// <code>New-OfficeExcel -Path .\Reports\Security-Review.xlsx {
///     foreach ($name in 'Findings', 'Accounts', 'Firewall') {
///         ExcelSheet $name {
///             ExcelTable (Import-OfficeCsv -Path ".\Data\$name.csv")
///             ExcelFooter -Left 'Internal' -Center '&amp;F' -Right 'Page &amp;P of &amp;N'
///         }
///     }
/// }</code>
/// <para>Prints the classification, file name and page number in the footer of each sheet of a new report.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelFooter", DefaultParameterSetName = "Text")]
[Alias("ExcelFooter")]
[OutputType(typeof(ExcelSheet))]
public sealed class SetOfficeExcelFooterCommand : ExcelRunningTextCmdlet {
    /// <inheritdoc />
    protected override bool IsHeader => false;
}
