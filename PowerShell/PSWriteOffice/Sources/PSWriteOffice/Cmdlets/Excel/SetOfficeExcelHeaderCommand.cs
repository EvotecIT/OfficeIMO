using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Updates worksheet header text, print flags or an image while retaining the footer.</summary>
/// <example><summary>Add a report title.</summary>
/// <code>ExcelSheet 'Data' { ExcelHeader -Center 'Quarterly review' }</code></example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelHeader", DefaultParameterSetName = "Text")]
[Alias("ExcelHeader")]
[OutputType(typeof(ExcelSheet))]
public sealed class SetOfficeExcelHeaderCommand : ExcelRunningTextCmdlet {
    /// <inheritdoc />
    protected override bool IsHeader => true;
}
