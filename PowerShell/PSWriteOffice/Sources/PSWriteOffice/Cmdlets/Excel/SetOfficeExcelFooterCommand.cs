using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Updates worksheet footer text, print flags or an image while retaining the header.</summary>
/// <example><summary>Number printed pages without replacing the header.</summary>
/// <code>ExcelSheet 'Data' { ExcelFooter -Right 'Page &amp;P of &amp;N' }</code></example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelFooter", DefaultParameterSetName = "Text")]
[Alias("ExcelFooter")]
[OutputType(typeof(ExcelSheet))]
public sealed class SetOfficeExcelFooterCommand : ExcelRunningTextCmdlet {
    /// <inheritdoc />
    protected override bool IsHeader => false;
}
