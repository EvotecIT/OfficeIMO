using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Sets the workbook date system used for numeric date serials.</summary>
/// <example>
///   <summary>Switch a workbook to the 1904 date system.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$workbook | Set-OfficeExcelDateSystem -DateSystem 1904 -PassThru | Save-OfficeExcel</code>
///   <para>Marks the workbook to use Excel's 1904 date system before saving.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelDateSystem")]
[Alias("ExcelDateSystem")]
[OutputType(typeof(ExcelDocument))]
public sealed class SetOfficeExcelDateSystemCommand : PSCmdlet
{
    /// <summary>Workbook to update.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Date system to use for Excel date serials.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    [ValidateSet("1900", "1904", "NineteenHundred", "NineteenFour")]
    public string DateSystem { get; set; } = "1900";

    /// <summary>Emit the workbook for further pipeline operations.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Document == null)
        {
            return;
        }

        Document.DateSystem = ExcelDateSystemService.Resolve(DateSystem, nameof(DateSystem));
        if (PassThru.IsPresent)
        {
            WriteObject(Document);
        }
    }
}
