using System.Management.Automation;
using OfficeIMO.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Lists built-in Excel table styles and compatibility recommendations.</summary>
/// <example>
/// <summary>List the recommended table styles</summary>
/// <code>Get-OfficeExcelTableStyle -RecommendedOnly</code>
/// <para>Shows the built-in table styles that render well in every supported host.</para>
/// </example>
/// <example>
/// <summary>Pick styles that render outside desktop Excel</summary>
/// <code>Get-OfficeExcelTableStyle -Profile CrossHost -RecommendedOnly | Format-Table</code>
/// <para>Lists the styles recommended for workbooks also opened in web and mobile spreadsheet apps.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelTableStyle", DefaultParameterSetName = "Default")]
public sealed class GetOfficeExcelTableStyleCommand : PSCmdlet
{
    /// <summary>Compatibility profile used to evaluate table styles.</summary>
    [Parameter(ParameterSetName = "Default")]
    public ExcelTableStyleCompatibilityProfile Profile { get; set; } = ExcelTableStyleCompatibilityProfile.CrossHost;

    /// <summary>Return only styles recommended for the selected profile.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter RecommendedOnly { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        foreach (var name in ExcelTableStyleCatalog.GetNames())
        {
            var info = ExcelTableStyleCatalog.Analyze(name, Profile);
            if (RecommendedOnly.IsPresent && !info.IsRecommended)
            {
                continue;
            }

            WriteObject(info);
        }
    }
}
