using System;
using System.Management.Automation;
using OfficeIMO.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Lists built-in Excel table styles and compatibility recommendations.</summary>
/// <example>
///   <summary>Show table styles recommended for cross-host workbooks.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeExcelTableStyle -RecommendedOnly |
///     Sort-Object Name |
///     Format-Table Name, Profile</code>
///   <para>Uses OfficeIMO's table style catalog to return styles that are broadly stable across desktop, web, and spreadsheet viewers.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelTableStyle")]
[Alias("ExcelTableStyle")]
public sealed class GetOfficeExcelTableStyleCommand : PSCmdlet
{
    /// <summary>Compatibility profile used to evaluate table styles.</summary>
    [Parameter]
    public ExcelTableStyleCompatibilityProfile Profile { get; set; } = ExcelTableStyleCompatibilityProfile.CrossHost;

    /// <summary>Return only styles recommended for the selected profile.</summary>
    [Parameter]
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

            var record = new PSObject();
            record.Properties.Add(new PSNoteProperty("Name", info.Name));
            record.Properties.Add(new PSNoteProperty("Style", info.Style));
            record.Properties.Add(new PSNoteProperty("Profile", info.Profile));
            record.Properties.Add(new PSNoteProperty("IsBuiltIn", info.IsBuiltIn));
            record.Properties.Add(new PSNoteProperty("IsRecommended", info.IsRecommended));
            record.Properties.Add(new PSNoteProperty("Warning", info.Warning));
            WriteObject(record);
        }
    }
}
