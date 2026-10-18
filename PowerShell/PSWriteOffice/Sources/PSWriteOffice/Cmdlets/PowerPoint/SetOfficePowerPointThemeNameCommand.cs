using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Sets the PowerPoint theme name.</summary>
/// <example>
///   <summary>Rename the theme across the presentation.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePowerPoint -Path .\Examples\Documents\PowerPointThemeName.pptx {
///     Set-OfficePowerPointThemeName -Name 'Service Brief' -AllMasters
///     Add-OfficePowerPointSlide -Layout 1 -PassThru | Set-OfficePowerPointSlideTitle -Title 'Named theme'
/// }</code>
///   <para>Applies a friendly theme name across every master before saving.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePowerPointThemeName")]
[Alias("PptThemeName")]
[OutputType(typeof(PowerPointPresentation))]
public sealed class SetOfficePowerPointThemeNameCommand : PSCmdlet
{
    /// <summary>Presentation to update (optional inside New-OfficePowerPoint).</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointPresentation? Presentation { get; set; }

    /// <summary>Theme name to apply.</summary>
    [Parameter(Mandatory = true)]
    public string Name { get; set; } = string.Empty;

    /// <summary>Apply the name across all slide masters.</summary>
    [Parameter]
    public SwitchParameter AllMasters { get; set; }

    /// <summary>Emit the presentation after update.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        try
        {
            var presentation = Presentation ?? PowerPointDslContext.Require(this).Presentation;
            if (AllMasters.IsPresent)
            {
                presentation.SetThemeNameForAllMasters(Name);
            }
            else
            {
                presentation.ThemeName = Name;
            }

            if (PassThru.IsPresent)
            {
                WriteObject(presentation);
            }
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "PowerPointSetThemeNameFailed", ErrorCategory.InvalidOperation, Presentation));
        }
    }
}
