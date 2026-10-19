using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Sets PowerPoint theme fonts.</summary>
/// <example>
///   <summary>Set theme Latin fonts.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePowerPoint -Path .\Examples\Documents\PowerPointThemeFonts.pptx {
///     Set-OfficePowerPointThemeFonts -MajorLatin 'Aptos Display' -MinorLatin 'Aptos' -AllMasters
///     Add-OfficePowerPointSlide -Layout 1 -PassThru | Set-OfficePowerPointSlideTitle -Title 'Theme fonts'
/// }</code>
///   <para>Updates theme fonts before creating slides.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePowerPointThemeFonts")]
[Alias("PptThemeFonts")]
[OutputType(typeof(PowerPointPresentation))]
public sealed class SetOfficePowerPointThemeFontsCommand : PSCmdlet
{
    /// <summary>Presentation to update (optional inside New-OfficePowerPoint).</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointPresentation? Presentation { get; set; }

    /// <summary>Major Latin font.</summary>
    [Parameter]
    public string? MajorLatin { get; set; }

    /// <summary>Minor Latin font.</summary>
    [Parameter]
    public string? MinorLatin { get; set; }

    /// <summary>Major East Asian font.</summary>
    [Parameter]
    public string? MajorEastAsian { get; set; }

    /// <summary>Minor East Asian font.</summary>
    [Parameter]
    public string? MinorEastAsian { get; set; }

    /// <summary>Major complex script font.</summary>
    [Parameter]
    public string? MajorComplexScript { get; set; }

    /// <summary>Minor complex script font.</summary>
    [Parameter]
    public string? MinorComplexScript { get; set; }

    /// <summary>Slide master index to update when not using <see cref="AllMasters"/>.</summary>
    [Parameter]
    [ValidateRange(0, int.MaxValue)]
    public int Master { get; set; }

    /// <summary>Apply the changes across all slide masters.</summary>
    [Parameter]
    public SwitchParameter AllMasters { get; set; }

    /// <summary>Clear unspecified font slots instead of keeping existing values.</summary>
    [Parameter]
    public SwitchParameter ClearMissing { get; set; }

    /// <summary>Emit the presentation after update.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        try
        {
            if (!HasRequestedChanges())
            {
                throw new PSArgumentException("Specify at least one theme font to update.");
            }

            var presentation = Presentation ?? PowerPointDslContext.Require(this).Presentation;
            var fonts = new PowerPointThemeFontSet(
                MajorLatin,
                MinorLatin,
                MajorEastAsian,
                MinorEastAsian,
                MajorComplexScript,
                MinorComplexScript);

            if (AllMasters.IsPresent)
            {
                presentation.SetThemeFontsForAllMasters(fonts, keepExistingWhenNull: !ClearMissing.IsPresent);
            }
            else
            {
                presentation.SetThemeFonts(fonts, Master, keepExistingWhenNull: !ClearMissing.IsPresent);
            }

            if (PassThru.IsPresent)
            {
                WriteObject(presentation);
            }
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "PowerPointSetThemeFontsFailed", ErrorCategory.InvalidOperation, Presentation));
        }
    }

    private bool HasRequestedChanges()
    {
        return !string.IsNullOrWhiteSpace(MajorLatin)
               || !string.IsNullOrWhiteSpace(MinorLatin)
               || !string.IsNullOrWhiteSpace(MajorEastAsian)
               || !string.IsNullOrWhiteSpace(MinorEastAsian)
               || !string.IsNullOrWhiteSpace(MajorComplexScript)
               || !string.IsNullOrWhiteSpace(MinorComplexScript);
    }
}
