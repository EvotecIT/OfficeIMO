using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using DocumentFormat.OpenXml.Presentation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Sets the text of the title placeholder on a slide.</summary>
/// <para>Targets the title placeholder when available; otherwise adds a new title shape.</para>
/// <example>
///   <summary>Rename a slide.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficePowerPointSlide -Presentation $ppt -Index 0 | Set-OfficePowerPointSlideTitle -Title 'Executive Summary'</code>
///   <para>Updates the first slide’s title.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePowerPointSlideTitle")]
[Alias("PptTitle")]
public class SetOfficePowerPointSlideTitleCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Slide whose title should change (optional inside DSL).</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointSlide? Slide { get; set; }

    /// <summary>New title text.</summary>
    [Parameter(Mandatory = true)]
    public string Title { get; set; } = null!;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try {
            var slide = Slide ?? PowerPointDslContext.Require(this).RequireSlide();
            var titleBox = slide.GetPlaceholder(PowerPointPlaceholderType.Title) ??
                           slide.GetPlaceholder(PowerPointPlaceholderType.CenteredTitle);

            if (titleBox != null) {
                titleBox.Text = Title;
            } else {
                slide.AddTitle(Title);
            }

            WritePassThru(slide);
        } catch (Exception ex) {
            WriteError(new ErrorRecord(ex, "PowerPointSetTitleFailed", ErrorCategory.InvalidOperation, Slide));
        }
    }
}