using System;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a section to a PowerPoint presentation.</summary>
/// <para>Creates a new section starting at the requested slide index or at the current slide inside the DSL.</para>
/// <example>
///   <summary>Create a section that starts at slide 3.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePowerPoint -Path .\Examples\Documents\PowerPointSections.pptx {
///     Add-OfficePowerPointSlide -Layout 1 -PassThru | Set-OfficePowerPointSlideTitle -Title 'Overview'
///     Add-OfficePowerPointSlide -Layout 1 -PassThru | Set-OfficePowerPointSlideTitle -Title 'Results'
///     Add-OfficePowerPointSection -Name 'Results' -StartSlideIndex 1
/// }</code>
///   <para>Creates a section named Results starting at the second slide.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointSection")]
[Alias("PptSection")]
[OutputType(typeof(PowerPointSectionInfo))]
public sealed class AddOfficePowerPointSectionCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Presentation to update (optional inside DSL).</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointPresentation? Presentation { get; set; }

    /// <summary>Name of the section to add.</summary>
    [Parameter(Mandatory = true)]
    public string Name { get; set; } = string.Empty;

    /// <summary>Zero-based slide index where the section should start.</summary>
    [Parameter]
    public int? StartSlideIndex { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try {
            var context = PowerPointDslContext.Current;
            var presentation = Presentation ?? context?.Presentation
                ?? throw new InvalidOperationException("Presentation was not provided. Use -Presentation or run inside New-OfficePowerPoint.");

            int startIndex = ResolveStartSlideIndex(presentation, context);
            WritePassThru(presentation.AddSection(Name, startIndex));
        } catch (Exception ex) {
            WriteError(new ErrorRecord(ex, "PowerPointAddSectionFailed", ErrorCategory.InvalidOperation, Presentation));
        }
    }

    private int ResolveStartSlideIndex(PowerPointPresentation presentation, PowerPointDslContext? context) {
        if (StartSlideIndex.HasValue) {
            return StartSlideIndex.Value;
        }

        var currentSlide = context?.CurrentSlide;
        if (currentSlide != null) {
            int index = presentation.Slides.ToList().IndexOf(currentSlide);
            if (index >= 0) {
                return index;
            }
        }

        throw new PSArgumentException("Specify -StartSlideIndex or run inside a slide scope within New-OfficePowerPoint.");
    }
}
