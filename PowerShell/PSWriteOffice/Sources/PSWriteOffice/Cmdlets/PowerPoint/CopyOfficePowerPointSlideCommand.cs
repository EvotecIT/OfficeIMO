using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Copies an existing slide within a PowerPoint presentation.</summary>
/// <para>Uses OfficeIMO slide duplication so charts, notes, and shapes are preserved.</para>
/// <example>
///   <summary>Duplicate the first slide and insert the copy after it.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePowerPoint -Path .\Examples\Documents\PowerPointCopySlide.pptx {
///     $slide = Add-OfficePowerPointSlide -Layout 1 -PassThru
///     Set-OfficePowerPointSlideTitle -Slide $slide -Title 'Original'
///     $copy = Copy-OfficePowerPointSlide -Index 0 -PassThru
///     Set-OfficePowerPointSlideTitle -Slide $copy -Title 'Copied appendix'
/// }</code>
///   <para>Duplicates a slide and updates the copied slide title.</para>
/// </example>
[Cmdlet(VerbsCommon.Copy, "OfficePowerPointSlide")]
[OutputType(typeof(PowerPointSlide))]
public sealed class CopyOfficePowerPointSlideCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Presentation to update (optional inside DSL).</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointPresentation? Presentation { get; set; }

    /// <summary>Zero-based slide index to duplicate.</summary>
    [Parameter(Mandatory = true)]
    public int Index { get; set; }

    /// <summary>Optional target index for the duplicate; omit to insert after the source slide.</summary>
    [Parameter]
    public int? InsertAt { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try {
            var presentation = Presentation ?? PowerPointDslContext.Current?.Presentation
                ?? throw new InvalidOperationException("Presentation was not provided. Use -Presentation or run inside New-OfficePowerPoint.");

            WritePassThru(presentation.DuplicateSlide(Index, InsertAt));
        } catch (Exception ex) {
            WriteError(new ErrorRecord(ex, "PowerPointCopySlideFailed", ErrorCategory.InvalidOperation, Presentation));
        }
    }
}
