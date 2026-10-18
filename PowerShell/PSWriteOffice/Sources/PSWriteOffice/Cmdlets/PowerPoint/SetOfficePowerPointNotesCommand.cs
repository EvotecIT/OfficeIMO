using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Sets speaker notes for a PowerPoint slide.</summary>
/// <example>
///   <summary>Attach speaker notes.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePowerPoint -Path .\Examples\Documents\PowerPointNotes.pptx {
///     $slide = Add-OfficePowerPointSlide -Layout 1 -PassThru
///     Set-OfficePowerPointSlideTitle -Slide $slide -Title 'Executive summary'
///     Set-OfficePowerPointNotes -Slide $slide -Text 'Keep this slide under five minutes and focus on decisions.'
/// }</code>
///   <para>Writes speaker notes to a generated slide.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePowerPointNotes")]
[Alias("PptNotes")]
[OutputType(typeof(PowerPointSlide))]
public sealed class SetOfficePowerPointNotesCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Slide whose notes should be updated (optional inside DSL).</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointSlide? Slide { get; set; }

    /// <summary>Notes text to apply.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Text { get; set; } = string.Empty;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try {
            var slide = Slide ?? PowerPointDslContext.Require(this).RequireSlide();
            var notes = slide.Notes;
            notes.Text = Text ?? string.Empty;
            WritePassThru(slide);
        } catch (Exception ex) {
            WriteError(new ErrorRecord(ex, "PowerPointSetNotesFailed", ErrorCategory.InvalidOperation, Slide));
        }
    }
}
