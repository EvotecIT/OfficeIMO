using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Sets speaker notes on a native slide or in the current slide content block.</summary>
/// <example><summary>Add talking points without placing them on the visible slide.</summary>
/// <code>New-OfficePowerPoint -Path ./Review.pptx {
///     PowerPointSlide -Title 'Service review' {
///         PowerPointParagraph 'All regions healthy'
///         PowerPointNotes 'Pause for questions before discussing capacity.'
///     }
/// }</code></example>
/// <example><summary>Update an existing slide's notes.</summary>
/// <code>Get-OfficePowerPointSlide -Document $deck -Slide 2 |
///     Set-OfficePowerPointNotes -Text 'Explain the margin before revenue.'
/// Save-OfficePowerPoint -Document $deck</code></example>
[Cmdlet(VerbsCommon.Set, "OfficePowerPointNotes")]
[Alias("PowerPointNotes")]
[OutputType(typeof(PowerPointSlide))]
public sealed class SetOfficePowerPointNotesCommand : OfficeCmdlet {
    /// <summary>Gets or sets a native slide or one-based position; omission uses the current slide content block.</summary>
    [Parameter(ValueFromPipeline = true), PowerPointSlideTransformation]
    public PowerPointSlideSelector? Slide { get; set; }

    /// <summary>Gets or sets the document for numeric selection; omission uses the enclosing document block.</summary>
    [Parameter, ValidateNotNull]
    public PowerPointPresentation? Document { get; set; }

    /// <summary>Gets or sets the notes text; an empty string clears the notes.</summary>
    [Parameter(Mandatory = true, Position = 0), AllowEmptyString]
    public string Text { get; set; } = string.Empty;

    /// <summary>Gets or sets whether to return the selected native slide after editing.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var document = Document ?? OfficeDocumentContext.Find<PowerPointPresentation>();
        try {
            if (Slide == null) {
                var content = PowerPointContentScope.Current;
                if (content == null || !ReferenceEquals(content.Document, document))
                    throw new ArgumentException("Supply -Slide, pipe a native slide, or invoke PowerPointNotes inside PowerPointSlide.");
                content.NativeSlide.Notes.Text = Text;
                if (PassThru) WriteObject(content.NativeSlide);
                return;
            }
            PowerPointSlide slide = Slide.Resolve(document);
            slide.Notes.Text = Text;
            if (PassThru) WriteObject(slide);
        } catch (ArgumentException error) { InvalidArgument(error, Slide); }
        catch (Exception error) when (error is not (PipelineStoppedException or ActionPreferenceStopException)) {
            ItemError(error, "powerpoint-notes-edit-failed", Slide);
        }
    }
}
