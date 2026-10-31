using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Copies a native slide within its document, preserving its content and relationships.</summary>
/// <example><summary>Copy the first slide to the end of a two-slide deck.</summary>
/// <code>Copy-OfficePowerPointSlide -Document $deck -Slide 1 -Position 3 -PassThru</code></example>
/// <example><summary>Copy a selected native slide immediately after itself.</summary>
/// <code>$slide | Copy-OfficePowerPointSlide -PassThru</code></example>
[Cmdlet(VerbsCommon.Copy, "OfficePowerPointSlide")]
[OutputType(typeof(PowerPointSlide))]
public sealed class CopyOfficePowerPointSlideCommand : OfficeCmdlet {
    /// <summary>Gets or sets the document for numeric selection; omission uses the enclosing scope or native slide owner.</summary>
    [Parameter, ValidateNotNull]
    public PowerPointDocument? Document { get; set; }
    /// <summary>Gets or sets a native slide or its one-based position.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true), PowerPointSlideTransformation]
    public PowerPointSlideSelector Slide { get; set; } = null!;
    /// <summary>Gets or sets the copy's one-based destination position; omission inserts it immediately after the source.</summary>
    [Parameter, ValidateRange(1, int.MaxValue)]
    public int? Position { get; set; }
    /// <summary>Gets or sets whether to return the new native slide.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var document = Document ?? OfficeDocumentContext.Find<PowerPointDocument>() ?? Slide.Slide?.OwnerDocument;
        try {
            document = ResolveDocument(document);
            PowerPointSlide source = Slide.Resolve(document);
            if (Position > document.Slides.Count + 1)
                throw new ArgumentOutOfRangeException(nameof(Position), "The copy position must be within the deck or immediately after its last slide.");
            PowerPointSlide copy = document.DuplicateSlide(source.Index, Position - 1);
            if (PassThru) WriteObject(copy);
        } catch (ArgumentException error) { InvalidArgument(error, Slide); }
        catch (Exception error) when (error is not (PipelineStoppedException or ActionPreferenceStopException)) {
            ItemError(error, "powerpoint-slide-copy-failed", Slide);
        }
    }
}
