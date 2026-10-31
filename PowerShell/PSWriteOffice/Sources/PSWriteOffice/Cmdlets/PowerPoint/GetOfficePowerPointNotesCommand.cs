using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Reads speaker notes without creating empty notes parts.</summary>
/// <example><summary>Read the presenter script.</summary>
/// <code>Get-OfficePowerPointNotes -Document $deck -IncludeEmpty | Select-Object SlideIndex, Text</code></example>
/// <example><summary>Read notes from a selected slide.</summary>
/// <code>Get-OfficePowerPointNotes -Document $deck -Slide 3</code></example>
[Cmdlet(VerbsCommon.Get, "OfficePowerPointNotes")]
[OutputType(typeof(PowerPointNotesInfo))]
public sealed class GetOfficePowerPointNotesCommand : OfficeCmdlet {
    /// <summary>Gets or sets a live document; omission uses the enclosing document block.</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointDocument? Document { get; set; }
    /// <summary>Gets or sets a native slide object or one-based slide position.</summary>
    [Parameter, PowerPointSlideTransformation]
    public PowerPointSlideSelector? Slide { get; set; }
    /// <summary>Gets or sets whether slides without notes are included.</summary>
    [Parameter]
    public SwitchParameter IncludeEmpty { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var document = Document ?? OfficeDocumentContext.Find<PowerPointDocument>();
        try {
            if (Slide != null) {
                PowerPointSlide slide = Slide.Resolve(document);
                WriteInfo(slide, slide.Index);
            } else {
                document = ResolveDocument(document);
                for (int i = 0; i < document.Slides.Count; i++) WriteInfo(document.Slides[i], i);
            }
        } catch (ArgumentException error) { InvalidArgument(error, Slide); }
    }
    private void WriteInfo(PowerPointSlide slide, int index) {
        var info = PowerPointNotesReader.Read(slide, index);
        if (IncludeEmpty || info.HasNotes || !string.IsNullOrWhiteSpace(info.Text)) WriteObject(info);
    }
}
