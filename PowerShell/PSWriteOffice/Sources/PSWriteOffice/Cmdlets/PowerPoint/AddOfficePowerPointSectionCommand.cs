using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Starts a named section at an existing, current or next appended slide.</summary>
/// <example><summary>Group the following slides into a section.</summary>
/// <code>New-OfficePowerPoint -Path ./Results.pptx {
///     PowerPointSection 'Results'
///     PowerPointSlide -Title 'Revenue' { PowerPointParagraph 'Revenue increased.' }
///     PowerPointSlide -Title 'Margin' { PowerPointParagraph 'Margin improved.' }
/// }</code></example>
/// <example><summary>Start an appendix at an existing slide.</summary>
/// <code>Add-OfficePowerPointSection -Document $deck -Name Appendix -Slide 9</code></example>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointSection", DefaultParameterSetName = "Name")]
[Alias("PowerPointSection")]
[OutputType(typeof(PowerPointSectionInfo))]
public sealed class AddOfficePowerPointSectionCommand : OfficeCmdlet {
    /// <summary>Gets or sets the document; omission uses the enclosing scope or selected native slide's owner.</summary>
    [Parameter(ValueFromPipeline = true), ValidateNotNull]
    public PowerPointDocument? Document { get; set; }
    /// <summary>Gets or sets the section name.</summary>
    [Parameter(Mandatory = true, Position = 0), ValidateNotNullOrEmpty]
    public string Name { get; set; } = string.Empty;
    /// <summary>Gets or sets a native slide or one-based slide position; omission uses the current or next slide.</summary>
    [Parameter, PowerPointSlideTransformation]
    public PowerPointSlideSelector? Slide { get; set; }
    /// <summary>Gets or sets whether to return section metadata, including its stable native identifier.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var document = Document ?? OfficeDocumentContext.Find<PowerPointDocument>() ?? Slide?.Slide?.OwnerDocument;
        try {
            document = ResolveDocument(document);
            if (string.IsNullOrWhiteSpace(Name)) throw new ArgumentException("Section name cannot be empty.", nameof(Name));
            PowerPointSlide? slide = Slide?.Resolve(document);
            var scope = PowerPointContentScope.Current;
            if (slide == null && ReferenceEquals(scope?.Document, document)) slide = scope!.NativeSlide;
            PowerPointSectionInfo section = slide == null
                ? document.StartSectionAtNextSlide(Name)
                : document.AddSection(Name, slide.Index);
            if (PassThru) WriteObject(section);
        } catch (ArgumentException error) { InvalidArgument(error, Name); }
        catch (Exception error) when (error is not (PipelineStoppedException or ActionPreferenceStopException)) {
            ItemError(error, "powerpoint-section-create-failed", Name);
        }
    }
}
