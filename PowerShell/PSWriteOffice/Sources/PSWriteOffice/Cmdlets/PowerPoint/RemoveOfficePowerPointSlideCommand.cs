using System;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Removes selected native slides or one-based slide positions.</summary>
/// <example><summary>Remove two positions without shifts changing the selection.</summary>
/// <code>Remove-OfficePowerPointSlide -Document $deck -Slide 1, 3 -Confirm:$false</code></example>
/// <example><summary>Preview removal of a native slide.</summary>
/// <code>$slide | Remove-OfficePowerPointSlide -WhatIf</code></example>
[Cmdlet(VerbsCommon.Remove, "OfficePowerPointSlide", SupportsShouldProcess = true)]
[OutputType(typeof(PowerPointPresentation))]
public sealed class RemoveOfficePowerPointSlideCommand : OfficeCmdlet {
    /// <summary>Return the live document after at least one slide is removed.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }
    /// <summary>Gets or sets the document for numeric selection; omission uses the enclosing scope or native slide owner.</summary>
    [Parameter, ValidateNotNull]
    public PowerPointPresentation? Document { get; set; }
    /// <summary>Gets or sets native slides or their one-based positions; each argument array is resolved before removal.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true), PowerPointSlideTransformation, ValidateNotNullOrEmpty]
    public PowerPointSlideSelector[] Slide { get; set; } = Array.Empty<PowerPointSlideSelector>();

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var document = Document ?? OfficeDocumentContext.Find<PowerPointPresentation>()
            ?? Slide.Select(selection => selection.Slide?.OwnerPresentation).FirstOrDefault(owner => owner != null);
        try {
            document = ResolveDocument(document);
            // Resolve the complete selection before mutation, then remove from the end to preserve numeric positions.
            var selected = Slide.Select(selection => selection.Resolve(document)).Distinct()
                .OrderByDescending(slide => slide.Index).ToArray();
            bool changed = false;
            foreach (PowerPointSlide slide in selected) {
                int index = slide.Index;
                if (ShouldProcess("Slide " + (index + 1), "Remove slide")) {
                    document.RemoveSlide(index);
                    changed = true;
                }
            }
            if (changed && PassThru.IsPresent) WriteObject(document);
        } catch (ArgumentException error) { InvalidArgument(error, Slide); }
        catch (Exception error) when (error is not (PipelineStoppedException or ActionPreferenceStopException)) {
            ItemError(error, "powerpoint-slide-remove-failed", Slide);
        }
    }
}
