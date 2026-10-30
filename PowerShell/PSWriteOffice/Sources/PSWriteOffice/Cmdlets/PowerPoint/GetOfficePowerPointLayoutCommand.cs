using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Lists template layouts from all masters, or one selected master.</summary>
/// <example><summary>List layouts in a live deck.</summary>
/// <code>Get-OfficePowerPointLayout -Document $deck</code></example>
/// <example><summary>Reuse a returned layout in a designed slide.</summary>
/// <code>New-OfficePowerPoint -Path ./Review.pptx {
///     $layout = @(Get-OfficePowerPointLayout -Master 1)[0]
///     PowerPointSlide -Title 'Review' -TemplateLayout $layout
/// }</code></example>
[Cmdlet(VerbsCommon.Get, "OfficePowerPointLayout")]
[OutputType(typeof(PowerPointSlideLayoutInfo))]
public sealed class GetOfficePowerPointLayoutCommand : OfficeCmdlet {
    /// <summary>Gets or sets the live document; omission uses the enclosing content scope.</summary>
    [Parameter(Position = 0, ValueFromPipeline = true)]
    public PowerPointPresentation? Document { get; set; }
    /// <summary>Gets or sets a one-based master position; omission lists all masters.</summary>
    [Parameter, ValidateRange(1, int.MaxValue)]
    public int? Master { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var document = ResolveDocument(Document);
        try {
            if (Master.HasValue) {
                foreach (var layout in document.GetSlideLayouts(Master.Value - 1)) WriteObject(layout);
            } else {
                foreach (var master in PowerPointTemplate.Inspect(document).Masters)
                    foreach (var layout in document.GetSlideLayouts(master.MasterIndex)) WriteObject(layout);
            }
        } catch (ArgumentException error) { InvalidArgument(error, document); }
    }
}
