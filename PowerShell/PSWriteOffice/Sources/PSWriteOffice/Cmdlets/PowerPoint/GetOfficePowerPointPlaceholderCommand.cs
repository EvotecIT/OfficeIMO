using System.Linq;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Gets native slide placeholders or metadata from the selected slide's template layout.</summary>
/// <example><summary>Read the first slide's title placeholder.</summary>
/// <code>Get-OfficePowerPointPlaceholder -Document $deck -Slide 1 -Type Title</code></example>
/// <example><summary>Inspect template positions before filling a slide.</summary>
/// <code>$slide | Get-OfficePowerPointPlaceholder -FromLayout</code></example>
[Cmdlet(VerbsCommon.Get, "OfficePowerPointPlaceholder", DefaultParameterSetName = "Slide")]
[OutputType(typeof(PowerPointTextBox), typeof(PowerPointLayoutPlaceholderInfo))]
public sealed class GetOfficePowerPointPlaceholderCommand : OfficeCmdlet {
    /// <summary>Gets or sets a native slide or one-based position; omission uses the current slide block.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = "Slide"),
        Parameter(ValueFromPipeline = true, ParameterSetName = "Layout"), PowerPointSlideTransformation]
    public PowerPointSlideSelector? Slide { get; set; }
    /// <summary>Gets or sets the document for numeric selection or explicit membership checks.</summary>
    [Parameter, ValidateNotNull]
    public PowerPointDocument? Document { get; set; }
    /// <summary>Gets or sets whether to return layout metadata instead of live text boxes.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Layout")]
    public SwitchParameter FromLayout { get; set; }
    /// <summary>Gets or sets the placeholder type to include.</summary>
    [Parameter]
    public PowerPointPlaceholderType? Type { get; set; }
    /// <summary>Gets or sets the native placeholder index to include.</summary>
    [Parameter]
    public uint? Index { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try {
            var slide = PowerPointManualPlacement.Resolve(Slide, Document).Slide;
            if (FromLayout) {
                var placeholders = slide.GetLayoutPlaceholders().Where(item =>
                    (!Type.HasValue || item.PlaceholderType == Type) &&
                    (!Index.HasValue || item.PlaceholderIndex == Index));
                WriteObject(placeholders, enumerateCollection: true);
            } else {
                var placeholders = slide.Placeholders.Where(item =>
                    (!Type.HasValue || item.PlaceholderType == Type) &&
                    (!Index.HasValue || item.PlaceholderIndex == Index));
                WriteObject(placeholders, enumerateCollection: true);
            }
        } catch (System.ArgumentException error) { InvalidArgument(error, Slide); }
        catch (System.Exception error) when (error is not (PipelineStoppedException or ActionPreferenceStopException)) {
            ItemError(error, "powerpoint-placeholder-read-failed", Slide);
        }
    }
}
