using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Changes the layout used by a slide.</summary>
/// <example>
///   <summary>Switch a slide to a layout by name.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficePowerPointSlide -Presentation $ppt -Index 0 | Set-OfficePowerPointSlideLayout -LayoutName 'Title and Content'</code>
///   <para>Updates the slide to use the requested layout.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePowerPointSlideLayout", DefaultParameterSetName = ParameterSetIndex)]
[Alias("PptSlideLayout")]
[OutputType(typeof(PowerPointSlide))]
public sealed class SetOfficePowerPointSlideLayoutCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    private const string ParameterSetIndex = "Index";
    private const string ParameterSetByName = "Name";
    private const string ParameterSetByType = "Type";

    /// <summary>Slide to update (optional inside a slide DSL scope).</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointSlide? Slide { get; set; }

    /// <summary>Slide master index to use.</summary>
    [Parameter]
    [ValidateRange(0, int.MaxValue)]
    public int Master { get; set; }

    /// <summary>Layout index to use.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetIndex)]
    [ValidateRange(0, int.MaxValue)]
    public int Layout { get; set; }

    /// <summary>Layout name to use.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetByName)]
    public string LayoutName { get; set; } = string.Empty;

    /// <summary>Layout type to use.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetByType)]
    public PowerPointSlideLayoutType LayoutType { get; set; }

    /// <summary>Use case-sensitive matching for layout names.</summary>
    [Parameter(ParameterSetName = ParameterSetByName)]
    public SwitchParameter CaseSensitive { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try {
            var slide = Slide ?? PowerPointDslContext.Require(this).RequireSlide();
            switch (ParameterSetName) {
                case ParameterSetByName:
                    slide.SetLayout(LayoutName, Master, ignoreCase: !CaseSensitive.IsPresent);
                    break;
                case ParameterSetByType:
                    slide.SetLayout(LayoutType, Master);
                    break;
                default:
                    slide.SetLayout(Master, Layout);
                    break;
            }

            WritePassThru(slide);
        } catch (Exception ex) {
            WriteError(new ErrorRecord(ex, "PowerPointSetSlideLayoutFailed", ErrorCategory.InvalidOperation, Slide));
        }
    }
}