using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a new slide to a PowerPoint presentation.</summary>
/// <para>Creates a slide using OfficeIMO master/layout indexes and can execute nested DSL content.</para>
/// <example>
///   <summary>Append a slide with the default layout.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$ppt = New-OfficePowerPoint -Path .\deck.pptx -NoSave
/// Add-OfficePowerPointSlide -Presentation $ppt
/// $ppt | Close-OfficePowerPoint -Save</code>
///   <para>Creates a deck and appends a new slide at the end.</para>
/// </example>
/// <example>
///   <summary>Create a slide using the DSL.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePowerPoint -Path .\deck.pptx { PptSlide { PptTitle -Title 'Status Update' } }</code>
///   <para>Creates a slide and sets the title using DSL aliases.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointSlide", DefaultParameterSetName = ParameterSetIndex)]
[Alias("PptSlide")]
public class AddOfficePowerPointSlideCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    private const string ParameterSetIndex = "Index";
    private const string ParameterSetByName = "Name";
    private const string ParameterSetByType = "Type";

    /// <summary>Presentation to update (optional inside New-OfficePowerPoint).</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointPresentation? Presentation { get; set; }

    /// <summary>Slide master index to use.</summary>
    [Parameter]
    [ValidateRange(0, int.MaxValue)]
    public int Master { get; set; } = 0;

    /// <summary>Layout index to use (matches the template’s built-in layouts).</summary>
    [Parameter(ParameterSetName = ParameterSetIndex)]
    [ValidateRange(0, int.MaxValue)]
    public int Layout { get; set; } = 1;

    /// <summary>Layout name to use (case-insensitive by default).</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetByName)]
    public string LayoutName { get; set; } = string.Empty;

    /// <summary>Layout type to use.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetByType)]
    public PowerPointSlideLayoutType LayoutType { get; set; }

    /// <summary>Use case-sensitive matching for layout names.</summary>
    [Parameter(ParameterSetName = ParameterSetByName)]
    public SwitchParameter CaseSensitive { get; set; }

    /// <summary>Nested DSL content executed within the slide scope.</summary>
    [Parameter(Position = 0)]
    public ScriptBlock? Content { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        PowerPointPresentation? presentation = Presentation;
        try {
            var context = PowerPointDslContext.Current;
            if (presentation == null) {
                presentation = (context ?? PowerPointDslContext.Require(this)).Presentation;
            }

            PowerPointSlide slide;
            switch (ParameterSetName) {
                case ParameterSetByName:
                    slide = presentation.AddSlide(LayoutName, Master, ignoreCase: !CaseSensitive.IsPresent);
                    break;
                case ParameterSetByType:
                    slide = presentation.AddSlide(LayoutType, Master);
                    break;
                default:
                    slide = presentation.AddSlide(Master, Layout);
                    break;
            }

            if (Content != null) {
                if (context != null) {
                    using (context.Push(slide)) {
                        Content.InvokeReturnAsIs();
                    }
                } else {
                    using (var scoped = PowerPointDslContext.Enter(presentation))
                    using (scoped.Push(slide)) {
                        Content.InvokeReturnAsIs();
                    }
                }
            }

            WritePassThru(slide);
        } catch (Exception ex) {
            WriteError(new ErrorRecord(ex, "PowerPointAddSlideFailed", ErrorCategory.InvalidOperation, presentation ?? Presentation));
        }
    }
}
