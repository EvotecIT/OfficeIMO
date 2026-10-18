using System;
using System.Management.Automation;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Sets layout placeholder text margins for a slide layout (points).</summary>
/// <example>
///   <summary>Update layout placeholder text margins.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Set-OfficePowerPointLayoutPlaceholderTextMargins -Presentation $ppt -Master 0 -Layout 1 -PlaceholderType Title -Left 12 -Top 8 -Right 12 -Bottom 8</code>
///   <para>Updates the text margins on the layout placeholder.</para>
/// </example>
/// <example>
///   <summary>Update margins inside the DSL.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePowerPoint -Path .\deck.pptx {
///     $layout = Get-OfficePowerPointLayout | Select-Object -First 1
///     Set-OfficePowerPointLayoutPlaceholderTextMargins -Master $layout.MasterIndex -Layout $layout.LayoutIndex -PlaceholderType Title -Left 12 -Top 8 -Right 12 -Bottom 8
///   }</code>
///   <para>Uses the DSL context to resolve the presentation.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePowerPointLayoutPlaceholderTextMargins")]
[Alias("PptLayoutPlaceholderMargins")]
[OutputType(typeof(PowerPointTextBox))]
public sealed class SetOfficePowerPointLayoutPlaceholderTextMarginsCommand : PSCmdlet
{
    /// <summary>Presentation to update (optional inside DSL).</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointPresentation? Presentation { get; set; }

    /// <summary>Slide master index.</summary>
    [Parameter]
    [ValidateRange(0, int.MaxValue)]
    public int Master { get; set; } = 0;

    /// <summary>Layout index within the master.</summary>
    [Parameter(Mandatory = true)]
    [ValidateRange(0, int.MaxValue)]
    public int Layout { get; set; }

    /// <summary>Placeholder type to target.</summary>
    [Parameter(Mandatory = true)]
    [Alias("Type")]
    public PowerPointPlaceholderType PlaceholderType { get; set; }

    /// <summary>Optional placeholder index.</summary>
    [Parameter]
    public uint? Index { get; set; }

    /// <summary>Left margin in points.</summary>
    [Parameter(Mandatory = true)]
    public double Left { get; set; }

    /// <summary>Top margin in points.</summary>
    [Parameter(Mandatory = true)]
    public double Top { get; set; }

    /// <summary>Right margin in points.</summary>
    [Parameter(Mandatory = true)]
    public double Right { get; set; }

    /// <summary>Bottom margin in points.</summary>
    [Parameter(Mandatory = true)]
    public double Bottom { get; set; }

    /// <summary>Create the placeholder if it is missing.</summary>
    [Parameter]
    public SwitchParameter CreateIfMissing { get; set; }

    /// <summary>Emit the placeholder textbox after update.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        PowerPointPresentation? presentation = null;
        try
        {
            presentation = Presentation ?? PowerPointDslContext.Require(this).Presentation;
            presentation.SetLayoutPlaceholderTextMarginsPoints(
                Master,
                Layout,
                PlaceholderType,
                Left,
                Top,
                Right,
                Bottom,
                Index,
                CreateIfMissing.IsPresent);

            if (PassThru.IsPresent)
            {
                var textBox = presentation.GetLayoutPlaceholderTextBox(Master, Layout, PlaceholderType, Index);
                if (textBox != null)
                {
                    WriteObject(textBox);
                }
            }
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "PowerPointSetLayoutPlaceholderTextMarginsFailed", ErrorCategory.InvalidOperation, presentation ?? Presentation));
        }
    }
}
