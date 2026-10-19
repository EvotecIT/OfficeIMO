using System;
using System.Management.Automation;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Sets layout placeholder bounds for a slide layout.</summary>
/// <example>
///   <summary>Update title placeholder bounds.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Set-OfficePowerPointLayoutPlaceholderBounds -Presentation $ppt -Master 0 -Layout 1 -PlaceholderType Title -Left 40 -Top 20 -Width 500 -Height 120</code>
///   <para>Moves/resizes the Title placeholder on the layout.</para>
/// </example>
/// <example>
///   <summary>Update placeholder bounds inside the DSL.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePowerPoint -Path .\deck.pptx {
///     $layout = Get-OfficePowerPointLayout | Select-Object -First 1
///     Set-OfficePowerPointLayoutPlaceholderBounds -Master $layout.MasterIndex -Layout $layout.LayoutIndex -PlaceholderType Title -Left 40 -Top 20 -Width 500 -Height 120
///   }</code>
///   <para>Uses the DSL context to resolve the presentation.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePowerPointLayoutPlaceholderBounds")]
[Alias("PptLayoutPlaceholderBounds")]
[OutputType(typeof(PowerPointTextBox))]
public sealed class SetOfficePowerPointLayoutPlaceholderBoundsCommand : PSCmdlet
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

    /// <summary>Left position in points.</summary>
    [Parameter(Mandatory = true)]
    public double Left { get; set; }

    /// <summary>Top position in points.</summary>
    [Parameter(Mandatory = true)]
    public double Top { get; set; }

    /// <summary>Width in points.</summary>
    [Parameter(Mandatory = true)]
    public double Width { get; set; }

    /// <summary>Height in points.</summary>
    [Parameter(Mandatory = true)]
    public double Height { get; set; }

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
            var bounds = PowerPointLayoutBox.FromPoints(Left, Top, Width, Height);
            presentation = Presentation ?? PowerPointDslContext.Require(this).Presentation;
            presentation.SetLayoutPlaceholderBounds(Master, Layout, PlaceholderType, bounds, Index, CreateIfMissing.IsPresent);

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
            WriteError(new ErrorRecord(ex, "PowerPointSetLayoutPlaceholderBoundsFailed", ErrorCategory.InvalidOperation, presentation ?? Presentation));
        }
    }
}
