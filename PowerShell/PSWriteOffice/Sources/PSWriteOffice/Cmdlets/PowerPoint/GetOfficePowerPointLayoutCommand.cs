using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Lists slide layouts available in a presentation.</summary>
/// <example>
///   <summary>Enumerate layouts for the first master.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficePowerPointLayout -Presentation $ppt</code>
///   <para>Returns layout metadata including name, type, and index.</para>
/// </example>
/// <example>
///   <summary>Get the first layout inside the DSL.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePowerPoint -Path .\deck.pptx { $layout = @(Get-OfficePowerPointLayout)[0] }</code>
///   <para>Uses the current DSL presentation context.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePowerPointLayout")]
[OutputType(typeof(PowerPointSlideLayoutInfo))]
public sealed class GetOfficePowerPointLayoutCommand : PSCmdlet
{
    /// <summary>Presentation to inspect (optional inside DSL).</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointPresentation? Presentation { get; set; }

    /// <summary>Slide master index.</summary>
    [Parameter]
    [ValidateRange(0, int.MaxValue)]
    public int Master { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        try
        {
            if (Master < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(Master), "Master index cannot be negative.");
            }

            var presentation = Presentation ?? PowerPointDslContext.Require(this).Presentation;
            var layouts = presentation.GetSlideLayouts(Master);
            foreach (var layout in layouts)
            {
                WriteObject(layout);
            }
        }
        catch (Exception ex) when (ex.GetType().Name != "StopUpstreamCommandsException")
        {
            WriteError(new ErrorRecord(ex, "PowerPointGetLayoutFailed", ErrorCategory.InvalidOperation, Presentation));
        }
    }
}
