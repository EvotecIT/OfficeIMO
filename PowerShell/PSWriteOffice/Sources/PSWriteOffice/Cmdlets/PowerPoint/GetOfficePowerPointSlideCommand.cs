using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Enumerates slides or retrieves a specific slide.</summary>
/// <para>Supports pipeline-friendly iteration over <see cref="PowerPointPresentation.Slides"/> or direct index selection.</para>
/// <example>
///   <summary>List slide titles.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficePowerPointSlide -Presentation $ppt | Get-OfficePowerPointPlaceholder -PlaceholderType Title | Select-Object -ExpandProperty Text</code>
///   <para>Streams each slide so you can read the title placeholder text.</para>
/// </example>
/// <example>
///   <summary>Enumerate slides inside the DSL.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePowerPoint -Path .\deck.pptx { Get-OfficePowerPointSlide | Select-Object -First 1 }</code>
///   <para>Uses the current DSL presentation context.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePowerPointSlide")]
public class GetOfficePowerPointSlideCommand : PSCmdlet
{
    /// <summary>Presentation to inspect (optional inside DSL).</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointPresentation? Presentation { get; set; }

    /// <summary>Optional zero-based index; omit to enumerate all slides.</summary>
    [Parameter]
    public int? Index { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var presentation = Presentation ?? PowerPointDslContext.Current?.Presentation
            ?? throw new InvalidOperationException("Presentation was not provided. Use -Presentation or run inside New-OfficePowerPoint.");

        if (Index.HasValue)
        {
            if (Index.Value < 0 || Index.Value >= presentation.Slides.Count)
            {
                WriteError(new ErrorRecord(new ArgumentOutOfRangeException(nameof(Index)), "PowerPointSlideIndexOutOfRange", ErrorCategory.InvalidArgument, Index));
                return;
            }

            WriteObject(presentation.Slides[Index.Value]);
        }
        else
        {
            foreach (var slide in presentation.Slides)
            {
                WriteObject(slide);
            }
        }
    }
}
