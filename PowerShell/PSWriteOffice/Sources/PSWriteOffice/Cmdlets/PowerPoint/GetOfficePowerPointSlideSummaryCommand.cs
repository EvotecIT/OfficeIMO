using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Reads high-level slide summaries from a presentation.</summary>
/// <para>Returns title, notes metadata, layout metadata, and content counts for each slide.</para>
/// <example>
///   <summary>Summarize a deck.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficePowerPointSlideSummary -Presentation $ppt</code>
///   <para>Returns one summary object per slide.</para>
/// </example>
/// <example>
///   <summary>Summarize one slide.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficePowerPointSlide -Presentation $ppt -Index 0 | Get-OfficePowerPointSlideSummary</code>
///   <para>Returns the summary for the selected slide.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePowerPointSlideSummary", DefaultParameterSetName = ParameterSetSlide)]
[OutputType(typeof(PowerPointSlideSummaryInfo))]
public sealed class GetOfficePowerPointSlideSummaryCommand : PSCmdlet
{
    private const string ParameterSetSlide = "Slide";
    private const string ParameterSetPresentation = "Presentation";

    /// <summary>Slide to inspect (optional inside the DSL).</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetSlide)]
    public PowerPointSlide? Slide { get; set; }

    /// <summary>Presentation whose slides should be summarized.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetPresentation)]
    public PowerPointPresentation? Presentation { get; set; }

    /// <summary>Optional zero-based slide index when reading from a presentation.</summary>
    [Parameter(ParameterSetName = ParameterSetPresentation)]
    public int? Index { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        try
        {
            if (ParameterSetName == ParameterSetPresentation)
            {
                var presentation = Presentation ?? PowerPointDslContext.Current?.Presentation
                    ?? throw new InvalidOperationException("Presentation was not provided. Use -Presentation or run inside New-OfficePowerPoint.");

                if (Index.HasValue)
                {
                    if (Index.Value < 0 || Index.Value >= presentation.Slides.Count)
                    {
                        WriteError(new ErrorRecord(new ArgumentOutOfRangeException(nameof(Index)), "PowerPointSlideSummaryIndexOutOfRange", ErrorCategory.InvalidArgument, Index));
                        return;
                    }

                    WriteSummary(presentation.Slides[Index.Value], Index.Value);
                    return;
                }

                for (var i = 0; i < presentation.Slides.Count; i++)
                {
                    WriteSummary(presentation.Slides[i], i);
                }

                return;
            }

            var context = PowerPointDslContext.Current;
            var slide = Slide ?? context?.RequireSlide()
                ?? throw new InvalidOperationException("Slide was not provided. Use -Slide or run inside Add-OfficePowerPointSlide / PptSlide.");

            var slideIndex = context?.Presentation != null
                ? ResolveSlideIndex(slide, context.Presentation)
                : PowerPointNotesReader.ResolveSlideIndex(slide);

            WriteSummary(slide, slideIndex);
        }
        catch (Exception ex)
        {
            var target = (object?)Slide ?? Presentation;
            WriteError(new ErrorRecord(ex, "PowerPointGetSlideSummaryFailed", ErrorCategory.InvalidOperation, target));
        }
    }

    private void WriteSummary(PowerPointSlide slide, int slideIndex)
    {
        WriteObject(PowerPointSlideSummaryReader.Read(slide, slideIndex));
    }

    private static int ResolveSlideIndex(PowerPointSlide slide, PowerPointPresentation? presentation)
    {
        if (presentation == null)
        {
            return -1;
        }

        for (var i = 0; i < presentation.Slides.Count; i++)
        {
            if (ReferenceEquals(presentation.Slides[i], slide))
            {
                return i;
            }
        }

        return -1;
    }
}
