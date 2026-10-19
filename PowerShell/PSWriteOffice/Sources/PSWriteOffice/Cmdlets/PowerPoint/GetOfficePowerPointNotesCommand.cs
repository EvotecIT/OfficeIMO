using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Reads speaker notes from a slide or presentation.</summary>
/// <para>Returns note metadata without creating empty notes parts on slides that do not already have them.</para>
/// <example>
///   <summary>Read notes from one slide.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficePowerPointSlide -Presentation $ppt -Index 0 | Get-OfficePowerPointNotes</code>
///   <para>Returns the notes text and metadata for the selected slide.</para>
/// </example>
/// <example>
///   <summary>Enumerate notes across a deck.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficePowerPointNotes -Presentation $ppt -IncludeEmpty</code>
///   <para>Lists slide indexes together with note text, including slides that have no notes yet.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePowerPointNotes", DefaultParameterSetName = ParameterSetSlide)]
[OutputType(typeof(PowerPointNotesInfo))]
public sealed class GetOfficePowerPointNotesCommand : PSCmdlet
{
    private const string ParameterSetSlide = "Slide";
    private const string ParameterSetPresentation = "Presentation";

    /// <summary>Slide to inspect (optional inside the DSL).</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetSlide)]
    public PowerPointSlide? Slide { get; set; }

    /// <summary>Presentation whose slides should be inspected.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetPresentation)]
    public PowerPointPresentation? Presentation { get; set; }

    /// <summary>Optional zero-based slide index when reading from a presentation.</summary>
    [Parameter(ParameterSetName = ParameterSetPresentation)]
    public int? Index { get; set; }

    /// <summary>Include slides that do not currently have speaker notes.</summary>
    [Parameter]
    public SwitchParameter IncludeEmpty { get; set; }

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
                        WriteError(new ErrorRecord(new ArgumentOutOfRangeException(nameof(Index)), "PowerPointNotesIndexOutOfRange", ErrorCategory.InvalidArgument, Index));
                        return;
                    }

                    WriteInfo(presentation.Slides[Index.Value], Index.Value);
                    return;
                }

                for (var i = 0; i < presentation.Slides.Count; i++)
                {
                    WriteInfo(presentation.Slides[i], i);
                }

                return;
            }

            var context = PowerPointDslContext.Current;
            var slide = Slide ?? context?.RequireSlide()
                ?? throw new InvalidOperationException("Slide was not provided. Use -Slide or run inside Add-OfficePowerPointSlide / PptSlide.");

            var slideIndex = context?.Presentation != null
                ? ResolveSlideIndex(slide, context.Presentation)
                : PowerPointNotesReader.ResolveSlideIndex(slide);

            WriteInfo(slide, slideIndex);
        }
        catch (Exception ex)
        {
            var target = (object?)Slide ?? Presentation;
            WriteError(new ErrorRecord(ex, "PowerPointGetNotesFailed", ErrorCategory.InvalidOperation, target));
        }
    }

    private void WriteInfo(PowerPointSlide slide, int slideIndex)
    {
        var info = PowerPointNotesReader.Read(slide, slideIndex);
        if (IncludeEmpty.IsPresent || info.HasNotes || !string.IsNullOrWhiteSpace(info.Text))
        {
            WriteObject(info);
        }
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
