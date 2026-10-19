using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Reads shape summaries from a slide or presentation.</summary>
/// <para>Returns PowerShell-friendly metadata for text boxes, pictures, tables, charts, and auto shapes.</para>
/// <example>
///   <summary>Inspect one slide.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficePowerPointSlide -Presentation $ppt -Index 0 | Get-OfficePowerPointShape</code>
///   <para>Returns shape summaries for the selected slide.</para>
/// </example>
/// <example>
///   <summary>Find pictures on a slide.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficePowerPointShape -Presentation $ppt -Index 0 -Kind Picture</code>
///   <para>Filters the slide output to picture shapes only.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePowerPointShape", DefaultParameterSetName = ParameterSetSlide)]
[OutputType(typeof(PowerPointShapeInfo))]
public sealed class GetOfficePowerPointShapeCommand : PSCmdlet
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

    /// <summary>Optional zero-based shape index filter.</summary>
    [Parameter]
    public int[]? ShapeIndex { get; set; }

    /// <summary>Optional wildcard filter for shape names.</summary>
    [Parameter]
    public string[]? Name { get; set; }

    /// <summary>Optional shape kind filter.</summary>
    [Parameter]
    public OfficePowerPointShapeKind[]? Kind { get; set; }

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
                        WriteError(new ErrorRecord(new ArgumentOutOfRangeException(nameof(Index)), "PowerPointShapeIndexOutOfRange", ErrorCategory.InvalidArgument, Index));
                        return;
                    }

                    WriteShapes(presentation.Slides[Index.Value], Index.Value);
                    return;
                }

                for (var i = 0; i < presentation.Slides.Count; i++)
                {
                    WriteShapes(presentation.Slides[i], i);
                }

                return;
            }

            var context = PowerPointDslContext.Current;
            var slide = Slide ?? context?.RequireSlide()
                ?? throw new InvalidOperationException("Slide was not provided. Use -Slide or run inside Add-OfficePowerPointSlide / PptSlide.");

            var slideIndex = context?.Presentation != null
                ? ResolveSlideIndex(slide, context.Presentation)
                : PowerPointNotesReader.ResolveSlideIndex(slide);

            WriteShapes(slide, slideIndex);
        }
        catch (Exception ex)
        {
            var target = (object?)Slide ?? Presentation;
            WriteError(new ErrorRecord(ex, "PowerPointGetShapeFailed", ErrorCategory.InvalidOperation, target));
        }
    }

    private void WriteShapes(PowerPointSlide slide, int slideIndex)
    {
        for (var i = 0; i < slide.Shapes.Count; i++)
        {
            var shape = slide.Shapes[i];
            var info = PowerPointShapeReader.Read(slide, shape, slideIndex, i);
            if (Matches(info))
            {
                WriteObject(info);
            }
        }
    }

    private bool Matches(PowerPointShapeInfo info)
    {
        if (ShapeIndex != null && ShapeIndex.Length > 0 && !ShapeIndex.Contains(info.ShapeIndex))
        {
            return false;
        }

        if (Kind != null && Kind.Length > 0 && !Kind.Any(kind => string.Equals(kind.ToString(), info.Kind, StringComparison.OrdinalIgnoreCase)))
        {
            return false;
        }

        if (Name != null && Name.Length > 0)
        {
            if (string.IsNullOrWhiteSpace(info.Name))
            {
                return false;
            }

            var matchesName = Name.Any(pattern =>
            {
                var wildcard = new WildcardPattern(pattern, WildcardOptions.IgnoreCase);
                return wildcard.IsMatch(info.Name);
            });

            if (!matchesName)
            {
                return false;
            }
        }

        return true;
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
