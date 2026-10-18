using System;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Finds PowerPoint shapes by text, name, kind, or slide.</summary>
/// <para>
/// Searches an open presentation or a single slide and returns <see cref="PowerPointShapeInfo"/>
/// records that include the slide index, shape index, shape kind, extracted text, shape name, and the
/// underlying OfficeIMO shape object. Text matching includes normal text boxes and table cell text, so
/// this command can locate the right object before piping it into modification commands.
/// </para>
/// <para>
/// Use <c>-Text</c> for literal contains matching or <c>-Pattern</c> for regular expressions, or omit
/// both text parameters when <c>-Kind</c>, <c>-Name</c>, <c>-Index</c>, and <c>-ShapeIndex</c> identify
/// the target shape without reading text content.
/// </para>
/// <example>
///   <summary>Find and update a text box.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Find-OfficePowerPointShape -Presentation $ppt -Text 'FY24' -Kind TextBox |
///     Set-OfficePowerPointShapeText -Text 'FY25'</code>
///   <para>Finds matching text shapes and updates them without using the PowerPoint DSL.</para>
/// </example>
/// <example>
///   <summary>Find a table by cell text and append a new status row.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$ppt = Get-OfficePowerPoint -Path .\Readiness.pptx
/// $table = Find-OfficePowerPointShape -Presentation $ppt -Text 'Risk' -Kind Table | Select-Object -First 1
/// $table | Add-OfficePowerPointTableRow -Values 'Latency', 'Investigating'
/// $ppt | Close-OfficePowerPoint -Save</code>
///   <para>Uses table-cell text as the locator, then pipes the table shape metadata into a table-row edit.</para>
/// </example>
[Cmdlet(VerbsCommon.Find, "OfficePowerPointShape", DefaultParameterSetName = ParameterSetPresentationText)]
[OutputType(typeof(PowerPointShapeInfo))]
public sealed class FindOfficePowerPointShapeCommand : PSCmdlet
{
    private const string ParameterSetPresentationText = "PresentationText";
    private const string ParameterSetPresentationRegex = "PresentationRegex";
    private const string ParameterSetSlideText = "SlideText";
    private const string ParameterSetSlideRegex = "SlideRegex";

    /// <summary>Open presentation whose slides should be searched.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetPresentationText)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetPresentationRegex)]
    public PowerPointPresentation? Presentation { get; set; }

    /// <summary>Single slide to search when the caller has already resolved the slide.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetSlideText)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetSlideRegex)]
    public PowerPointSlide? Slide { get; set; }

    /// <summary>Literal text to find in text boxes and table cells.</summary>
    [Parameter(Position = 0, ParameterSetName = ParameterSetPresentationText)]
    [Parameter(Position = 0, ParameterSetName = ParameterSetSlideText)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Regular expression to find in text boxes and table cells.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPresentationRegex)]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetSlideRegex)]
    public string Pattern { get; set; } = string.Empty;

    /// <summary>Use case-sensitive matching for text and name filters.</summary>
    [Parameter]
    public SwitchParameter CaseSensitive { get; set; }

    /// <summary>Optional zero-based slide index when reading from a presentation.</summary>
    [Parameter(ParameterSetName = ParameterSetPresentationText)]
    [Parameter(ParameterSetName = ParameterSetPresentationRegex)]
    public int? Index { get; set; }

    /// <summary>Optional zero-based shape index filter, useful when several shapes contain the same text.</summary>
    [Parameter]
    public int[]? ShapeIndex { get; set; }

    /// <summary>Optional wildcard filter for shape names.</summary>
    [Parameter]
    public string[]? Name { get; set; }

    /// <summary>Optional shape kind filter such as TextBox or Table.</summary>
    [Parameter]
    public OfficePowerPointShapeKind[]? Kind { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var hasTextFilter = ParameterSetName.EndsWith("Regex", StringComparison.Ordinal)
            || MyInvocation.BoundParameters.ContainsKey(nameof(Text));
        var matcher = ParameterSetName.EndsWith("Regex", StringComparison.Ordinal)
            ? PowerPointObjectSearch.CreateRegexMatcher(Pattern, CaseSensitive.IsPresent)
            : PowerPointObjectSearch.CreateTextMatcher(Text, CaseSensitive.IsPresent);

        if (ParameterSetName.StartsWith("Slide", StringComparison.Ordinal))
        {
            WriteMatchingShapes(Slide ?? throw new InvalidOperationException("Slide was not provided."), ResolveSlideIndex(Slide, Presentation));
            return;
        }

        var presentation = Presentation ?? throw new InvalidOperationException("Presentation was not provided.");
        if (Index.HasValue)
        {
            if (Index.Value < 0 || Index.Value >= presentation.Slides.Count)
            {
                throw new PSArgumentOutOfRangeException(nameof(Index), Index.Value, $"Presentation contains {presentation.Slides.Count} slides.");
            }

            WriteMatchingShapes(presentation.Slides[Index.Value], Index.Value);
            return;
        }

        for (var slideIndex = 0; slideIndex < presentation.Slides.Count; slideIndex++)
        {
            WriteMatchingShapes(presentation.Slides[slideIndex], slideIndex);
        }

        void WriteMatchingShapes(PowerPointSlide slide, int slideIndex)
        {
            for (var shapeIndex = 0; shapeIndex < slide.Shapes.Count; shapeIndex++)
            {
                var info = PowerPointShapeReader.Read(slide, slide.Shapes[shapeIndex], slideIndex, shapeIndex);
                if (MatchesMetadata(info) && (!hasTextFilter || PowerPointObjectSearch.MatchesShape(info, matcher)))
                {
                    WriteObject(info);
                }
            }
        }
    }

    private bool MatchesMetadata(PowerPointShapeInfo info)
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

            var options = CaseSensitive.IsPresent ? WildcardOptions.None : WildcardOptions.IgnoreCase;
            return Name.Any(pattern => new WildcardPattern(pattern, options).IsMatch(info.Name));
        }

        return true;
    }

    private static int ResolveSlideIndex(PowerPointSlide? slide, PowerPointPresentation? presentation)
    {
        if (slide == null || presentation == null)
        {
            return -1;
        }

        for (var index = 0; index < presentation.Slides.Count; index++)
        {
            if (ReferenceEquals(presentation.Slides[index], slide))
            {
                return index;
            }
        }

        return -1;
    }
}
