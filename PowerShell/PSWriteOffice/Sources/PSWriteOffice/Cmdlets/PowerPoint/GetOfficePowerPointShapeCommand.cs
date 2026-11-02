using System;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Gets editable shape summaries, with optional text, name, kind and position filters.</summary>
/// <para>Searches every slide of the document unless a native slide or one-based slide number is supplied. Inside PowerPointSlide, omission uses the current slide. Text matching includes table cells.</para>
/// <example><summary>Inspect one slide.</summary>
/// <code>Get-OfficePowerPointShape -Document $deck -Slide 2 -Kind Picture</code></example>
/// <example><summary>Find and replace a label.</summary>
/// <code>Get-OfficePowerPointShape -Document $deck -Text 'Draft' -Kind TextBox |
///     Set-OfficePowerPointShapeText -Text 'Ready'</code></example>
[Cmdlet(VerbsCommon.Get, "OfficePowerPointShape", DefaultParameterSetName = "All")]
[OutputType(typeof(PowerPointShapeInfo))]
public sealed class GetOfficePowerPointShapeCommand : OfficeCmdlet {
    /// <summary>Live document; omission uses the enclosing document block.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = "All")]
    [Parameter(ValueFromPipeline = true, ParameterSetName = "Text")]
    [Parameter(ValueFromPipeline = true, ParameterSetName = "Pattern")]
    public PowerPointDocument? Document { get; set; }

    /// <summary>Native slide or one-based slide number.</summary>
    [Parameter, PowerPointSlideTransformation]
    public PowerPointSlideSelector? Slide { get; set; }

    /// <summary>One-based shape positions within each selected slide.</summary>
    [Parameter, ValidateRange(1, int.MaxValue)]
    public int[]? Index { get; set; }

    /// <summary>Wildcard shape-name filters.</summary>
    [Parameter, SupportsWildcards]
    public string[]? Name { get; set; }

    /// <summary>Shape kinds to include.</summary>
    [Parameter]
    public OfficePowerPointShapeKind[]? Kind { get; set; }

    /// <summary>Literal text to find in text boxes or table cells.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Text"), AllowEmptyString]
    public string Text { get; set; } = string.Empty;

    /// <summary>Regular expression to find in text boxes or table cells.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Pattern"), ValidateNotNullOrEmpty]
    public string Pattern { get; set; } = string.Empty;

    /// <summary>Match text and names with case sensitivity.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Pattern")]
    public SwitchParameter CaseSensitive { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try {
            Func<string, bool>? matcher = ParameterSetName switch {
                "Pattern" => PowerPointObjectSearch.CreateRegexMatcher(Pattern, CaseSensitive),
                "Text" => PowerPointObjectSearch.CreateTextMatcher(Text, CaseSensitive),
                _ => null
            };
            var document = Document ?? OfficeDocumentContext.Find<PowerPointDocument>();
            var current = Document == null ? PowerPointContentScope.Current?.NativeSlide : null;
            if (Slide != null) WriteShapes(Slide.Resolve(document), matcher);
            else if (current != null) WriteShapes(current, matcher);
            else foreach (var slide in ResolveDocument(document).Slides) WriteShapes(slide, matcher);
        } catch (Exception error) {
            if (error is PipelineStoppedException or ActionPreferenceStopException) throw;
            if (error is ArgumentException argument) InvalidArgument(argument, Slide);
            else ItemError(error, "powerpoint-shape-read-failed", (object?)Slide ?? Document);
        }
    }

    private void WriteShapes(PowerPointSlide slide, Func<string, bool>? matcher) {
        var wildcardOptions = CaseSensitive.IsPresent ? WildcardOptions.None : WildcardOptions.IgnoreCase;
        var names = Name?.Select(pattern => new WildcardPattern(pattern, wildcardOptions)).ToArray();
        for (int index = 0; index < slide.Shapes.Count; index++) {
            var info = PowerPointShapeReader.Read(slide, slide.Shapes[index], slide.Index, index);
            if (Index?.Length > 0 && !Index.Contains(info.ShapeIndex)) continue;
            if (Kind?.Length > 0 && !Kind.Any(kind => string.Equals(kind.ToString(), info.Kind, StringComparison.OrdinalIgnoreCase))) continue;
            if (names?.Length > 0 && !names.Any(pattern => pattern.IsMatch(info.Name ?? string.Empty))) continue;
            if (matcher != null && !PowerPointObjectSearch.MatchesShape(info, matcher)) continue;
            WriteObject(info);
        }
    }
}
