namespace OfficeIMO.IWork;

/// <summary>Immutable spacing and custom tab stops for an owned native Pages body paragraph.</summary>
/// <remarks>Paragraphs remain left-aligned, with zero indents and a 36-point default tab interval.
/// Custom layout uses a period for decimal tabs. No tab leaders are authored.</remarks>
public sealed class IWorkPagesParagraphLayout {
    internal static IWorkPagesParagraphLayout Default { get; } = new();

    /// <summary>Creates a paragraph layout with positive relative line spacing and nonnegative before/after spacing.</summary>
    /// <param name="lineSpacingMultiplier">Relative line height; 1 means single spacing.</param>
    /// <param name="spaceBeforePoints">Additional space before the paragraph in points.</param>
    /// <param name="spaceAfterPoints">Additional space after the paragraph in points.</param>
    /// <param name="tabStops">Custom stops, copied and sorted by position. Positions must be unique.</param>
    public IWorkPagesParagraphLayout(float lineSpacingMultiplier = 1, float spaceBeforePoints = 0,
        float spaceAfterPoints = 0, IEnumerable<IWorkTabStop>? tabStops = null) {
        IWorkCreationValidation.ValidatePositive(lineSpacingMultiplier, nameof(lineSpacingMultiplier));
        foreach (var (value, name) in new[] { (spaceBeforePoints, nameof(spaceBeforePoints)), (spaceAfterPoints, nameof(spaceAfterPoints)) }) {
            IWorkCreationValidation.ValidateFinite(value, name);
            if (value < 0) throw new ArgumentOutOfRangeException(name, "Paragraph spacing cannot be negative.");
        }
        IWorkTabStop[] stops = (tabStops ?? Array.Empty<IWorkTabStop>()).ToArray();
        if (stops.Any(stop => stop == null)) throw new ArgumentException("Tab stops cannot contain null.", nameof(tabStops));
        stops = stops.Select(stop => IWorkTabStop.Create((float)stop.PositionPoints, stop.Alignment))
            .OrderBy(stop => stop.PositionPoints).ToArray();
        if (stops.Select(stop => stop.PositionPoints).Distinct().Count() != stops.Length)
            throw new ArgumentException("Custom tab positions must be unique.", nameof(tabStops));
        LineSpacingMultiplier = lineSpacingMultiplier;
        SpaceBeforePoints = spaceBeforePoints == 0 ? 0 : spaceBeforePoints;
        SpaceAfterPoints = spaceAfterPoints == 0 ? 0 : spaceAfterPoints;
        TabStops = Array.AsReadOnly(stops);
    }

    /// <summary>Gets the relative line-height multiplier.</summary>
    public float LineSpacingMultiplier { get; }
    /// <summary>Gets additional space before the paragraph in points.</summary>
    public float SpaceBeforePoints { get; }
    /// <summary>Gets additional space after the paragraph in points.</summary>
    public float SpaceAfterPoints { get; }
    /// <summary>Gets custom tab stops in increasing position order.</summary>
    public IReadOnlyList<IWorkTabStop> TabStops { get; }

    internal bool HasCustomProperties => LineSpacingMultiplier != 1 || SpaceBeforePoints != 0
        || SpaceAfterPoints != 0 || TabStops.Count != 0;
}
