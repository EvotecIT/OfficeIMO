using System.Globalization;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Reader;

public sealed record PdfSearchHit(int PageNumber, string Snippet) {
    private static readonly IStudioLocalizer DefaultLocalizer = new StudioLocalizer(CultureInfo.GetCultureInfo("en"));

    internal IStudioLocalizer Localizer { get; init; } = DefaultLocalizer;

    public Avalonia.Rect Bounds { get; init; }

    /// <summary>Per-line highlight rectangles; a match that wraps across lines has one rectangle per line segment.</summary>
    public IReadOnlyList<Avalonia.Rect> LineBounds { get; init; } = Array.Empty<Avalonia.Rect>();

    internal IReadOnlyList<Avalonia.Rect> Highlights => LineBounds.Count > 0 ? LineBounds : new[] { Bounds };

    public int OccurrenceNumber { get; init; }

    public string Label => Localizer.Format("Search.ResultLabel", PageNumber, Snippet, OccurrenceNumber);

    internal PdfSearchHit WithLocalizer(IStudioLocalizer localizer) =>
        this with { Localizer = localizer ?? throw new ArgumentNullException(nameof(localizer)) };
}
