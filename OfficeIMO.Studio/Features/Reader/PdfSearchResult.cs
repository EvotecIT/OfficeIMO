using System.Globalization;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Reader;

public sealed record PdfSearchResult(int PageNumber, string Snippet) {
    private static readonly IStudioLocalizer DefaultLocalizer = new StudioLocalizer(CultureInfo.GetCultureInfo("en"));

    internal IStudioLocalizer Localizer { get; init; } = DefaultLocalizer;

    internal static PdfSearchResult FromHit(PdfSearchHit hit, IStudioLocalizer localizer) => new(hit.PageNumber, hit.Snippet) {
        Bounds = new(hit.Bounds.X, hit.Bounds.Y, hit.Bounds.Width, hit.Bounds.Height),
        LineBounds = hit.LineBounds.Select(static line => new Avalonia.Rect(line.X, line.Y, line.Width, line.Height)).ToArray(),
        OccurrenceNumber = hit.OccurrenceNumber, Localizer = localizer
    };

    public Avalonia.Rect Bounds { get; init; }

    /// <summary>Per-line highlight rectangles; a match that wraps across lines has one rectangle per line segment.</summary>
    public IReadOnlyList<Avalonia.Rect> LineBounds { get; init; } = Array.Empty<Avalonia.Rect>();

    internal IReadOnlyList<Avalonia.Rect> Highlights => LineBounds.Count > 0 ? LineBounds : new[] { Bounds };

    public int OccurrenceNumber { get; init; }

    public string Label => Localizer.Format("Search.ResultLabel", PageNumber, Snippet, OccurrenceNumber);

}
