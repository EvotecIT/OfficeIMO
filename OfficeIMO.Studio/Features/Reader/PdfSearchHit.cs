using System.Globalization;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Reader;

public sealed record PdfSearchHit(int PageNumber, string Snippet) {
    private static readonly IStudioLocalizer DefaultLocalizer = new StudioLocalizer(CultureInfo.GetCultureInfo("en"));

    internal IStudioLocalizer Localizer { get; init; } = DefaultLocalizer;

    public string Label => Localizer.Format("Search.ResultLabel", PageNumber, Snippet);

    internal PdfSearchHit WithLocalizer(IStudioLocalizer localizer) =>
        this with { Localizer = localizer ?? throw new ArgumentNullException(nameof(localizer)) };
}
