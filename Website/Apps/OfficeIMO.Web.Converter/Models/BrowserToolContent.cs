namespace OfficeIMO.Web.Converter.Models;

public sealed record BrowserToolStep(string Title, string Description);

public sealed record BrowserToolContent(
    string Title,
    string Summary,
    string Input,
    string Output,
    string Expectation,
    string GuideUrl,
    string GuideLabel,
    string SeoTitle,
    IReadOnlyList<BrowserToolStep> Steps);
