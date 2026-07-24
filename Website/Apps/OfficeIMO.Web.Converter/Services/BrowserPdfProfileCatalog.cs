using OfficeIMO.Web.Converter.Models;

namespace OfficeIMO.Web.Converter.Services;

public static class BrowserPdfProfileCatalog {
    public static BrowserPdfProfile Faithful { get; } = new(
        BrowserPdfProfileKind.Faithful,
        "faithful",
        "Faithful",
        "Best browser fidelity with embedded fonts, full shaping, semantic tags, and explicit substitutions.");

    public static BrowserPdfProfile Portable { get; } = new(
        BrowserPdfProfileKind.Portable,
        "portable",
        "Portable",
        "Pinned browser assets and deterministic output with no system-font or external-resource dependency.");

    public static BrowserPdfProfile Accessible { get; } = new(
        BrowserPdfProfileKind.Accessible,
        "accessible",
        "Accessible",
        "Tagged output plus PDF/UA-1 identification groundwork; the report does not claim validator conformance.");

    public static BrowserPdfProfile Diagnostic { get; } = new(
        BrowserPdfProfileKind.Diagnostic,
        "diagnostic",
        "Diagnostic",
        "Faithful output plus a first-page SVG overlay for words, lines, regions, and reading order.");

    public static IReadOnlyList<BrowserPdfProfile> All { get; } = [
        Faithful,
        Portable,
        Accessible,
        Diagnostic
    ];

    public static BrowserPdfProfile Find(string? id) =>
        All.FirstOrDefault(profile => string.Equals(profile.Id, id, StringComparison.OrdinalIgnoreCase))
        ?? Faithful;
}
