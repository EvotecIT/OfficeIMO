using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Html.Pdf;

/// <summary>
/// Describes the product contract for one HTML to PDF adapter profile.
/// </summary>
public sealed class HtmlPdfProfileContract {
    internal HtmlPdfProfileContract(HtmlPdfProfile profile, string id, string displayName, string pipeline, string intendedUse, string fidelityContract, string unsupportedScope) {
        Profile = profile;
        Id = id;
        DisplayName = displayName;
        Pipeline = pipeline;
        IntendedUse = intendedUse;
        FidelityContract = fidelityContract;
        UnsupportedScope = unsupportedScope;
    }

    /// <summary>Adapter profile represented by this contract.</summary>
    public HtmlPdfProfile Profile { get; }

    /// <summary>Stable profile identifier for wrappers, manifests, and documentation.</summary>
    public string Id { get; }

    /// <summary>Human-readable profile name.</summary>
    public string DisplayName { get; }

    /// <summary>First-party conversion pipeline used by this profile.</summary>
    public string Pipeline { get; }

    /// <summary>Recommended use cases for this profile.</summary>
    public string IntendedUse { get; }

    /// <summary>Fidelity and output contract callers can rely on.</summary>
    public string FidelityContract { get; }

    /// <summary>Known unsupported or intentionally simplified scope.</summary>
    public string UnsupportedScope { get; }
}

/// <summary>
/// Stable contract descriptions for the HTML to PDF adapter profiles.
/// </summary>
public static class HtmlPdfProfileContracts {
    private static readonly IReadOnlyList<HtmlPdfProfileContract> Contracts = new ReadOnlyCollection<HtmlPdfProfileContract>(new[] {
        new HtmlPdfProfileContract(
            HtmlPdfProfile.Semantic,
            "html-pdf-semantic",
            "Semantic HTML to PDF",
            "HTML -> OfficeIMO.Markdown.Html -> MarkdownDoc -> OfficeIMO.Markdown.Pdf -> OfficeIMO.Pdf",
            "Articles, documentation, simple reports, and semantic HTML where clean structure matters more than authored CSS layout.",
            "Preserves headings, paragraphs, lists, links, tables, local/data images supported by the Markdown PDF path, and shared conversion warnings.",
            "Not a browser renderer; CSS layout, scripts, forms, media, complex paged media, and unsupported resources are simplified or reported through conversion diagnostics."),
        new HtmlPdfProfileContract(
            HtmlPdfProfile.Document,
            "html-pdf-document",
            "Document HTML to PDF",
            "HTML -> OfficeIMO.Word.Html -> WordDocument -> OfficeIMO.Word.Pdf -> OfficeIMO.Pdf",
            "Trusted or controlled print-oriented HTML with CSS, images, links, tables, and page-break hints where the Word model is the best intermediate document shape.",
            "Preserves the Word HTML adapter's supported document structure, styling, images, links, tables, page-break hints, and first-party Word PDF diagnostics.",
            "Not a browser renderer; unsupported CSS, script behavior, media, interactive HTML, and complex layout are simplified according to Word HTML and Word PDF adapter diagnostics.")
    });

    /// <summary>All supported HTML to PDF profile contracts.</summary>
    public static IReadOnlyList<HtmlPdfProfileContract> All => Contracts;

    /// <summary>Gets the contract for a supported HTML to PDF profile.</summary>
    public static HtmlPdfProfileContract Get(HtmlPdfProfile profile) {
        for (int i = 0; i < Contracts.Count; i++) {
            if (Contracts[i].Profile == profile) {
                return Contracts[i];
            }
        }

        throw new ArgumentOutOfRangeException(nameof(profile), profile, "HTML to PDF profile is not supported.");
    }
}
