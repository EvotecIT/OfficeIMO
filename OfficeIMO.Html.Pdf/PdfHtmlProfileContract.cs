using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Html.Pdf;

/// <summary>
/// Describes the product contract for one PDF to HTML adapter profile.
/// </summary>
public sealed class PdfHtmlProfileContract {
    internal PdfHtmlProfileContract(PdfHtmlProfile profile, string id, string displayName, string pipeline, string intendedUse, string fidelityContract, string unsupportedScope) {
        Profile = profile;
        Id = id;
        DisplayName = displayName;
        Pipeline = pipeline;
        IntendedUse = intendedUse;
        FidelityContract = fidelityContract;
        UnsupportedScope = unsupportedScope;
    }

    /// <summary>Adapter profile represented by this contract.</summary>
    public PdfHtmlProfile Profile { get; }

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
/// Stable contract descriptions for the PDF to HTML adapter profiles.
/// </summary>
public static class PdfHtmlProfileContracts {
    private static readonly IReadOnlyList<PdfHtmlProfileContract> Contracts = new ReadOnlyCollection<PdfHtmlProfileContract>(new[] {
        new PdfHtmlProfileContract(
            PdfHtmlProfile.Semantic,
            "pdf-html-semantic",
            "Semantic PDF to HTML",
            "PDF -> OfficeIMO.Pdf logical model -> semantic HTML",
            "Search, indexing, review, export, and pipeline flows where readable structure matters more than exact visual positioning.",
            "Emits headings, paragraphs, lists, detected tables, metadata, image placeholders or data URI images when available, and optional link/form sections.",
            "Born-digital parser-supported PDFs work best; scanned PDFs need an OCR adapter, and complex/unsafe PDF structures remain governed by OfficeIMO.Pdf read diagnostics."),
        new PdfHtmlProfileContract(
            PdfHtmlProfile.PositionedReview,
            "pdf-html-positioned-review",
            "Positioned Review PDF to HTML",
            "PDF -> OfficeIMO.Pdf logical model -> page wrappers with positioned review hints",
            "Human review and diagnostics where page coordinates, text blocks, table hints, links, forms, and image placements need to be inspected in a browser.",
            "Emits one page wrapper per source page with positioned text/table/image/link/form hints and safe handling for unsafe links.",
            "Review output is not a full PDF renderer; complex graphics, arbitrary content streams, optional content, scans, and unsupported parser structures are simplified or reported.")
    });

    /// <summary>All supported PDF to HTML profile contracts.</summary>
    public static IReadOnlyList<PdfHtmlProfileContract> All => Contracts;

    /// <summary>Gets the contract for a supported PDF to HTML profile.</summary>
    public static PdfHtmlProfileContract Get(PdfHtmlProfile profile) {
        for (int i = 0; i < Contracts.Count; i++) {
            if (Contracts[i].Profile == profile) {
                return Contracts[i];
            }
        }

        throw new ArgumentOutOfRangeException(nameof(profile), profile, "PDF to HTML profile is not supported.");
    }
}
