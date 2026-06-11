using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Html.Pdf;

/// <summary>
/// Describes the product contract for one PDF to HTML adapter profile.
/// </summary>
public sealed class PdfHtmlProfileContract {
    internal PdfHtmlProfileContract(
        PdfHtmlProfile profile,
        string id,
        string displayName,
        string pipeline,
        string intendedUse,
        string fidelityContract,
        string unsupportedScope,
        IReadOnlyList<string> preservedSignals,
        IReadOnlyList<string> reviewSignals,
        IReadOnlyList<string> outputArtifacts,
        IReadOnlyList<string> diagnosticGuarantees,
        IReadOnlyList<string> rendererBoundaries) {
        Profile = profile;
        Id = id;
        DisplayName = displayName;
        Pipeline = pipeline;
        IntendedUse = intendedUse;
        FidelityContract = fidelityContract;
        UnsupportedScope = unsupportedScope;
        PreservedSignals = preservedSignals;
        ReviewSignals = reviewSignals;
        OutputArtifacts = outputArtifacts;
        DiagnosticGuarantees = diagnosticGuarantees;
        RendererBoundaries = rendererBoundaries;
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

    /// <summary>Logical PDF signals this profile treats as part of its supported v1 contract.</summary>
    public IReadOnlyList<string> PreservedSignals { get; }

    /// <summary>Review-oriented diagnostics or geometry hints this profile can expose.</summary>
    public IReadOnlyList<string> ReviewSignals { get; }

    /// <summary>Output artifact shapes supported by this profile.</summary>
    public IReadOnlyList<string> OutputArtifacts { get; }

    /// <summary>Diagnostics callers can rely on when PDF content is simplified, rejected, or degraded.</summary>
    public IReadOnlyList<string> DiagnosticGuarantees { get; }

    /// <summary>Explicit non-goals that prevent callers from treating this adapter as a full PDF renderer or editable reconstruction engine.</summary>
    public IReadOnlyList<string> RendererBoundaries { get; }
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
            "Born-digital parser-supported PDFs work best; scanned PDFs need an OCR adapter, and complex/unsafe PDF structures remain governed by OfficeIMO.Pdf read diagnostics.",
            new[] { "metadata", "headings", "paragraphs", "lists", "tables", "images", "links", "form-fields" },
            new[] { "page-numbers", "image-placeholders", "optional-link-sections", "optional-form-sections" },
            new[] { "html", "export-summary" },
            new[] { "conversion-report-warnings", "image-embedding-policy-warnings" },
            new[] { "no-ocr", "no-pixel-perfect-rendering", "no-editable-office-reconstruction" }),
        new PdfHtmlProfileContract(
            PdfHtmlProfile.PositionedReview,
            "pdf-html-positioned-review",
            "Positioned Review PDF to HTML",
            "PDF -> OfficeIMO.Pdf logical model -> page wrappers with positioned review hints",
            "Human review and diagnostics where page coordinates, text blocks, table hints, links, forms, and image placements need to be inspected in a browser.",
            "Emits one page wrapper per source page with positioned text/table/image/link/form hints and safe handling for unsafe links.",
            "Review output is not a full PDF renderer; complex graphics, arbitrary content streams, optional content, scans, and unsupported parser structures are simplified or reported.",
            new[] { "page-geometry", "text-blocks", "tables", "images", "links", "form-widgets" },
            new[] { "absolute-text-positions", "table-bounds", "image-placements", "link-frames", "form-widget-frames", "unsafe-link-inertness" },
            new[] { "html", "export-summary" },
            new[] { "conversion-report-warnings", "image-embedding-policy-warnings", "unsafe-link-sanitization" },
            new[] { "no-full-graphics-renderer", "no-optional-content-composition", "no-scan-ocr", "no-editable-office-reconstruction" })
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
