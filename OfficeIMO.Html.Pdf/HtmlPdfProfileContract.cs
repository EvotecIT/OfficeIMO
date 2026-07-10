using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Html.Pdf;

/// <summary>
/// Describes the product contract for one HTML to PDF adapter profile.
/// </summary>
public sealed class HtmlPdfProfileContract {
    internal HtmlPdfProfileContract(
        HtmlPdfProfile profile,
        HtmlConversionProfile sharedProfile,
        string id,
        string displayName,
        string pipeline,
        string intendedUse,
        string fidelityContract,
        string unsupportedScope,
        IReadOnlyList<string> supportedHtmlFeatures,
        IReadOnlyList<string> supportedCssFeatures,
        IReadOnlyList<string> supportedResourceFeatures,
        IReadOnlyList<string> diagnosticGuarantees,
        IReadOnlyList<string> rendererBoundaries) {
        Profile = profile;
        SharedProfile = sharedProfile;
        Id = id;
        DisplayName = displayName;
        Pipeline = pipeline;
        IntendedUse = intendedUse;
        FidelityContract = fidelityContract;
        UnsupportedScope = unsupportedScope;
        SupportedHtmlFeatures = supportedHtmlFeatures;
        SupportedCssFeatures = supportedCssFeatures;
        SupportedResourceFeatures = supportedResourceFeatures;
        DiagnosticGuarantees = diagnosticGuarantees;
        RendererBoundaries = rendererBoundaries;
    }

    /// <summary>Adapter profile represented by this contract.</summary>
    public HtmlPdfProfile Profile { get; }

    /// <summary>Shared OfficeIMO HTML conversion profile represented by this adapter profile.</summary>
    public HtmlConversionProfile SharedProfile { get; }

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

    /// <summary>HTML structures this profile treats as part of its supported v1 contract.</summary>
    public IReadOnlyList<string> SupportedHtmlFeatures { get; }

    /// <summary>CSS features this profile treats as part of its supported v1 contract.</summary>
    public IReadOnlyList<string> SupportedCssFeatures { get; }

    /// <summary>Resource loading and embedding behavior this profile exposes as a supported v1 contract.</summary>
    public IReadOnlyList<string> SupportedResourceFeatures { get; }

    /// <summary>Diagnostics callers can rely on when content is simplified, rejected, or degraded.</summary>
    public IReadOnlyList<string> DiagnosticGuarantees { get; }

    /// <summary>Explicit non-goals that prevent callers from treating this adapter as a browser-grade renderer.</summary>
    public IReadOnlyList<string> RendererBoundaries { get; }
}

/// <summary>
/// Stable contract descriptions for the HTML to PDF adapter profiles.
/// </summary>
public static class HtmlPdfProfileContracts {
    private static readonly IReadOnlyList<HtmlPdfProfileContract> Contracts = new ReadOnlyCollection<HtmlPdfProfileContract>(new[] {
        new HtmlPdfProfileContract(
            HtmlPdfProfile.Semantic,
            HtmlConversionProfile.Semantic,
            "html-pdf-semantic",
            "Semantic HTML to PDF",
            "HTML -> OfficeIMO.Markdown.Html -> MarkdownDoc -> OfficeIMO.Markdown.Pdf -> OfficeIMO.Pdf",
            "Articles, documentation, simple reports, and semantic HTML where clean structure matters more than authored CSS layout.",
            "Preserves headings, paragraphs, lists, links, tables, local/data images supported by the Markdown PDF path, and shared conversion warnings.",
            "Not a browser renderer; CSS layout, scripts, forms, media, complex paged media, and unsupported resources are simplified or reported through conversion diagnostics.",
            new[] { "headings", "paragraphs", "lists", "links", "tables", "local-and-data-images" },
            new[] { "table-cell-alignment", "table-cell-background", "inline-emphasis", "inline-color" },
            new[] { "local-image-files", "data-uri-images", "markdown-resource-policy" },
            new[] { "shared-markdown-pdf-warnings", "unsupported-image-warning" },
            new[] { "no-browser-layout-engine", "no-script-execution", "no-interactive-html", "no-paged-media-engine" }),
        new HtmlPdfProfileContract(
            HtmlPdfProfile.Document,
            HtmlConversionProfile.Document,
            "html-pdf-document",
            "Document HTML to PDF",
            "HTML -> OfficeIMO.Word.Html -> WordDocument -> OfficeIMO.Word.Pdf -> OfficeIMO.Pdf",
            "Trusted or controlled print-oriented HTML with CSS, images, links, tables, and page-break hints where the Word model is the best intermediate document shape.",
            "Preserves the Word HTML adapter's supported document structure, styling, images, links, tables, page-break hints, and first-party Word PDF diagnostics.",
            "Not a browser renderer; unsupported CSS, script behavior, media, interactive HTML, and complex layout are simplified according to Word HTML and Word PDF adapter diagnostics.",
            new[] { "headings", "paragraphs", "lists", "links", "tables", "images", "page-break-hints" },
            new[] { "inline-css", "linked-stylesheets-when-enabled", "table-styles", "font-styles", "colors", "page-break-before-after" },
            new[] { "base-path-resolution", "trusted-stylesheet-policy", "trusted-image-policy", "resource-policy-summary" },
            new[] { "html-import-diagnostics", "word-pdf-warnings", "shared-conversion-report" },
            new[] { "no-browser-layout-engine", "no-script-execution", "no-css-grid-or-flex-layout-contract", "no-interactive-html" }),
        new HtmlPdfProfileContract(
            HtmlPdfProfile.Rendered,
            HtmlConversionProfile.HighFidelityPrint,
            "html-pdf-rendered",
            "Direct Rendered HTML to PDF",
            "HTML -> OfficeIMO.Html layout -> shared render visuals -> OfficeIMO.Pdf",
            "Authored print HTML that needs direct CSS geometry, searchable text, links, images, and the same layout used by HTML image export.",
            "Uses the first-party paged renderer directly and reports every currently unsupported or approximated layout feature through stable diagnostics.",
            "The first implementation contract covers normal flow, inline text, basic boxes, data and asynchronously resolved external images, simple tables, generic page rules, and stable line/row fragmentation; flex, grid, advanced positioning and fragmentation, linked stylesheets, font loading, and complex shaping remain diagnostic-backed roadmap items.",
            new[] { "headings", "paragraphs", "lists", "links", "simple-tables", "data-uri-images", "explicit-page-breaks" },
            new[] { "cascade-and-inheritance", "custom-properties", "print-media", "font-and-color", "margin-padding-border-background", "width-height", "text-alignment", "generic-page-size-margin", "page-break-before-after" },
            new[] { "data-uri-images", "caller-resolved-external-images", "base-uri-link-resolution", "shared-url-policy", "resource-timeout-and-byte-budgets", "resource-manifest-diagnostics" },
            new[] { "shared-html-render-diagnostics", "resource-policy-diagnostics", "pending-layout-feature-diagnostics", "shared-conversion-report" },
            new[] { "no-script-execution", "no-browser-process", "no-new-external-dependencies", "advanced-layout-remains-explicitly-diagnosed" })
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
