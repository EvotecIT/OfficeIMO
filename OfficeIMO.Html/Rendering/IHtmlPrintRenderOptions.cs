namespace OfficeIMO.Html;

/// <summary>
/// Implemented by render options from the optional print backend (OfficeIMO.Html.Pdf).
/// The provider-neutral engine asks these options to prepare print font measurement,
/// font retention and embedded-package resource policy; plain options keep neutral layout.
/// </summary>
internal interface IHtmlPrintRenderOptions {
    /// <summary>
    /// True when the print resource policy governs resources. False for shared options wrapped for image
    /// export, whose own URL and resource policy, including archive and container resolvers, stays authoritative.
    /// </summary>
    bool UsesPrintResourcePolicy { get; }

    /// <summary>Resolves options for PDF-backed encoders using the print font and resource policy.</summary>
    HtmlRenderOptions PrepareRenderOptions(HtmlRenderRequest request);

    /// <summary>Retains explicitly configured default faces for shared layout and paint.</summary>
    IReadOnlyList<string> RetainConfiguredDefaultFonts(OfficeFontFaceCollection fonts, HtmlResourceSession resources,
        HtmlDiagnosticReport diagnostics, CancellationToken cancellationToken);

    /// <summary>True when document font families may be resolved from installed system fonts.</summary>
    bool AllowsDocumentSystemFontFamilies { get; }

    /// <summary>Freezes the library-selected Unicode fallback before authoritative layout.</summary>
    bool TryPrepareLibraryFontMetrics(IEnumerable<HtmlRenderVisual> flow,
        IEnumerable<(string Text, OfficeFontInfo Font)> marginRequirements, OfficeFontFaceCollection fonts,
        HtmlResourceSession resources, HtmlDiagnosticReport diagnostics, CancellationToken cancellationToken);

    /// <summary>Applies the embedded-package policy of a site bundle to a print request.</summary>
    HtmlRenderRequest PrepareSiteBundleRequest(HtmlSiteBundle bundle, HtmlRenderRequest request);

    /// <summary>True when archive-owned resources may be read during print rendering.</summary>
    bool AllowsEmbeddedPackageResources { get; }

    /// <summary>Archive resolver admitted by the embedded-package policy.</summary>
    HtmlRenderResourceResolver? EmbeddedPackageResourceResolver { get; set; }

    /// <summary>URL policy for host resources resolved next to archive resources.</summary>
    HtmlUrlPolicy? EmbeddedPackageHostResourceUrlPolicy { get; set; }

    /// <summary>Resolver configured by the archive owner, recognised when it is reused.</summary>
    HtmlRenderResourceResolver? EmbeddedPackageConfiguredResourceResolver { get; set; }

    /// <summary>Host-only resolver used after the package permission check.</summary>
    HtmlRenderResourceResolver? EmbeddedPackageHostResourceResolver { get; set; }
}
