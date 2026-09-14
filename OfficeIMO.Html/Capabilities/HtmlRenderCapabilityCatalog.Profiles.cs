namespace OfficeIMO.Html;

public static partial class HtmlRenderCapabilityCatalog {
    private const string HtmlRevision = "3657f01294111708b4c554af11372acdaed777d3";
    private const string DomRevision = "a2331a45360129e8645ef7e0a04740241b6e3726";
    private const string UrlRevision = "8e14777cfa145b08a9fb735fe580ec0c366564c3";
    private const string EncodingRevision = "a985b62a9b45c17da3e17a9f0a0b4e30c34c4a8a";
    private const string CssRevision = "5db94ab2f11278678b3a5014a7f6a92f3c69ff64";
    private const string SvgRevision = "1ecc83537a872592a01ac1e02930555066e56ace";
    private const string MathMlRevision = "f69212ad1f21bbfb7cde24264608c23b96b0e1b7";
    private const string WptRevision = "d63987ce254c38880525061eb87949f3f56716a8";
    private const string Html5LibRevision = "224991ec10db04f056a89eed8b0bd8695fd2950e";

    private static readonly HtmlCapabilityProviderPin[] DocumentProviders = {
        Provider(HtmlCapabilityProviderIds.OfficeIMOHtmlCore, "OfficeIMO.Html.Core", "3.4.3", HtmlCapabilityProviderOwnership.OfficeIMO),
        Provider(HtmlCapabilityProviderIds.AngleSharpHtml, "AngleSharp HTML parser adapter", "AngleSharp 1.7.1 / OfficeIMO.Html.AngleSharp 3.4.3", HtmlCapabilityProviderOwnership.ThirdParty)
    };

    private static readonly HtmlCapabilityProviderPin[] StaticProviders = {
        Provider(HtmlCapabilityProviderIds.OfficeIMOHtmlCore, "OfficeIMO.Html.Core", "3.4.3", HtmlCapabilityProviderOwnership.OfficeIMO),
        Provider(HtmlCapabilityProviderIds.OfficeIMOHtml, "OfficeIMO.Html", "3.4.3", HtmlCapabilityProviderOwnership.OfficeIMO),
        Provider(HtmlCapabilityProviderIds.OfficeIMOCore, "OfficeIMO.Core drawing and text", "3.4.3", HtmlCapabilityProviderOwnership.OfficeIMO),
        Provider(HtmlCapabilityProviderIds.AngleSharpHtml, "AngleSharp HTML parser adapter", "AngleSharp 1.7.1 / OfficeIMO.Html.AngleSharp 3.4.3", HtmlCapabilityProviderOwnership.ThirdParty),
        Provider(HtmlCapabilityProviderIds.AngleSharpCss, "AngleSharp.Css parser", "1.0.1", HtmlCapabilityProviderOwnership.ThirdParty),
        Provider(HtmlCapabilityProviderIds.CallerTextShaper, "Caller-selected text shaper", "caller-selected; optional", HtmlCapabilityProviderOwnership.Caller)
    };

    private static readonly HtmlCapabilityProviderPin[] PagedProviders = StaticProviders.Concat(new[] {
        Provider(HtmlCapabilityProviderIds.OfficeIMOHtmlPdf, "OfficeIMO.Html.Pdf", "3.4.3", HtmlCapabilityProviderOwnership.OfficeIMO),
        Provider(HtmlCapabilityProviderIds.OfficeIMOPdf, "OfficeIMO.Pdf", "3.4.3", HtmlCapabilityProviderOwnership.OfficeIMO)
    }).ToArray();

    private static readonly HtmlCapabilitySpecificationPin[] DocumentSpecifications = {
        Specification(HtmlCapabilitySpecificationIds.Html, "WHATWG HTML Living Standard", "https://html.spec.whatwg.org/commit-snapshots/" + HtmlRevision + "/", HtmlRevision, "syntax, parsing, elements, document modes, resources, forms, and static active-content boundaries"),
        Specification(HtmlCapabilitySpecificationIds.Dom, "WHATWG DOM", "https://dom.spec.whatwg.org/commit-snapshots/" + DomRevision + "/", DomRevision, "owned node, tree, mutation, and query contracts selected by web-document-v1"),
        Specification(HtmlCapabilitySpecificationIds.Url, "WHATWG URL", "https://url.spec.whatwg.org/commit-snapshots/" + UrlRevision + "/", UrlRevision, "URL parsing and resolution used by document and resource processing"),
        Specification(HtmlCapabilitySpecificationIds.Encoding, "WHATWG Encoding", "https://encoding.spec.whatwg.org/commit-snapshots/" + EncodingRevision + "/", EncodingRevision, "labels, sniffing precedence, decoding, and error behavior selected by web-document-v1")
    };

    private static readonly HtmlCapabilitySpecificationPin[] StaticSpecifications = DocumentSpecifications.Concat(new[] {
        Specification(HtmlCapabilitySpecificationIds.CssSnapshot, "CSS Snapshot 2026", "https://www.w3.org/TR/2026/NOTE-css-2026-20260622/", "22 June 2026", "module index only; capability rows select exact modules and subsets"),
        CssModule(HtmlCapabilitySpecificationIds.CssCascade, "CSS cascade, variables, and conditional rules", "css-cascade, css-variables, css-conditional, css-nesting"),
        CssModule(HtmlCapabilitySpecificationIds.Selectors, "Selectors", "selectors"),
        CssModule(HtmlCapabilitySpecificationIds.CssValues, "CSS values, units, and sizing", "css-values, css-sizing"),
        CssModule(HtmlCapabilitySpecificationIds.CssColor, "CSS color", "css-color"),
        CssModule(HtmlCapabilitySpecificationIds.CssBackgrounds, "CSS backgrounds, borders, and shadows", "css-backgrounds, css-borders, css-ui"),
        CssModule(HtmlCapabilitySpecificationIds.CssImages, "CSS images", "css-images"),
        CssModule(HtmlCapabilitySpecificationIds.CssEffects, "CSS transforms, masking, filters, and compositing", "css-transforms, css-masking, filter-effects, compositing"),
        CssModule(HtmlCapabilitySpecificationIds.CssBoxLayout, "CSS display, box, sizing, and inline layout", "css-display, css-box, css-sizing, css-inline"),
        CssModule(HtmlCapabilitySpecificationIds.CssText, "CSS text, overflow, and text decoration", "css-text, css-text-decor, css-overflow"),
        CssModule(HtmlCapabilitySpecificationIds.CssWritingModes, "CSS writing modes and ruby", "css-writing-modes, css-ruby, css-pseudo"),
        CssModule(HtmlCapabilitySpecificationIds.CssFlexbox, "CSS flexible box layout", "css-flexbox, css-align"),
        CssModule(HtmlCapabilitySpecificationIds.CssGrid, "CSS grid layout", "css-grid, css-align"),
        CssModule(HtmlCapabilitySpecificationIds.CssMulticol, "CSS multi-column layout", "css-multicol"),
        CssModule(HtmlCapabilitySpecificationIds.CssPosition, "CSS positioned layout", "css-position, css-position-3"),
        CssModule(HtmlCapabilitySpecificationIds.CssTables, "CSS table layout", "css-tables"),
        CssModule(HtmlCapabilitySpecificationIds.CssGeneratedContent, "CSS generated content, lists, and counters", "css-content, css-lists, css-counter-styles, css-gcpm"),
        CssModule(HtmlCapabilitySpecificationIds.CssPagedMedia, "CSS paged media and fragmentation", "css-page, css-break, css-gcpm"),
        CssModule(HtmlCapabilitySpecificationIds.CssConditional, "Media, conditional, and container queries", "mediaqueries, css-conditional, css-contain"),
        CssModule(HtmlCapabilitySpecificationIds.CssFonts, "CSS fonts", "css-fonts"),
        Specification(HtmlCapabilitySpecificationIds.Unicode17, "Unicode Standard 17.0.0", "https://www.unicode.org/versions/Unicode17.0.0/", "17.0.0", "bidirectional text, line breaking, character properties, and generated Unicode data used by the selected text subset"),
        Specification(HtmlCapabilitySpecificationIds.Svg2, "SVG 2 editor source", "https://github.com/w3c/svgwg/tree/" + SvgRevision, SvgRevision, "bounded static SVG geometry, paint, text, references, masks, and filters declared by the SVG capability rows"),
        Specification(HtmlCapabilitySpecificationIds.MathMlCore, "MathML Core editor source", "https://github.com/w3c/mathml-core/tree/" + MathMlRevision, MathMlRevision, "bounded Presentation MathML structures declared by the MathML capability rows")
    }).ToArray();

    private static readonly HtmlCapabilitySpecificationPin[] PagedSpecifications = StaticSpecifications.Concat(new[] {
        Specification(HtmlCapabilitySpecificationIds.Pdf20, "ISO 32000-2:2020 PDF 2.0", "https://www.iso.org/standard/75839.html", "ISO 32000-2:2020", "metadata, tagged structure, annotations, forms, page geometry, and searchable content emitted by the declared PDF output subset")
    }).ToArray();

    private static readonly HtmlCapabilityEvidencePin[] DocumentEvidence = {
        Evidence(HtmlCapabilityEvidenceIds.OfficeIMOHtmlTests, "OfficeIMO.Html.Tests executable contract suite", "package RepositoryCommit; catalog schema 2", HtmlCapabilityEvidenceRole.Regression, "owned document, provider, conversion, resource, rendering, output, and catalog contracts selected by each capability row"),
        Evidence(HtmlCapabilityEvidenceIds.DocumentV1, "OfficeIMO H1-H2 owned web-document qualification selection", "H1-H2/v1", HtmlCapabilityEvidenceRole.Qualification, "selected executable byte decoding, parsing, serialization, query, mutation, identity, foreign-attribute, and source-position cases", required: 4, passed: 4, failed: 0, excluded: 0, untested: 0, caseIds: new[] { "ForeignAttributesAndSourcePositionsSurviveEdits", "HtmlConversionDocument_LoadDetectsMetaCharsetFromByteInput", "RecoveredDoctypeSurvivesEditingAndSerialization", "Snapshot_EditPreservesIdentityWithoutMutatingSourceOrLeakedHandles" }),
        Evidence(HtmlCapabilityEvidenceIds.Html5Lib, "html5lib-tests", Html5LibRevision, HtmlCapabilityEvidenceRole.ConformanceReference, "pinned upstream parser corpus; individual selected paths and counts remain unqualified until H5 manifests adopt them"),
        Evidence(HtmlCapabilityEvidenceIds.WebPlatformTests, "web-platform-tests", WptRevision, HtmlCapabilityEvidenceRole.ConformanceReference, "pinned upstream source; only explicitly selected paths and counts can qualify a capability")
    };

    private static readonly HtmlCapabilityEvidencePin[] StaticScreenEvidence = DocumentEvidence.Concat(new[] {
        Evidence(HtmlCapabilityEvidenceIds.H4ScreenV1, "OfficeIMO H4 continuous-mode HTML rendering corpus", "H4/v1", HtmlCapabilityEvidenceRole.Qualification, "continuous-mode semantic, geometry, raster, SVG, provider, and environment evidence", required: 6, passed: 6, failed: 0, excluded: 0, untested: 0, caseIds: new[] { "dashboard-print", "email-render", "invoice", "multilingual-bidi", "product-catalog", "quarterly-report" })
    }).ToArray();

    private static readonly HtmlCapabilityEvidencePin[] PagedEvidence = DocumentEvidence.Concat(new[] {
        Evidence(HtmlCapabilityEvidenceIds.H4PagedV1, "OfficeIMO H4 paged-mode HTML rendering corpus", "H4/v1", HtmlCapabilityEvidenceRole.Qualification, "paged-mode semantic, geometry, raster, SVG, searchable-PDF, provider, and environment evidence", required: 5, passed: 5, failed: 0, excluded: 0, untested: 0, caseIds: new[] { "account-statement", "business-letter", "certificate", "legal-contract", "static-standards-showcase" })
    }).ToArray();

    private static readonly IReadOnlyList<HtmlCapabilityProfileManifest> Profiles = new[] {
        new HtmlCapabilityProfileManifest(
            HtmlCapabilityProfileIds.WebDocumentV1,
            "1.0",
            "Web document v1",
            HtmlCapabilityPromotionState.QualifiedOptIn,
            DocumentProviders,
            DocumentSpecifications,
            DocumentEvidence,
            new[] { "Windows", "Linux", "macOS", "browser WebAssembly where separately qualified" },
            new[] { "owned document", "owned snapshot", "serialized HTML" }),
        new HtmlCapabilityProfileManifest(
            HtmlCapabilityProfileIds.StaticScreenV1,
            "1.0",
            "Static screen v1",
            HtmlCapabilityPromotionState.StableDefault,
            StaticProviders,
            StaticSpecifications,
            StaticScreenEvidence,
            new[] { "Windows", "Linux", "macOS" },
            new[] { "geometry", "display list", "raster", "SVG" }),
        new HtmlCapabilityProfileManifest(
            HtmlCapabilityProfileIds.PagedPrintV1,
            "1.0",
            "Paged print v1",
            HtmlCapabilityPromotionState.StableDefault,
            PagedProviders,
            PagedSpecifications,
            PagedEvidence,
            new[] { "Windows", "Linux", "macOS" },
            new[] { "page set", "paged raster", "paged SVG", "searchable PDF" })
    }.OrderBy(profile => profile.Id, StringComparer.Ordinal).ToList().AsReadOnly();

    private static readonly IReadOnlyDictionary<string, HtmlCapabilityProfileManifest> ProfilesById =
        Profiles.ToDictionary(profile => profile.Id, StringComparer.OrdinalIgnoreCase);

    /// <summary>Gets versioned provider, specification, evidence, platform, and output manifests in stable order.</summary>
    public static IReadOnlyList<HtmlCapabilityProfileManifest> ProfileManifests => Profiles;

    /// <summary>Gets a compatibility profile manifest by stable identifier.</summary>
    public static HtmlCapabilityProfileManifest GetProfile(string id) {
        if (!TryGetProfile(id, out HtmlCapabilityProfileManifest profile)) {
            throw new ArgumentOutOfRangeException(nameof(id), id, "Unknown HTML compatibility profile.");
        }
        return profile;
    }

    /// <summary>Attempts to get a compatibility profile manifest by stable identifier.</summary>
    public static bool TryGetProfile(string? id, out HtmlCapabilityProfileManifest profile) {
        if (!string.IsNullOrWhiteSpace(id)
            && ProfilesById.TryGetValue(id!.Trim(), out HtmlCapabilityProfileManifest? found)
            && found != null) {
            profile = found;
            return true;
        }
        profile = null!;
        return false;
    }

    private static HtmlCapabilityProviderPin Provider(string id, string name, string version, HtmlCapabilityProviderOwnership ownership) =>
        new HtmlCapabilityProviderPin(id, name, version, ownership);

    private static HtmlCapabilitySpecificationPin Specification(string id, string title, string uri, string revision, string scope) =>
        new HtmlCapabilitySpecificationPin(id, title, uri, revision, scope);

    private static HtmlCapabilitySpecificationPin CssModule(string id, string title, string modules) =>
        Specification(id, title, "https://github.com/w3c/csswg-drafts/tree/" + CssRevision, CssRevision, modules);

    private static HtmlCapabilityEvidencePin Evidence(
        string id,
        string source,
        string revision,
        HtmlCapabilityEvidenceRole role,
        string scope,
        int? required = null,
        int? passed = null,
        int? failed = null,
        int? excluded = null,
        int? untested = null,
        IEnumerable<string>? caseIds = null) =>
        new HtmlCapabilityEvidencePin(id, source, revision, role, scope, required, passed, failed, excluded, untested, caseIds);
}
