namespace OfficeIMO.Html;

/// <summary>Describes how much of a declared capability subset has qualification evidence.</summary>
public enum HtmlCapabilityCoverage {
    /// <summary>The exact declared subset has the required evidence for the selected profile.</summary>
    Qualified,
    /// <summary>Only part of the declared feature family is implemented and qualified.</summary>
    Partial,
    /// <summary>The declared feature is outside the supported subset.</summary>
    Unsupported,
    /// <summary>The implementation exists but has not passed the profile's qualification gate.</summary>
    Unqualified
}

/// <summary>Describes what the engine does when it encounters a declared capability.</summary>
public enum HtmlCapabilityHandling {
    /// <summary>The selected OfficeIMO pipeline handles the declared subset directly.</summary>
    Native,
    /// <summary>The engine preserves the source or semantic value without applying its behavior.</summary>
    Preserved,
    /// <summary>The engine applies a documented substitute or initial-value behavior.</summary>
    Fallback,
    /// <summary>The engine deliberately omits the feature and reports the boundary.</summary>
    Ignored,
    /// <summary>The engine blocks the feature or resource under an explicit policy.</summary>
    Rejected
}

/// <summary>Describes whether a capability is required by, optional in, or experimental for a profile.</summary>
public enum HtmlCapabilityMaturity {
    /// <summary>The profile requires this outcome.</summary>
    Required,
    /// <summary>The profile can use this outcome when explicitly requested or available.</summary>
    Optional,
    /// <summary>The outcome is still experimental and can change within its preview contract.</summary>
    Experimental
}

/// <summary>Describes how a capability is promoted into released OfficeIMO profiles.</summary>
public enum HtmlCapabilityPromotionState {
    /// <summary>The capability is available only in integration builds and internal evidence runs.</summary>
    Incubating,
    /// <summary>The capability is available through an explicit preview profile, option, or provider.</summary>
    ExperimentalOptIn,
    /// <summary>The capability passed its declared gate and is released for explicit selection.</summary>
    QualifiedOptIn,
    /// <summary>The capability is part of the normal supported profile.</summary>
    StableDefault,
    /// <summary>The capability was replaced and remains only under a versioned compatibility policy.</summary>
    Superseded
}

/// <summary>Identifies the processing stages to which a capability claim applies.</summary>
[Flags]
public enum HtmlCapabilityStage {
    /// <summary>No processing stage has been declared.</summary>
    None = 0,
    /// <summary>Byte decoding, encoding selection, URL input, and source retention.</summary>
    SourceAndDecoding = 1,
    /// <summary>Syntax recognition, recovery, source preservation, and serialization.</summary>
    ParseAndPreserve = 2,
    /// <summary>Owned DOM nodes, namespaces, query, and mutation behavior.</summary>
    DomAndQuery = 4,
    /// <summary>CSS cascade, specified values, inheritance, and computed values.</summary>
    CascadeAndCompute = 8,
    /// <summary>Box construction, intrinsic sizing, line layout, geometry, and fragmentation.</summary>
    Layout = 16,
    /// <summary>Display-list construction and target output behavior.</summary>
    PaintAndOutput = 32,
    /// <summary>Script, events, navigation, web APIs, and interaction behavior.</summary>
    RuntimeAndInteraction = 64
}

/// <summary>Classifies who owns a provider implementation used by a profile.</summary>
public enum HtmlCapabilityProviderOwnership {
    /// <summary>The implementation is owned by OfficeIMO.</summary>
    OfficeIMO,
    /// <summary>The implementation is supplied by a replaceable third-party package.</summary>
    ThirdParty,
    /// <summary>The implementation is supplied by the platform or runtime.</summary>
    Platform,
    /// <summary>The implementation is supplied explicitly by the caller.</summary>
    Caller
}

/// <summary>Classifies how an evidence source contributes to a compatibility claim.</summary>
public enum HtmlCapabilityEvidenceRole {
    /// <summary>The evidence is a required passing qualification gate.</summary>
    Qualification,
    /// <summary>The evidence protects an established behavior against regression.</summary>
    Regression,
    /// <summary>The evidence is pinned for conformance selection but does not yet qualify the whole source.</summary>
    ConformanceReference,
    /// <summary>The evidence compares providers or independent engines without defining the contract alone.</summary>
    DifferentialReference
}

/// <summary>Stable identifiers for the first versioned HTML compatibility profiles.</summary>
public static class HtmlCapabilityProfileIds {
    /// <summary>Owned web-document parsing, DOM, query, edit, and serialization profile.</summary>
    public const string WebDocumentV1 = "web-document-v1";
    /// <summary>Static screen-media cascade, layout, paint, resource, and output profile.</summary>
    public const string StaticScreenV1 = "static-screen-v1";
    /// <summary>Print-media cascade, paged layout, fragmentation, and paged-output profile.</summary>
    public const string PagedPrintV1 = "paged-print-v1";
}

/// <summary>Stable identifiers for implementation providers referenced by HTML capability manifests.</summary>
public static class HtmlCapabilityProviderIds {
    /// <summary>The dependency-free OfficeIMO HTML document contract leaf.</summary>
    public const string OfficeIMOHtmlCore = "officeimo-html-core";
    /// <summary>The OfficeIMO static HTML/CSS renderer.</summary>
    public const string OfficeIMOHtml = "officeimo-html";
    /// <summary>The shared OfficeIMO drawing and text implementation.</summary>
    public const string OfficeIMOCore = "officeimo-core";
    /// <summary>The OfficeIMO HTML-to-PDF adapter.</summary>
    public const string OfficeIMOHtmlPdf = "officeimo-html-pdf";
    /// <summary>The OfficeIMO PDF writer and structure implementation.</summary>
    public const string OfficeIMOPdf = "officeimo-pdf";
    /// <summary>The temporary AngleSharp HTML parser adapter.</summary>
    public const string AngleSharpHtml = "anglesharp-html";
    /// <summary>The temporary AngleSharp CSS parser.</summary>
    public const string AngleSharpCss = "anglesharp-css";
    /// <summary>An optional caller-selected text shaping provider.</summary>
    public const string CallerTextShaper = "caller-text-shaper";
}

/// <summary>Stable identifiers for specifications pinned by HTML capability manifests.</summary>
public static class HtmlCapabilitySpecificationIds {
    /// <summary>WHATWG HTML.</summary>
    public const string Html = "whatwg-html";
    /// <summary>WHATWG DOM.</summary>
    public const string Dom = "whatwg-dom";
    /// <summary>WHATWG URL.</summary>
    public const string Url = "whatwg-url";
    /// <summary>WHATWG Encoding.</summary>
    public const string Encoding = "whatwg-encoding";
    /// <summary>W3C CSS Snapshot 2026.</summary>
    public const string CssSnapshot = "css-snapshot-2026";
    /// <summary>CSS cascade and inheritance.</summary>
    public const string CssCascade = "css-cascade";
    /// <summary>CSS selectors.</summary>
    public const string Selectors = "selectors";
    /// <summary>CSS values and units.</summary>
    public const string CssValues = "css-values";
    /// <summary>CSS color.</summary>
    public const string CssColor = "css-color";
    /// <summary>CSS backgrounds and borders.</summary>
    public const string CssBackgrounds = "css-backgrounds";
    /// <summary>CSS images.</summary>
    public const string CssImages = "css-images";
    /// <summary>CSS transforms and compositing.</summary>
    public const string CssEffects = "css-effects";
    /// <summary>CSS display, box, and inline layout.</summary>
    public const string CssBoxLayout = "css-box-layout";
    /// <summary>CSS text.</summary>
    public const string CssText = "css-text";
    /// <summary>CSS writing modes and ruby.</summary>
    public const string CssWritingModes = "css-writing-modes";
    /// <summary>CSS flexible box layout.</summary>
    public const string CssFlexbox = "css-flexbox";
    /// <summary>CSS grid layout.</summary>
    public const string CssGrid = "css-grid";
    /// <summary>CSS multi-column layout.</summary>
    public const string CssMulticol = "css-multicol";
    /// <summary>CSS positioning.</summary>
    public const string CssPosition = "css-position";
    /// <summary>CSS table layout.</summary>
    public const string CssTables = "css-tables";
    /// <summary>CSS generated content, lists, and counters.</summary>
    public const string CssGeneratedContent = "css-generated-content";
    /// <summary>CSS paged media and fragmentation.</summary>
    public const string CssPagedMedia = "css-paged-media";
    /// <summary>CSS media and container queries.</summary>
    public const string CssConditional = "css-conditional";
    /// <summary>CSS fonts.</summary>
    public const string CssFonts = "css-fonts";
    /// <summary>Unicode 17 algorithms and data.</summary>
    public const string Unicode17 = "unicode-17";
    /// <summary>SVG 2.</summary>
    public const string Svg2 = "svg-2";
    /// <summary>MathML Core.</summary>
    public const string MathMlCore = "mathml-core";
    /// <summary>PDF 2.0.</summary>
    public const string Pdf20 = "pdf-2.0";
}

/// <summary>Stable identifiers for evidence pinned by HTML capability manifests.</summary>
public static class HtmlCapabilityEvidenceIds {
    /// <summary>The executable OfficeIMO HTML contract and regression suite from the package source revision.</summary>
    public const string OfficeIMOHtmlTests = "officeimo-html-tests";
    /// <summary>The selected H1-H2/v1 owned web-document qualification cases.</summary>
    public const string DocumentV1 = "officeimo-html-document-v1";
    /// <summary>The nine continuous-mode cases in the frozen H4/v1 rendering corpus.</summary>
    public const string H4ScreenV1 = "officeimo-html-h4-screen-v1";
    /// <summary>The six paged-mode cases in the frozen H4/v1 rendering corpus.</summary>
    public const string H4PagedV1 = "officeimo-html-h4-paged-v1";
    /// <summary>The pinned html5lib parser test source.</summary>
    public const string Html5Lib = "html5lib-tests";
    /// <summary>The pinned Web Platform Tests source.</summary>
    public const string WebPlatformTests = "web-platform-tests";
}
