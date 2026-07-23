using System.Collections.ObjectModel;

namespace OfficeIMO.Html;

/// <summary>
/// Stable diagnostic codes emitted by the first-party HTML renderer.
/// </summary>
public static class HtmlRenderDiagnosticCodes {
    /// <summary>CSS background-image layers beyond the configured per-element limit were omitted.</summary>
    public const string BackgroundImageLayerLimit = "HtmlRenderBackgroundImageLayerLimit";
    /// <summary>A CSS background-repeat value used a single-image fallback.</summary>
    public const string BackgroundImageRepeatUnsupported = "HtmlRenderBackgroundImageRepeatUnsupported";
    /// <summary>A CSS background image value used a deterministic supported fallback or was omitted.</summary>
    public const string BackgroundImageValueUnsupported = "HtmlRenderBackgroundImageValueUnsupported";
    /// <summary>Repeated CSS background images exceeded the configured operation-wide tile limit.</summary>
    public const string BackgroundImageTileLimitExceeded = "HtmlRenderBackgroundImageTileLimitExceeded";
    /// <summary>CSS gradients exceeded the configured color-stop limit.</summary>
    public const string GradientStopLimitExceeded = "HtmlRenderGradientStopLimitExceeded";
    /// <summary>Layout exceeded the configured nesting-depth limit.</summary>
    public const string DepthLimitExceeded = "HtmlRenderDepthLimitExceeded";
    /// <summary>Layout exceeded the configured operation limit.</summary>
    public const string LayoutOperationLimitExceeded = "HtmlRenderLayoutOperationLimitExceeded";
    /// <summary>The source HTML string exceeded the configured character limit.</summary>
    public const string InputCharacterLimitExceeded = "HtmlRenderInputCharacterLimitExceeded";
    /// <summary>The parsed HTML document exceeded the configured DOM node limit.</summary>
    public const string NodeLimitExceeded = "HtmlNodeLimitExceeded";
    /// <summary>A table contained no renderable rows or cells.</summary>
    public const string EmptyTable = "HtmlRenderEmptyTable";
    /// <summary>A table exceeded the configured row or column limit.</summary>
    public const string TableLimitExceeded = "HtmlRenderTableLimitExceeded";
    /// <summary>Collapsed table border resolution exceeded the configured segment limit.</summary>
    public const string CollapsedTableBorderLimitExceeded = "HtmlRenderCollapsedTableBorderLimitExceeded";
    /// <summary>An external image requires asynchronous resource resolution.</summary>
    public const string ExternalImagePending = "HtmlRenderExternalImagePending";
    /// <summary>An external stylesheet requires asynchronous resource resolution.</summary>
    public const string ExternalStylesheetPending = "HtmlRenderExternalStylesheetPending";
    /// <summary>A font data URI could not be decoded.</summary>
    public const string FontDataUriInvalid = "HtmlRenderFontDataUriInvalid";
    /// <summary>An @font-face rule had no usable family descriptor.</summary>
    public const string FontFaceInvalid = "HtmlRenderFontFaceInvalid";
    /// <summary>No source from an @font-face rule was available to the renderer.</summary>
    public const string FontFaceUnavailable = "HtmlRenderFontFaceUnavailable";
    /// <summary>A font source was not a supported TrueType glyf-outline font.</summary>
    public const string FontFormatUnsupported = "HtmlRenderFontFormatUnsupported";
    /// <summary>Right-to-left inline content requires bidi positioning not yet active in the shared renderer.</summary>
    public const string BidiLayoutUnsupported = "HtmlRenderBidiLayoutUnsupported";
    /// <summary>A joining script is outside the bounded core-Arabic contextual shaper.</summary>
    public const string ComplexTextShapingUnsupported = "HtmlRenderComplexTextShapingUnsupported";
    /// <summary>Flex layout used the documented normal-flow fallback.</summary>
    public const string FlexLayoutPending = "HtmlRenderFlexLayoutPending";
    /// <summary>A flex property value used a documented deterministic fallback.</summary>
    public const string FlexValueUnsupported = "HtmlRenderFlexValueUnsupported";
    /// <summary>A float or clear property value used a documented deterministic fallback.</summary>
    public const string FloatValueUnsupported = "HtmlRenderFloatValueUnsupported";
    /// <summary>Multi-column generation exceeded the configured safety limit.</summary>
    public const string MultiColumnLimitExceeded = "HtmlRenderMultiColumnLimitExceeded";
    /// <summary>A multi-column property value used a documented deterministic fallback.</summary>
    public const string MultiColumnValueUnsupported = "HtmlRenderMultiColumnValueUnsupported";
    /// <summary>Content without a safe page break was force-fragmented.</summary>
    public const string ForcedFragment = "HtmlRenderForcedFragment";
    /// <summary>Grid layout used the documented normal-flow fallback.</summary>
    public const string GridLayoutPending = "HtmlRenderGridLayoutPending";
    /// <summary>A grid property value used a documented deterministic fallback.</summary>
    public const string GridValueUnsupported = "HtmlRenderGridValueUnsupported";
    /// <summary>Grid track expansion exceeded the configured safety limit.</summary>
    public const string GridTrackLimitExceeded = "HtmlRenderGridTrackLimitExceeded";
    /// <summary>A generated-content expression was omitted because it could not be represented.</summary>
    public const string GeneratedContentUnsupported = "HtmlRenderGeneratedContentUnsupported";
    /// <summary>A CSS counter declaration was ignored because it could not be represented.</summary>
    public const string GeneratedCounterUnsupported = "HtmlRenderGeneratedCounterUnsupported";
    /// <summary>A replaced-element sizing or object-placement value used a deterministic fallback.</summary>
    public const string ReplacedElementValueUnsupported = "HtmlRenderReplacedElementValueUnsupported";
    /// <summary>A positioned inset could not be resolved by the current length model.</summary>
    public const string PositionInsetUnsupported = "HtmlRenderPositionInsetUnsupported";
    /// <summary>A positioned layout mode used the documented normal-flow fallback.</summary>
    public const string PositioningModeUnsupported = "HtmlRenderPositioningModeUnsupported";
    /// <summary>An automatic positioned inset could not use its hypothetical normal-flow anchor.</summary>
    public const string PositionStaticAnchorFallback = "HtmlRenderPositionStaticAnchorFallback";
    /// <summary>A scrollable overflow box was captured at its initial static scroll position.</summary>
    public const string OverflowScrollSnapshot = "HtmlRenderOverflowScrollSnapshot";
    /// <summary>An overflow property value used a documented visible fallback.</summary>
    public const string OverflowValueUnsupported = "HtmlRenderOverflowValueUnsupported";
    /// <summary>An overflow-clip-margin value used its initial padding-box zero fallback.</summary>
    public const string OverflowClipMarginValueUnsupported = "HtmlRenderOverflowClipMarginValueUnsupported";
    /// <summary>A CSS transform or transform-origin value used the identity fallback.</summary>
    public const string TransformValueUnsupported = "HtmlRenderTransformValueUnsupported";
    /// <summary>A CSS opacity value used the opaque fallback.</summary>
    public const string OpacityValueUnsupported = "HtmlRenderOpacityValueUnsupported";
    /// <summary>A non-atomic inline paint effect could not form an isolated group.</summary>
    public const string InlinePaintEffectUnsupported = "HtmlRenderInlinePaintEffectUnsupported";
    /// <summary>A CSS border radius contained invalid or unsupported syntax and used square-corner fallback.</summary>
    public const string BorderRadiusValueUnsupported = "HtmlRenderBorderRadiusValueUnsupported";
    /// <summary>A CSS box shadow contained invalid or unsupported syntax.</summary>
    public const string BoxShadowValueUnsupported = "HtmlRenderBoxShadowValueUnsupported";
    /// <summary>CSS box-shadow layers beyond the configured per-element limit were omitted.</summary>
    public const string BoxShadowLayerLimit = "HtmlRenderBoxShadowLayerLimit";
    /// <summary>A CSS border paint declaration used invalid or unsupported syntax.</summary>
    public const string BorderPaintValueUnsupported = "HtmlRenderBorderPaintValueUnsupported";
    /// <summary>A CSS outline paint declaration could not be represented by the current outline model.</summary>
    public const string OutlinePaintValueUnsupported = "HtmlRenderOutlinePaintValueUnsupported";
    /// <summary>A sticky-positioned element was captured at its stable static document position.</summary>
    public const string PositionStickyStatic = "HtmlRenderPositionStickyStatic";
    /// <summary>A positioned element declared stacking behavior that is not active yet.</summary>
    public const string PositionZIndexPending = "HtmlRenderPositionZIndexPending";
    /// <summary>A complex page selector could not be applied per page.</summary>
    public const string PageSelectorPending = "HtmlRenderPageSelectorPending";
    /// <summary>A pseudo-page geometry declaration requires page-by-page content reflow.</summary>
    public const string PagePseudoGeometryPending = "HtmlRenderPagePseudoGeometryPending";
    /// <summary>A page-margin generated-content expression was unsupported.</summary>
    public const string PageMarginContentUnsupported = "HtmlRenderPageMarginContentUnsupported";
    /// <summary>A page-margin position was unsupported by the current visual model.</summary>
    public const string PageMarginPositionUnsupported = "HtmlRenderPageMarginPositionUnsupported";
    /// <summary>An <c>@page</c> size declaration was unsupported.</summary>
    public const string PageSizeUnsupported = "HtmlRenderPageSizeUnsupported";
    /// <summary>SVG content could not be represented by the bounded shared vector scene.</summary>
    public const string SvgContentUnsupported = "HtmlRenderSvgContentUnsupported";
    /// <summary>A resource exceeded the configured per-resource byte limit.</summary>
    public const string ResourceByteLimitExceeded = "HtmlRenderResourceByteLimitExceeded";
    /// <summary>Resolved resources exceeded the operation-wide count limit.</summary>
    public const string ResourceCountLimitExceeded = "HtmlRenderResourceCountLimitExceeded";
    /// <summary>Resolver invocations exceeded the operation-wide request limit.</summary>
    public const string ResourceRequestLimitExceeded = "HtmlRenderResourceRequestLimitExceeded";
    /// <summary>A resolver returned an incompatible media type.</summary>
    public const string ResourceContentTypeRejected = "HtmlRenderResourceContentTypeRejected";
    /// <summary>The configured resource resolver failed.</summary>
    public const string ResourceLoadFailed = "HtmlRenderResourceLoadFailed";
    /// <summary>Resource resolution exceeded its configured timeout.</summary>
    public const string ResourceTimeout = "HtmlRenderResourceTimeout";
    /// <summary>The configured resource resolver returned no resource.</summary>
    public const string ResourceUnavailable = "HtmlRenderResourceUnavailable";
    /// <summary>A resource reference could not be represented as an absolute URI.</summary>
    public const string ResourceUriInvalid = "HtmlRenderResourceUriInvalid";
    /// <summary>A resolved stylesheet could not be decoded as supported CSS text.</summary>
    public const string StylesheetEncodingUnsupported = "HtmlRenderStylesheetEncodingUnsupported";
    /// <summary>A recursive stylesheet import cycle was suppressed.</summary>
    public const string StylesheetImportCycle = "HtmlRenderStylesheetImportCycle";
    /// <summary>Stylesheet imports exceeded the configured recursion depth.</summary>
    public const string StylesheetImportDepthExceeded = "HtmlRenderStylesheetImportDepthExceeded";
    /// <summary>A resolved stylesheet referenced URL resources that are not active in the current paint model.</summary>
    public const string StylesheetUrlResourcesPending = "HtmlRenderStylesheetUrlResourcesPending";
    /// <summary>A repeated table header was suppressed because it left no safe body-row break.</summary>
    public const string TableHeaderRepeatSuppressed = "HtmlRenderTableHeaderRepeatSuppressed";
    /// <summary>A repeated table footer was suppressed because it left no safe body-row break.</summary>
    public const string TableFooterRepeatSuppressed = "HtmlRenderTableFooterRepeatSuppressed";
    /// <summary>A table formatting property used its documented fallback.</summary>
    public const string TableValueUnsupported = "HtmlRenderTableValueUnsupported";
    /// <summary>Resolved resources exceeded the operation-wide byte budget.</summary>
    public const string TotalResourceByteLimitExceeded = "HtmlRenderTotalResourceByteLimitExceeded";
    /// <summary>A visual could not cross a forced page boundary safely.</summary>
    public const string VisualFragmentUnsupported = "HtmlRenderVisualFragmentUnsupported";

    /// <summary>All stable renderer diagnostic codes.</summary>
    public static IReadOnlyList<string> All { get; } = new ReadOnlyCollection<string>(new[] {
        BackgroundImageLayerLimit,
        BackgroundImageRepeatUnsupported,
        BackgroundImageValueUnsupported,
        BackgroundImageTileLimitExceeded,
        GradientStopLimitExceeded,
        DepthLimitExceeded,
        LayoutOperationLimitExceeded,
        InputCharacterLimitExceeded,
        NodeLimitExceeded,
        EmptyTable,
        TableLimitExceeded,
        CollapsedTableBorderLimitExceeded,
        ExternalImagePending,
        ExternalStylesheetPending,
        FontDataUriInvalid,
        FontFaceInvalid,
        FontFaceUnavailable,
        FontFormatUnsupported,
        BidiLayoutUnsupported,
        ComplexTextShapingUnsupported,
        FlexLayoutPending,
        FlexValueUnsupported,
        FloatValueUnsupported,
        MultiColumnLimitExceeded,
        MultiColumnValueUnsupported,
        ForcedFragment,
        GeneratedContentUnsupported,
        GeneratedCounterUnsupported,
        GridLayoutPending,
        GridTrackLimitExceeded,
        GridValueUnsupported,
        ReplacedElementValueUnsupported,
        OverflowClipMarginValueUnsupported,
        OverflowScrollSnapshot,
        OverflowValueUnsupported,
        TransformValueUnsupported,
        OpacityValueUnsupported,
        InlinePaintEffectUnsupported,
        BorderRadiusValueUnsupported,
        BoxShadowLayerLimit,
        BoxShadowValueUnsupported,
        BorderPaintValueUnsupported,
        OutlinePaintValueUnsupported,
        PositionInsetUnsupported,
        PositioningModeUnsupported,
        PositionStaticAnchorFallback,
        PositionStickyStatic,
        PositionZIndexPending,
        PageMarginContentUnsupported,
        PageMarginPositionUnsupported,
        PagePseudoGeometryPending,
        PageSelectorPending,
        PageSizeUnsupported,
        SvgContentUnsupported,
        ResourceByteLimitExceeded,
        ResourceCountLimitExceeded,
        ResourceRequestLimitExceeded,
        ResourceContentTypeRejected,
        ResourceLoadFailed,
        ResourceTimeout,
        ResourceUnavailable,
        ResourceUriInvalid,
        StylesheetEncodingUnsupported,
        StylesheetImportCycle,
        StylesheetImportDepthExceeded,
        StylesheetUrlResourcesPending,
        TableFooterRepeatSuppressed,
        TableHeaderRepeatSuppressed,
        TableValueUnsupported,
        TotalResourceByteLimitExceeded,
        VisualFragmentUnsupported
    });
}
