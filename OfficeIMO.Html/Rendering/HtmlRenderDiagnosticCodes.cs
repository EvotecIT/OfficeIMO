using System.Collections.ObjectModel;

namespace OfficeIMO.Html;

/// <summary>
/// Stable diagnostic codes emitted by the first-party HTML renderer.
/// </summary>
public static class HtmlRenderDiagnosticCodes {
    /// <summary>Layout exceeded the configured nesting-depth limit.</summary>
    public const string DepthLimitExceeded = "HtmlRenderDepthLimitExceeded";
    /// <summary>A table contained no renderable rows or cells.</summary>
    public const string EmptyTable = "HtmlRenderEmptyTable";
    /// <summary>An external image requires asynchronous resource resolution.</summary>
    public const string ExternalImagePending = "HtmlRenderExternalImagePending";
    /// <summary>Flex layout used the documented normal-flow fallback.</summary>
    public const string FlexLayoutPending = "HtmlRenderFlexLayoutPending";
    /// <summary>Content without a safe page break was force-fragmented.</summary>
    public const string ForcedFragment = "HtmlRenderForcedFragment";
    /// <summary>Grid layout used the documented normal-flow fallback.</summary>
    public const string GridLayoutPending = "HtmlRenderGridLayoutPending";
    /// <summary>An inline image used its alternative-text fallback.</summary>
    public const string InlineImageFallback = "HtmlRenderInlineImageFallback";
    /// <summary>A named or pseudo-page selector could not yet be applied per page.</summary>
    public const string PageSelectorPending = "HtmlRenderPageSelectorPending";
    /// <summary>An <c>@page</c> size declaration was unsupported.</summary>
    public const string PageSizeUnsupported = "HtmlRenderPageSizeUnsupported";
    /// <summary>The dependency-free PNG backend could not decode a retained raster format.</summary>
    public const string RasterDecoderUnavailable = "HtmlRenderRasterDecoderUnavailable";
    /// <summary>A resource exceeded the configured per-resource byte limit.</summary>
    public const string ResourceByteLimitExceeded = "HtmlRenderResourceByteLimitExceeded";
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
    /// <summary>A table row span could not yet be fragmented across pages.</summary>
    public const string TableRowSpanPending = "HtmlRenderTableRowSpanPending";
    /// <summary>A repeated table header was suppressed because it left no safe body-row break.</summary>
    public const string TableHeaderRepeatSuppressed = "HtmlRenderTableHeaderRepeatSuppressed";
    /// <summary>Resolved resources exceeded the operation-wide byte budget.</summary>
    public const string TotalResourceByteLimitExceeded = "HtmlRenderTotalResourceByteLimitExceeded";
    /// <summary>A visual could not cross a forced page boundary safely.</summary>
    public const string VisualFragmentUnsupported = "HtmlRenderVisualFragmentUnsupported";

    /// <summary>All stable renderer diagnostic codes.</summary>
    public static IReadOnlyList<string> All { get; } = new ReadOnlyCollection<string>(new[] {
        DepthLimitExceeded,
        EmptyTable,
        ExternalImagePending,
        FlexLayoutPending,
        ForcedFragment,
        GridLayoutPending,
        InlineImageFallback,
        PageSelectorPending,
        PageSizeUnsupported,
        RasterDecoderUnavailable,
        ResourceByteLimitExceeded,
        ResourceContentTypeRejected,
        ResourceLoadFailed,
        ResourceTimeout,
        ResourceUnavailable,
        ResourceUriInvalid,
        TableHeaderRepeatSuppressed,
        TableRowSpanPending,
        TotalResourceByteLimitExceeded,
        VisualFragmentUnsupported
    });
}
