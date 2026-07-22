using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>
/// Shared layout and safety options used by HTML image and PDF rendering.
/// </summary>
public class HtmlRenderOptions : OfficeImageExportOptions {
    /// <summary>CSS reference pixel density used for physical page conversion.</summary>
    public const double CssPixelsPerInch = 96D;

    /// <summary>Creates the default continuous render configuration.</summary>
    public HtmlRenderOptions() {
        PageSize = OfficePageSizes.A4;
        Margins = HtmlRenderMargins.All(48D);
    }

    /// <summary>Creates an independent copy of shared HTML rendering settings.</summary>
    /// <param name="source">Settings to copy.</param>
    protected HtmlRenderOptions(HtmlRenderOptions source) : this() {
        if (source == null) throw new ArgumentNullException(nameof(source));
        source.CopyTo(this);
    }

    /// <summary>Continuous or paged layout mode.</summary>
    public HtmlRenderMode Mode { get; set; } = HtmlRenderMode.Continuous;

    /// <summary>Viewport width for continuous rendering, in CSS pixels.</summary>
    public double ViewportWidth { get; set; } = 816D;

    /// <summary>Optional minimum continuous-surface height, in CSS pixels.</summary>
    public double? ViewportHeight { get; set; }

    /// <summary>Physical page size used by paged rendering.</summary>
    public OfficePageSize PageSize { get; set; }

    /// <summary>When true, generic print <c>@page</c> size and margin declarations override the paged defaults.</summary>
    public bool HonorCssPageRules { get; set; } = true;

    /// <summary>Page or continuous-surface margins measured in CSS pixels.</summary>
    public HtmlRenderMargins Margins { get; set; }

    /// <summary>Default font family used when CSS does not select one.</summary>
    public string DefaultFontFamily { get; set; } = "Arial";

    /// <summary>Default CSS font size in pixels.</summary>
    public double DefaultFontSize { get; set; } = 16D;

    /// <summary>Default line-height multiplier.</summary>
    public double DefaultLineHeight { get; set; } = 1.2D;

    /// <summary>Optional base URI used to resolve links and resource references.</summary>
    public Uri? BaseUri { get; set; }

    /// <summary>URL policy applied before links or resources enter the rendered result.</summary>
    public HtmlUrlPolicy UrlPolicy { get; set; } = HtmlUrlPolicy.CreateOfficeIMOProfile();

    /// <summary>
    /// Optional separate policy for images, stylesheets, fonts, and other render resources.
    /// When omitted, <see cref="UrlPolicy"/> is used. This lets package or application resolvers
    /// authorize a resource origin without authorizing the same scheme for emitted hyperlinks.
    /// </summary>
    public HtmlUrlPolicy? ResourceUrlPolicy { get; set; }

    /// <summary>Optional application-supplied asynchronous resolver for policy-approved external resources.</summary>
    public HtmlRenderResourceResolver? ResourceResolver { get; set; }

    // Package-backed adapters use this path for resources that are already retained locally.
    // External application resolvers remain asynchronous and are intentionally not blocked on
    // by the synchronous rendering API.
    internal HtmlRenderSynchronousResourceResolver? SynchronousResourceResolver { get; set; }

    /// <summary>Maximum time allowed for one resolver invocation.</summary>
    public TimeSpan ResourceTimeout { get; set; } = TimeSpan.FromSeconds(30D);

    /// <summary>Maximum asynchronous resolver invocations allowed to run concurrently.</summary>
    public int MaxConcurrentResourceLoads { get; set; } = 8;

    /// <summary>Maximum bytes accepted from one resolved resource.</summary>
    public long MaxResourceBytes { get; set; } = 10L * 1024L * 1024L;

    /// <summary>Maximum total bytes accepted from external resources in one render operation.</summary>
    public long MaxTotalResourceBytes { get; set; } = 50L * 1024L * 1024L;

    /// <summary>Maximum number of external resources accepted in one render operation.</summary>
    public int MaxResourceCount { get; set; } = 256;

    /// <summary>Maximum resolver invocations attempted in one render operation, including misses and failures.</summary>
    public int MaxResourceRequests { get; set; } = 512;

    /// <summary>Maximum recursive <c>@import</c> depth accepted for external stylesheets.</summary>
    public int MaxStylesheetImportDepth { get; set; } = 16;

    /// <summary>Maximum width of any generated image surface in output pixels.</summary>
    public int MaxSurfaceWidth { get; set; } = 32768;

    /// <summary>Maximum height of any generated image surface in output pixels.</summary>
    public int MaxSurfaceHeight { get; set; } = 32768;

    /// <summary>Maximum page count accepted from one paged render operation.</summary>
    public int MaxPageCount { get; set; } = 1000;

    /// <summary>Maximum element nesting depth processed by the layout engine.</summary>
    public int MaxLayoutDepth { get; set; } = 256;

    /// <summary>
    /// Maximum UTF-16 characters accepted in the source HTML string. The default leaves enough room for
    /// the default total resource budget when resources are embedded as base64 data URIs.
    /// </summary>
    public int MaxInputCharacters { get; set; } = 72 * 1024 * 1024;

    /// <summary>Maximum DOM nodes accepted after parsing and before style or layout work begins.</summary>
    public int MaxHtmlNodes { get; set; } = 100000;

    /// <summary>Maximum total repeated background-image tiles accepted in one render operation.</summary>
    public int MaxBackgroundImageTiles { get; set; } = 16384;

    /// <summary>Maximum background-image layers accepted on one element.</summary>
    public int MaxBackgroundImageLayers { get; set; } = 32;

    /// <summary>Maximum CSS box-shadow layers accepted on one element.</summary>
    public int MaxBoxShadowLayers { get; set; } = 32;

    /// <summary>Maximum color stops accepted in one CSS gradient.</summary>
    public int MaxGradientStops { get; set; } = 64;

    /// <summary>Maximum explicit or implicit tracks accepted on either grid axis.</summary>
    public int MaxGridTracks { get; set; } = 256;

    /// <summary>Maximum generated columns accepted in one multi-column formatting context.</summary>
    public int MaxColumnCount { get; set; } = 64;

    /// <summary>Gets the CSS media context selected by the current render mode.</summary>
    public HtmlCssMediaContext MediaContext => Mode == HtmlRenderMode.Paged ? HtmlCssMediaContext.Print : HtmlCssMediaContext.Screen;

    /// <summary>Gets the paged surface width in CSS pixels.</summary>
    public double PageWidth => PageSize.WidthInches * CssPixelsPerInch;

    /// <summary>Gets the paged surface height in CSS pixels.</summary>
    public double PageHeight => PageSize.HeightInches * CssPixelsPerInch;

    /// <summary>Creates an independent options snapshot.</summary>
    public virtual HtmlRenderOptions Clone() => CopyTo(new HtmlRenderOptions());

    /// <summary>Copies shared layout and resource settings into a target-specific options instance.</summary>
    protected internal T CopyTo<T>(T target) where T : HtmlRenderOptions {
        CopyImageExportOptionsTo(target);
        target.Mode = Mode;
        target.ViewportWidth = ViewportWidth;
        target.ViewportHeight = ViewportHeight;
        target.PageSize = PageSize;
        target.HonorCssPageRules = HonorCssPageRules;
        target.Margins = Margins;
        target.DefaultFontFamily = DefaultFontFamily;
        target.DefaultFontSize = DefaultFontSize;
        target.DefaultLineHeight = DefaultLineHeight;
        target.BaseUri = BaseUri;
        target.UrlPolicy = (UrlPolicy ?? HtmlUrlPolicy.CreateOfficeIMOProfile()).Clone();
        target.ResourceUrlPolicy = ResourceUrlPolicy?.Clone();
        target.ResourceResolver = ResourceResolver;
        target.SynchronousResourceResolver = SynchronousResourceResolver;
        target.ResourceTimeout = ResourceTimeout;
        target.MaxConcurrentResourceLoads = MaxConcurrentResourceLoads;
        target.MaxResourceBytes = MaxResourceBytes;
        target.MaxTotalResourceBytes = MaxTotalResourceBytes;
        target.MaxResourceCount = MaxResourceCount;
        target.MaxResourceRequests = MaxResourceRequests;
        target.MaxStylesheetImportDepth = MaxStylesheetImportDepth;
        target.MaxSurfaceWidth = MaxSurfaceWidth;
        target.MaxSurfaceHeight = MaxSurfaceHeight;
        target.MaxPageCount = MaxPageCount;
        target.MaxLayoutDepth = MaxLayoutDepth;
        target.MaxInputCharacters = MaxInputCharacters;
        target.MaxHtmlNodes = MaxHtmlNodes;
        target.MaxBackgroundImageTiles = MaxBackgroundImageTiles;
        target.MaxBackgroundImageLayers = MaxBackgroundImageLayers;
        target.MaxBoxShadowLayers = MaxBoxShadowLayers;
        target.MaxGradientStops = MaxGradientStops;
        target.MaxGridTracks = MaxGridTracks;
        target.MaxColumnCount = MaxColumnCount;
        target.ResponsiveImageCandidateLimit = ResponsiveImageCandidateLimit;
        return target;
    }

    internal HtmlUrlPolicy GetResourceUrlPolicy() =>
        ResourceUrlPolicy ?? UrlPolicy ?? HtmlUrlPolicy.CreateOfficeIMOProfile();

    internal int? ResponsiveImageCandidateLimit { get; set; }

    internal void Validate() {
        ValidateImageExportOptions();
        ValidatePositive(ViewportWidth, nameof(ViewportWidth));
        if (ViewportHeight.HasValue) {
            ValidatePositive(ViewportHeight.Value, nameof(ViewportHeight));
        }

        ValidatePositive(DefaultFontSize, nameof(DefaultFontSize));
        ValidatePositive(DefaultLineHeight, nameof(DefaultLineHeight));
        if (string.IsNullOrWhiteSpace(DefaultFontFamily)) {
            throw new ArgumentException("A default font family is required.", nameof(DefaultFontFamily));
        }

        if (MaxSurfaceWidth <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxSurfaceWidth), "Maximum surface width must be positive.");
        }

        if (MaxSurfaceHeight <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxSurfaceHeight), "Maximum surface height must be positive.");
        }

        if (MaxPageCount <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxPageCount), "Maximum page count must be positive.");
        }

        if (MaxLayoutDepth <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxLayoutDepth), "Maximum layout depth must be positive.");
        }

        if (MaxInputCharacters <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxInputCharacters), "Maximum source HTML character count must be positive.");
        }

        if (MaxHtmlNodes <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxHtmlNodes), "Maximum HTML DOM node count must be positive.");
        }

        if (MaxBackgroundImageTiles <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxBackgroundImageTiles), "Maximum background-image tile count must be positive.");
        }

        if (MaxBackgroundImageLayers <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxBackgroundImageLayers), "Maximum background-image layer count must be positive.");
        }

        if (MaxBoxShadowLayers <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxBoxShadowLayers), "Maximum box-shadow layer count must be positive.");
        }

        if (MaxGradientStops < 2) {
            throw new ArgumentOutOfRangeException(nameof(MaxGradientStops), "Maximum gradient stop count must be at least two.");
        }
        if (MaxGridTracks <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxGridTracks), "Maximum grid track count must be positive.");
        }
        if (MaxColumnCount <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxColumnCount), "Maximum multi-column count must be positive.");
        }

        if (ResourceTimeout <= TimeSpan.Zero || ResourceTimeout == System.Threading.Timeout.InfiniteTimeSpan) {
            throw new ArgumentOutOfRangeException(nameof(ResourceTimeout), "Resource timeout must be a finite positive duration.");
        }

        if (MaxConcurrentResourceLoads <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxConcurrentResourceLoads), "Maximum concurrent resource loads must be positive.");
        }

        if (MaxResourceBytes <= 0L) {
            throw new ArgumentOutOfRangeException(nameof(MaxResourceBytes), "Maximum resource bytes must be positive.");
        }

        if (MaxTotalResourceBytes <= 0L || MaxTotalResourceBytes < MaxResourceBytes) {
            throw new ArgumentOutOfRangeException(nameof(MaxTotalResourceBytes), "Maximum total resource bytes must be positive and at least the per-resource limit.");
        }

        if (MaxResourceCount <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxResourceCount), "Maximum resource count must be positive.");
        }

        if (MaxResourceRequests <= 0 || MaxResourceRequests < MaxResourceCount) {
            throw new ArgumentOutOfRangeException(nameof(MaxResourceRequests), "Maximum resource requests must be positive and at least the accepted resource count limit.");
        }

        if (MaxStylesheetImportDepth <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxStylesheetImportDepth), "Maximum stylesheet import depth must be positive.");
        }

        double surfaceWidth = Mode == HtmlRenderMode.Paged ? PageWidth : ViewportWidth;
        double surfaceHeight = Mode == HtmlRenderMode.Paged ? PageHeight : ViewportHeight ?? 1D;
        if (Margins.Left + Margins.Right >= surfaceWidth || Margins.Top + Margins.Bottom >= surfaceHeight && Mode == HtmlRenderMode.Paged) {
            throw new ArgumentException("Render margins must leave a positive content area.", nameof(Margins));
        }
    }

    private static void ValidatePositive(double value, string parameterName) {
        if (value <= 0D || double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(parameterName, "HTML render values must be finite positive numbers.");
        }
    }
}
