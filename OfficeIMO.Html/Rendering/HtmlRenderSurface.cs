using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>
/// Immutable executable view over one retained HTML output surface. Preview, geometry,
/// automation, and encoder consumers share this view and its source mapping.
/// </summary>
public sealed class HtmlRenderSurface {
    private readonly HtmlRenderResult _result;

    internal HtmlRenderSurface(HtmlRenderResult result, HtmlRenderPage page, HtmlRenderSurfaceResult descriptor) {
        _result = result ?? throw new ArgumentNullException(nameof(result));
        Page = page ?? throw new ArgumentNullException(nameof(page));
        Descriptor = descriptor ?? throw new ArgumentNullException(nameof(descriptor));
    }

    /// <summary>Zero-based position in the resolved output set.</summary>
    public int OutputIndex => Descriptor.OutputIndex;
    /// <summary>Retained backend-neutral page and display list.</summary>
    public HtmlRenderPage Page { get; }
    /// <summary>Output geometry, clipping, and source-placement descriptor.</summary>
    public HtmlRenderSurfaceResult Descriptor { get; }
    /// <summary>Complete output-surface rectangle.</summary>
    public HtmlRenderRectangle Bounds => Descriptor.Bounds;
    /// <summary>Requested output scale before bounded encoder reduction.</summary>
    public double RequestedScale => _result.RequestedScale;
    /// <summary>Requested preview and image background.</summary>
    public OfficeColor BackgroundColor => _result.BackgroundColor;

    /// <summary>Creates a detached Drawing preview of this surface.</summary>
    public OfficeDrawing CreateDrawing() => Page.CreateDrawing();

    /// <summary>Creates a detached Drawing preview with cooperative cancellation.</summary>
    public OfficeDrawing CreateDrawing(CancellationToken cancellationToken) => Page.CreateDrawing(cancellationToken);

    /// <summary>Maps an output point to the source page or source-canvas slice.</summary>
    public bool TryMapToSource(HtmlRenderPoint outputPoint, out HtmlRenderSourcePoint? sourcePoint) =>
        Descriptor.TryMapToSource(outputPoint, out sourcePoint);

    /// <summary>Maps finite output coordinates to the source page or source-canvas slice.</summary>
    public bool TryMapToSource(double x, double y, out HtmlRenderSourcePoint? sourcePoint) =>
        TryMapToSource(new HtmlRenderPoint(x, y), out sourcePoint);

    /// <summary>Runs a bounded, topmost-first hit test against the retained scene.</summary>
    public HtmlRenderHitTestReport HitTest(HtmlRenderPoint point, HtmlRenderHitTestOptions? options = null,
        CancellationToken cancellationToken = default) =>
        HtmlRenderHitTester.HitTest(this, point, options, cancellationToken);

    /// <summary>Runs a bounded, topmost-first hit test at finite surface coordinates.</summary>
    public HtmlRenderHitTestReport HitTest(double x, double y, HtmlRenderHitTestOptions? options = null,
        CancellationToken cancellationToken = default) =>
        HitTest(new HtmlRenderPoint(x, y), options, cancellationToken);

    /// <summary>Returns the topmost matching visual, or <see langword="null"/>.</summary>
    public HtmlRenderHitTestResult? HitTestTopmost(HtmlRenderPoint point,
        HtmlRenderHitTestOptions? options = null, CancellationToken cancellationToken = default) =>
        HitTest(point, (options ?? new HtmlRenderHitTestOptions()).WithMaximumResults(1), cancellationToken)
            .Matches.FirstOrDefault();
}
