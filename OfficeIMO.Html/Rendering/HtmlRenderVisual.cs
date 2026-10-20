namespace OfficeIMO.Html;

/// <summary>
/// Immutable positioned visual emitted by the shared HTML layout engine.
/// </summary>
public abstract partial class HtmlRenderVisual {
    internal HtmlRenderVisual(HtmlRenderVisualKind kind, double x, double y, double width, double height, int paintOrder, string? linkUri, string? source, double? layoutY = null, double? layoutHeight = null, bool allowZeroWidth = false) {
        ValidateFinite(x, nameof(x));
        ValidateFinite(y, nameof(y));
        ValidateFinite(layoutY ?? y, nameof(layoutY));
        if (!allowZeroWidth || width != 0D) ValidatePositive(width, nameof(width));
        ValidatePositive(height, nameof(height));
        ValidatePositive(layoutHeight ?? height, nameof(layoutHeight));
        Kind = kind;
        X = x;
        Y = y;
        Width = width;
        Height = height;
        PaintOrder = paintOrder;
        LinkUri = linkUri;
        Source = source;
        LayoutY = layoutY ?? y;
        LayoutHeight = layoutHeight ?? height;
    }

    /// <summary>Visual operation kind.</summary>
    public HtmlRenderVisualKind Kind { get; }

    /// <summary>Left coordinate in CSS pixels.</summary>
    public double X { get; }

    /// <summary>Top coordinate in CSS pixels.</summary>
    public double Y { get; }

    /// <summary>Visual width in CSS pixels. A semantic container may have zero width.</summary>
    public double Width { get; }

    /// <summary>Visual height in CSS pixels.</summary>
    public double Height { get; }

    /// <summary>Stable paint order within the rendered page.</summary>
    public int PaintOrder { get; }

    /// <summary>Optional safe hyperlink associated with the visual.</summary>
    public string? LinkUri { get; }

    /// <summary>Optional source element description used by diagnostics and inspection.</summary>
    public string? Source { get; }

    /// <summary>
    /// Gets the normal-flow top coordinate used internally for fragmentation.
    /// Paint-only CSS positioning deliberately leaves this coordinate unchanged.
    /// </summary>
    internal double LayoutY { get; }

    internal HtmlRenderPaintPhase PaintPhase { get; set; }

    // Source adapters retain semantic placement separately from paint and fragment arrival.
    internal int? SemanticSourceOrder { get; set; }

    /// <summary>
    /// A boundary whose children use a different DOM source coordinate space.
    /// Extraction orders the boundary at <see cref="SemanticSourceOrder"/> in its
    /// parent document, then orders the child document independently.
    /// </summary>
    internal HtmlRenderSourceOrderScope? SourceOrderScope { get; set; }

    // A repeated table fragment reopens only its enclosing formatting context.
    // Remember this closure's immediate phase, including an already-atomic
    // nested box, instead of guessing from the visual's shape or children.
    internal HtmlRenderPaintPhase? AtomicSourcePaintPhase { get; private set; }

    internal void CloseFlowPaint() {
        AtomicSourcePaintPhase = PaintPhase;
        PaintPhase = HtmlRenderPaintPhase.Atomic;
    }

    /// <summary>Normal-flow height used for fragmentation, independent of paint overhang.</summary>
    internal double LayoutHeight { get; }

    internal HtmlRenderStackingContext? StackingContext { get; private set; }

    // Split paint envelopes still describe the same layout box. Identity, rather
    // than equal coordinates, keeps distinct overlapping source elements apart.
    internal object? PaintProjectionIdentity { get; private set; }

    // Called only for a freshly reconstructed paint envelope.
    internal HtmlRenderVisual IdentifyPaintProjection(object identity) {
        PaintProjectionIdentity = identity;
        return this;
    }

    internal HtmlRenderVisual WithStackingContext(HtmlRenderStackingContext context) {
        HtmlRenderVisual result = Translate(0D, 0D, PaintOrder);
        result.StackingContext = context;
        return result;
    }

    internal T CopyStackingContextTo<T>(T result) where T : HtmlRenderVisual {
        result.PaintPhase = PaintPhase;
        // A rebuilt embedded viewport retains its parent-document source entry
        // even when a promoted child supplies the surrounding paint metadata.
        if (result.SourceOrderScope == null) {
            result.SemanticSourceOrder = SemanticSourceOrder;
            result.SourceOrderScope = SourceOrderScope;
        }
        result.AtomicSourcePaintPhase = AtomicSourcePaintPhase;
        result.StackingContext = StackingContext;
        result.PaintProjectionIdentity = PaintProjectionIdentity;
        result.RelativePaintOffsetY = RelativePaintOffsetY;
        result.IsOutOfFlowPaint = IsOutOfFlowPaint;
        result.FollowsColumnStaticPosition = FollowsColumnStaticPosition;
        return result;
    }

    internal HtmlRenderVisual Translate(double offsetX, double offsetY, int paintOrder) =>
        CopyStackingContextTo(TranslateCore(offsetX, offsetY, paintOrder));

    internal HtmlRenderVisual TranslatePaint(double offsetX, double offsetY, int paintOrder) =>
        CopyStackingContextTo(TranslatePaintCore(offsetX, offsetY, paintOrder));

    internal abstract HtmlRenderVisual TranslateCore(double offsetX, double offsetY, int paintOrder);

    internal abstract HtmlRenderVisual TranslatePaintCore(double offsetX, double offsetY, int paintOrder);

    private static void ValidateFinite(double value, string parameterName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(parameterName, "Render coordinates must be finite.");
        }
    }

    private static void ValidatePositive(double value, string parameterName) {
        if (value <= 0D || double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(parameterName, "Render dimensions must be finite positive numbers.");
        }
    }
}

internal enum HtmlRenderPaintPhase { Content, BlockBackground, Float, Atomic, Outline }

/// <summary>Shared identity of projected fragments from one independently ordered embedded document.</summary>
internal sealed class HtmlRenderSourceOrderScope { }
