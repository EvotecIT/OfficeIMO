namespace OfficeIMO.Drawing;

/// <summary>Converts a surface pixel after all source paints have been composited.</summary>
internal interface IOfficePostCompositionColorTransform {
    bool TryTransform(OfficeColor source, out OfficeColor transformed);
}

public sealed partial class OfficeDrawing {
    internal IOfficePostCompositionColorTransform? PostCompositionColorTransform { get; private set; }
    internal OfficeColor? PostCompositionBackdrop { get; private set; }

    /// <summary>Retains scene semantics while deferring destination color conversion to rasterization.</summary>
    internal void SetPostCompositionColorTransform(IOfficePostCompositionColorTransform transform, OfficeColor backdrop) {
        ArgumentNullException.ThrowIfNull(transform);
        PostCompositionColorTransform = transform;
        PostCompositionBackdrop = backdrop;
    }

    /// <summary>Creates a render-only view of the authored paints for region composition.</summary>
    internal OfficeDrawing CreateCompositionSource() {
        if (PostCompositionColorTransform == null) return this;
        var source = new OfficeDrawing(Width, Height) {
            _fonts = _fonts,
            TextShapingProvider = TextShapingProvider,
            TextShapingLanguage = TextShapingLanguage,
            SvgTextAsPaths = SvgTextAsPaths
        };
        source._elements.AddRange(_elements);
        return source;
    }
}
