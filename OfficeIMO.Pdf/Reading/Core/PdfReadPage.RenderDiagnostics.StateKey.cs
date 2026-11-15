namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    /// <summary>The paint state that decides whether a form invocation needs its own capability scan.</summary>
    /// <remarks>Equality is component-wise, as for a value tuple of the same components, without boxing each one.</remarks>
    private readonly struct XObjectInvocationStateKey : IEquatable<XObjectInvocationStateKey> {
        private readonly string? _fillName;
        private readonly bool _hasFillTint;
        private readonly PdfPageColorSpace? _fillBase;
        private readonly PdfPageTilingPatternResource? _fillResource;
        private readonly PdfPageShadingPatternResource? _fillShadingResource;
        private readonly Matrix2D _fillPaintTransform;
        private readonly PdfPageColorSpace _fillColorSpace;
        private readonly string? _strokeName;
        private readonly bool _hasStrokeTint;
        private readonly PdfPageColorSpace? _strokeBase;
        private readonly PdfPageTilingPatternResource? _strokeResource;
        private readonly PdfPageShadingPatternResource? _strokeShadingResource;
        private readonly Matrix2D _strokePaintTransform;
        private readonly PdfPageColorSpace _strokeColorSpace;
        private readonly Matrix2D _invocationTransform;
        private readonly PdfPageClipPath? _invocationClip;

        internal XObjectInvocationStateKey(PdfPageXObjectInvocation invocation) {
            _fillName = invocation.FillPattern?.Name;
            _hasFillTint = invocation.FillPattern?.Tint.HasValue == true;
            _fillBase = invocation.FillPatternBaseColorSpace;
            _fillResource = invocation.FillPattern?.TilingPattern;
            _fillShadingResource = invocation.FillPattern?.ShadingPattern;
            _fillPaintTransform = invocation.FillPattern?.PaintTransform ?? Matrix2D.Identity;
            _fillColorSpace = invocation.FillColorSpace;
            _strokeName = invocation.StrokePattern?.Name;
            _hasStrokeTint = invocation.StrokePattern?.Tint.HasValue == true;
            _strokeBase = invocation.StrokePatternBaseColorSpace;
            _strokeResource = invocation.StrokePattern?.TilingPattern;
            _strokeShadingResource = invocation.StrokePattern?.ShadingPattern;
            _strokePaintTransform = invocation.StrokePattern?.PaintTransform ?? Matrix2D.Identity;
            _strokeColorSpace = invocation.StrokeColorSpace;
            _invocationTransform = invocation.Transform;
            _invocationClip = invocation.ClipPath;
        }

        public bool Equals(XObjectInvocationStateKey other) =>
            Same(_invocationTransform, other._invocationTransform) &&
            Same(_fillName, other._fillName) &&
            Same(_hasFillTint, other._hasFillTint) &&
            Same(_fillBase, other._fillBase) &&
            Same(_fillResource, other._fillResource) &&
            Same(_fillShadingResource, other._fillShadingResource) &&
            Same(_fillPaintTransform, other._fillPaintTransform) &&
            Same(_fillColorSpace, other._fillColorSpace) &&
            Same(_strokeName, other._strokeName) &&
            Same(_hasStrokeTint, other._hasStrokeTint) &&
            Same(_strokeBase, other._strokeBase) &&
            Same(_strokeResource, other._strokeResource) &&
            Same(_strokeShadingResource, other._strokeShadingResource) &&
            Same(_strokePaintTransform, other._strokePaintTransform) &&
            Same(_strokeColorSpace, other._strokeColorSpace) &&
            Same(_invocationClip, other._invocationClip);

        public override bool Equals(object? obj) => obj is XObjectInvocationStateKey other && Equals(other);

        public override int GetHashCode() {
            unchecked {
                int hash = _invocationTransform.GetHashCode();
                hash = (hash * 31) + _fillPaintTransform.GetHashCode();
                hash = (hash * 31) + _strokePaintTransform.GetHashCode();
                hash = (hash * 31) + (_fillName is null ? 0 : StringComparer.Ordinal.GetHashCode(_fillName));
                hash = (hash * 31) + (_strokeName is null ? 0 : StringComparer.Ordinal.GetHashCode(_strokeName));
                return hash;
            }
        }

        private static bool Same<T>(T left, T right) => EqualityComparer<T>.Default.Equals(left, right);
    }
}
