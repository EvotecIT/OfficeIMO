using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeTrueTypeFont {
    private double _trackingEvaluationScale = 1D;

    // A raster resolution change scales geometry, while tracking retains the authored text size.
    internal OfficeTrueTypeFont ForOutputScale(double scale) {
        if (_tracking == null || scale == _trackingEvaluationScale) return this;
        if (scale <= 0 || double.IsNaN(scale) || double.IsInfinity(scale)) throw new ArgumentOutOfRangeException(nameof(scale));
        var instance = (OfficeTrueTypeFont)MemberwiseClone();
        instance._opticalSizeFonts = null;
        instance._trackingEvaluationScale = scale;
        return instance;
    }

    private double HorizontalTracking(double fontSize) =>
        (_tracking?.GetAdjustment(fontSize / _trackingEvaluationScale) ?? 0D) * ScaleFor(fontSize);

    private bool[]? TrackingBoundaries(string text, IReadOnlyList<int> indexes, bool negative = false) {
        if (_tracking == null || indexes.Count == 0) return null;
        return OfficeOpenTypeTracking.GetBoundaries(text, indexes, negative);
    }
}
