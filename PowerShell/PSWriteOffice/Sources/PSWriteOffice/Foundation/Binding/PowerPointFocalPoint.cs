using System;
using System.Globalization;
using System.Management.Automation;

namespace PSWriteOffice.Foundation;

/// <summary>Fractional image coordinates kept in view when cropping an image to fill its frame.</summary>
public sealed class PowerPointFocalPoint {
    /// <summary>Creates a focal point with both coordinates in the inclusive range zero to one.</summary>
    public PowerPointFocalPoint(double x, double y) {
        if (double.IsNaN(x) || double.IsInfinity(x) || x < 0 || x > 1) throw new ArgumentOutOfRangeException(nameof(x));
        if (double.IsNaN(y) || double.IsInfinity(y) || y < 0 || y > 1) throw new ArgumentOutOfRangeException(nameof(y));
        X = x; Y = y;
    }
    /// <summary>Horizontal position as a fraction of image width.</summary>
    public double X { get; }
    /// <summary>Vertical position as a fraction of image height.</summary>
    public double Y { get; }

    /// <summary>Reads Center, Top, Bottom, Left, Right or an invariant x,y pair such as 0.4,0.3.</summary>
    public static PowerPointFocalPoint From(object value) {
        value = OfficeArgumentValue.Unwrap(value) ?? throw OfficeArgumentValue.Invalid("A focal point is required.");
        if (value is PowerPointFocalPoint point) return point;
        if (value is string text) {
            switch (text.Trim().ToUpperInvariant()) {
                case "CENTER": return new(0.5, 0.5);
                case "TOP": return new(0.5, 0);
                case "BOTTOM": return new(0.5, 1);
                case "LEFT": return new(0, 0.5);
                case "RIGHT": return new(1, 0.5);
            }
            string[] pair = text.Split(',');
            if (pair.Length == 2 && double.TryParse(pair[0], NumberStyles.Float, CultureInfo.InvariantCulture, out double x)
                && double.TryParse(pair[1], NumberStyles.Float, CultureInfo.InvariantCulture, out double y)) return new(x, y);
        }
        throw OfficeArgumentValue.Invalid("Choose Center, Top, Bottom, Left, Right or x,y fractions from zero to one.");
    }
}

/// <summary>Binds a named or fractional native image focal point.</summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public sealed class PowerPointFocalPointTransformationAttribute : ArgumentTransformationAttribute {
    /// <inheritdoc />
    public override object Transform(EngineIntrinsics engineIntrinsics, object inputData) => PowerPointFocalPoint.From(inputData);
}
