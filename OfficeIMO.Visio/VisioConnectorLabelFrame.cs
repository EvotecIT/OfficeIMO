using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

internal enum VisioConnectorLabelAnchorKind { Path, Page, Native }

/// <summary>Immutable source coordinates, retained independently of the projected frame and graph copies.</summary>
internal sealed class VisioConnectorNativeTextFrame {
    internal VisioConnectorNativeTextFrame(double x, double y, double angle) { Pin = new OfficePoint(x, y); Angle = angle; }
    internal OfficePoint Pin { get; }
    internal double Angle { get; }
}

/// <summary>Canonical cached connector text-frame projection for writing, drawing, layout and inspection.</summary>
internal static class VisioConnectorLabelFrame {
    internal static VisioConnectorLabelPlacement? ResolvePlacement(VisioConnector connector) {
        VisioConnectorLabelPlacement? source = connector.LabelPlacement;
        if (source == null) return null;
        if (source.AnchorKind == VisioConnectorLabelAnchorKind.Path) {
            // Native files also retain a cached page pin. It must not freeze a restored path anchor.
            if (!source.AbsolutePinX.HasValue && !source.AbsolutePinY.HasValue) return source;
            var path = source.Clone(); path.SetAbsolutePin(null, null); return path;
        }
        if (source.AnchorKind != VisioConnectorLabelAnchorKind.Native || source.NativeFrame == null ||
            connector.NativeGeometry?.AppliesTo(connector) != true) return source;
        var placement = source.Clone();
        OfficePoint pin = connector.NativeGeometry.TransformSourcePoint(connector, source.NativeFrame.Pin);
        (double scale, _) = connector.NativeGeometry.GetTransform(connector);
        placement.SetAbsolutePin(pin.X, pin.Y);
        double dimensionScale = scale / source.NativeDimensionScale;
        placement.Width *= dimensionScale; placement.Height *= dimensionScale;
        placement.LocPinX *= dimensionScale; placement.LocPinY *= dimensionScale;
        placement.AnchorKind = VisioConnectorLabelAnchorKind.Page;
        placement.NativeFrame = null;
        return placement;
    }

    /// <summary>Stores a layout move in source coordinates without discarding native frame intent.</summary>
    internal static void ApplyLayoutPlacement(VisioConnector connector, VisioConnectorLabelPlacement projected) {
        VisioConnectorLabelPlacement? source = connector.LabelPlacement;
        if (source?.AnchorKind != VisioConnectorLabelAnchorKind.Native || source.NativeFrame == null ||
            connector.NativeGeometry?.AppliesTo(connector) != true) {
            connector.LabelPlacement = projected;
            return;
        }
        OfficePoint current = connector.NativeGeometry.TransformSourcePoint(connector, source.NativeFrame.Pin);
        (double scale, double angle) = connector.NativeGeometry.GetTransform(connector);
        if (scale < 1e-12)
            throw new System.NotSupportedException("A collapsed native connector frame requires explicit label placement before layout editing.");
        double dx = (projected.AbsolutePinX!.Value - current.X) / scale;
        double dy = (projected.AbsolutePinY!.Value - current.Y) / scale;
        double x = source.NativeFrame.Pin.X + dx * System.Math.Cos(angle) + dy * System.Math.Sin(angle);
        double y = source.NativeFrame.Pin.Y - dx * System.Math.Sin(angle) + dy * System.Math.Cos(angle);
        var edited = source.Clone();
        edited.SetAbsolutePin(x, y);
        edited.NativeFrame = new VisioConnectorNativeTextFrame(x, y, source.NativeFrame.Angle);
        connector.LabelPlacement = edited;
    }

    /// <summary>Rebases physical text measurements to the current native scale, retaining pin and angle binding.</summary>
    internal static void SetMeasuredDimensions(VisioConnector connector, VisioConnectorLabelPlacement placement, double width, double height) {
        if (placement.AnchorKind == VisioConnectorLabelAnchorKind.Native && placement.NativeFrame != null &&
            connector.NativeGeometry?.AppliesTo(connector) == true) {
            double scale = connector.NativeGeometry.GetTransform(connector).Scale;
            if (scale < 1e-12)
                throw new System.NotSupportedException("A collapsed native connector frame requires explicit label placement before text fitting.");
            double ratio = scale / placement.NativeDimensionScale;
            placement.LocPinX *= ratio;
            placement.LocPinY *= ratio;
            placement.NativeDimensionScale = scale;
        }
        placement.Width = width;
        placement.Height = height;
    }

    internal static double ResolveAngle(VisioConnector connector) {
        double angle = connector.TextStyle?.TextAngle ?? 0;
        VisioConnectorLabelPlacement? placement = connector.LabelPlacement;
        if (placement?.AnchorKind == VisioConnectorLabelAnchorKind.Native && !placement.KeepPageAngle && placement.NativeFrame != null &&
            angle == placement.NativeFrame.Angle && connector.NativeGeometry?.AppliesTo(connector) == true)
            angle += connector.NativeGeometry.GetTransform(connector).Angle;
        return angle;
    }
}
