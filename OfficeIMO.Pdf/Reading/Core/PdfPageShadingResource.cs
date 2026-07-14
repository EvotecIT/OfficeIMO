using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal readonly struct PdfPageShadingResource {
    public PdfPageShadingResource(double x0, double y0, double x1, double y1, OfficeColor startColor, OfficeColor endColor) {
        IsRadial = false;
        X0 = x0;
        Y0 = y0;
        R0 = 0D;
        X1 = x1;
        Y1 = y1;
        R1 = 0D;
        Stops = new[] { new OfficeGradientStop(0D, startColor), new OfficeGradientStop(1D, endColor) };
    }

    public PdfPageShadingResource(double x0, double y0, double x1, double y1, IReadOnlyList<OfficeGradientStop> stops) {
        IsRadial = false;
        X0 = x0;
        Y0 = y0;
        R0 = 0D;
        X1 = x1;
        Y1 = y1;
        R1 = 0D;
        Stops = SnapshotStops(stops);
    }

    public PdfPageShadingResource(double x0, double y0, double r0, double x1, double y1, double r1, OfficeColor startColor, OfficeColor endColor) {
        IsRadial = true;
        X0 = x0;
        Y0 = y0;
        R0 = r0;
        X1 = x1;
        Y1 = y1;
        R1 = r1;
        Stops = new[] { new OfficeGradientStop(0D, startColor), new OfficeGradientStop(1D, endColor) };
    }

    public PdfPageShadingResource(double x0, double y0, double r0, double x1, double y1, double r1, IReadOnlyList<OfficeGradientStop> stops) {
        IsRadial = true;
        X0 = x0;
        Y0 = y0;
        R0 = r0;
        X1 = x1;
        Y1 = y1;
        R1 = r1;
        Stops = SnapshotStops(stops);
    }

    public bool IsRadial { get; }

    public double X0 { get; }

    public double Y0 { get; }

    public double R0 { get; }

    public double X1 { get; }

    public double Y1 { get; }

    public double R1 { get; }

    public IReadOnlyList<OfficeGradientStop> Stops { get; }

    public OfficeColor StartColor => Stops[0].Color;

    public OfficeColor EndColor => Stops[Stops.Count - 1].Color;

    private static System.Collections.ObjectModel.ReadOnlyCollection<OfficeGradientStop> SnapshotStops(IReadOnlyList<OfficeGradientStop> stops) {
        if (stops == null || stops.Count < 2) {
            throw new System.ArgumentException("PDF shading resources require at least two gradient stops.", nameof(stops));
        }

        var snapshot = new OfficeGradientStop[stops.Count];
        double previousOffset = -1D;
        for (int index = 0; index < stops.Count; index++) {
            OfficeGradientStop stop = stops[index];
            if (stop.Offset < previousOffset) {
                throw new System.ArgumentException("PDF shading gradient stops must be ordered by offset.", nameof(stops));
            }

            snapshot[index] = stop;
            previousOffset = stop.Offset;
        }

        return System.Array.AsReadOnly(snapshot);
    }
}
