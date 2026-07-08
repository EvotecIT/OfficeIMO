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
        StartColor = startColor;
        EndColor = endColor;
    }

    public PdfPageShadingResource(double x0, double y0, double r0, double x1, double y1, double r1, OfficeColor startColor, OfficeColor endColor) {
        IsRadial = true;
        X0 = x0;
        Y0 = y0;
        R0 = r0;
        X1 = x1;
        Y1 = y1;
        R1 = r1;
        StartColor = startColor;
        EndColor = endColor;
    }

    public bool IsRadial { get; }

    public double X0 { get; }

    public double Y0 { get; }

    public double R0 { get; }

    public double X1 { get; }

    public double Y1 { get; }

    public double R1 { get; }

    public OfficeColor StartColor { get; }

    public OfficeColor EndColor { get; }
}
