using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static class PdfPageColorConverter {
    public static OfficeColor FromLab(double lightness, double a, double b) => OfficeColorSpaceConverter.FromLab(lightness, a, b);
    public static OfficeColor FromCalGray(double gray) => OfficeColorSpaceConverter.FromCalibratedGray(gray, 0.9505D, 1D, 1.089D);
    public static OfficeColor FromCalRgb(double red, double green, double blue) => OfficeColorSpaceConverter.FromCalibratedRgb(red, green, blue, 0.9505D, 1D, 1.089D);
    public static OfficeColor FromCalRgb(double red, double green, double blue, PdfPageColorSpace colorSpace) => colorSpace.ConvertCalRgb(red, green, blue);
}
