using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfAppearancePlacementTests {
    [Theory]
    [InlineData(0.0001)]
    [InlineData(1E-20)]
    [InlineData(-1.23456789E20)]
    [InlineData(double.Epsilon)]
    [InlineData(double.MaxValue)]
    public void GeometryNumbersRoundTripWithoutExponentNotation(double value) {
        string operand = PdfSyntaxEscaper.Number(value);
        Assert.DoesNotContain("E", operand);
        Assert.Equal(value, double.Parse(operand, System.Globalization.CultureInfo.InvariantCulture));
    }

    [Theory]
    [InlineData(1, 0, 0, 1, 0, 0)]
    [InlineData(2, 0, 0, 2, 4, 6)]
    [InlineData(0, 1, -1, 0, 15, 25)]
    [InlineData(1, 0.5, 0.25, 1, -15, 25)]
    [InlineData(-1, 0, 0, 1, 0, 0)]
    public void TransformedBoundsFitTheWidgetWithoutErasingOrientation(double a, double b, double c, double d, double e, double f) {
        var dictionary = new PdfDictionary();
        dictionary.Items["BBox"] = Numbers(10, 20, 190, 42);
        dictionary.Items["Matrix"] = Numbers(a, b, c, d, e, f);
        Matrix2D outer = PdfAppearancePlacement.Read(dictionary, value => value, 30, 300, 180, 22, out Matrix2D matrix);
        Matrix2D combined = Matrix2D.Multiply(outer, matrix);
        var corners = new[] { combined.Transform(10, 20), combined.Transform(190, 20),
            combined.Transform(10, 42), combined.Transform(190, 42) };
        Assert.Equal(30, corners.Min(point => point.X), 8);
        Assert.Equal(210, corners.Max(point => point.X), 8);
        Assert.Equal(300, corners.Min(point => point.Y), 8);
        Assert.Equal(322, corners.Max(point => point.Y), 8);
        Assert.Equal(Math.Sign(b), Math.Sign(combined.B));
        Assert.Equal(Math.Sign(c), Math.Sign(combined.C));
    }

    [Fact]
    public void OverflowedExtentsCannotCollapseAnAppearanceToZeroScale() {
        var dictionary = new PdfDictionary();
        dictionary.Items["BBox"] = Numbers(-double.MaxValue, 0, double.MaxValue, 22);
        Assert.Throws<InvalidOperationException>(() => PdfAppearancePlacement.Read(dictionary,
            value => value, 30, 300, 180, 22, out _));
    }

    [Fact]
    public void DegenerateAppearanceCannotGenerateInfiniteContentCoordinates() {
        var dictionary = new PdfDictionary();
        dictionary.Items["BBox"] = Numbers(0, 0, 180, 22);
        dictionary.Items["Matrix"] = Numbers(0, 0, 0, 0, 0, 0);
        Assert.Throws<InvalidOperationException>(() => PdfAppearancePlacement.Read(dictionary,
            value => value, 30, 300, 180, 22, out _));
    }

    private static PdfArray Numbers(params double[] values) {
        var array = new PdfArray();
        foreach (double value in values) array.Items.Add(new PdfNumber(value));
        return array;
    }
}
