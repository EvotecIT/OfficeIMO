using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfGradientPlacementConsistencyTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BodyHeaderAndFooterPreserveTheSameNonSquareGradientField(bool transformed) {
        var shape = OfficeShape.Rectangle(200, 100);
        shape.FillGradient = OfficeLinearGradient.DiagonalDown(OfficeColor.Red, OfficeColor.Blue);
        shape.StrokeWidth = 0;
        if (transformed) shape.Transform = OfficeTransform.RotateDegrees(17, 100, 50);
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(500), PageHeight = global::OfficeIMO.OfficeLength.Points(700),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(50), MarginRight = global::OfficeIMO.OfficeLength.Points(50), MarginTop = global::OfficeIMO.OfficeLength.Points(150), MarginBottom = global::OfficeIMO.OfficeLength.Points(150)
            })
            .Compose(document => document.Page(page => page
                .Header(header => header.Shape(shape))
                .Footer(footer => footer.Shape(shape))
                .Content(content => content.Spacer(150).Shape(shape))))
            .ToBytes();

        MatchCollection coordinates = Regex.Matches(Encoding.ASCII.GetString(bytes), @"/Coords\s*\[([^\]]+)\]");
        Assert.NotEmpty(coordinates.Cast<Match>());
        foreach (Match match in coordinates) {
            double[] values = match.Groups[1].Value.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(value => double.Parse(value, CultureInfo.InvariantCulture)).ToArray();
            Assert.Equal(4, values.Length);
            // Untransformed PDF coordinates point upward; transformed shapes keep
            // their top-down local coordinates until the single placement matrix.
            // The gradient's normal must survive the non-square coordinate scale.
            Assert.InRange(values[2] - values[0], 79.999, 80.001);
            double expectedY = transformed ? 160D : -160D;
            Assert.InRange(values[3] - values[1], expectedY - .001D, expectedY + .001D);
        }
    }
}
