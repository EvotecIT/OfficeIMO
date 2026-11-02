using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfThemePageAndListDefaultsTests {
    [Fact]
    public void ThemeMarginsFollowTheSelectedThemeThroughCloningAndExplicitSidesKeepTheirValues() {
        var theme = OfficeTheme.Modern.WithPage(new OfficeThemePage(
            OfficeLength.Points(29), OfficeLength.Points(31), OfficeLength.Points(37), OfficeLength.Points(23)));
        var options = new PdfCreateOptions { Theme = theme, MarginRight = OfficeLength.Points(9) };
        Assert.Equal(23D, options.MarginLeft.ToPoints());
        Assert.Equal(29D, options.MarginTop.ToPoints());
        Assert.Equal(37D, options.MarginBottom.ToPoints());
        var clone = options.Clone();
        clone.Theme = OfficeTheme.Minimal;
        Assert.Equal(OfficeTheme.Minimal.Page.MarginLeft.ToPoints(), clone.MarginLeft.ToPoints());
        Assert.Equal(OfficeTheme.Minimal.Page.MarginTop.ToPoints(), clone.MarginTop.ToPoints());
        Assert.Equal(OfficeTheme.Minimal.Page.MarginBottom.ToPoints(), clone.MarginBottom.ToPoints());
        Assert.Equal(9D, clone.MarginRight.ToPoints());
        Assert.Equal(23D, options.MarginLeft.ToPoints());
        using var document = PdfDocument.Create(clone).Paragraph(paragraph => paragraph.Text("Margin anchor"));
        using var independent = UglyToad.PdfPig.PdfDocument.Open(document.ToBytes());
        Assert.Equal(clone.MarginLeft.ToPoints(), independent.GetPage(1).Letters[0].StartBaseLine.X, 2);
    }

    [Fact]
    public void ThemeBulletsLeaveAReadableGapAndAnExplicitMarkerGapStillWins() {
        double Gap(PdfListStyle? style) {
            using var document = PdfDocument.Create().Bullets(new[] { "Readable item" }, style: style);
            using var independent = UglyToad.PdfPig.PdfDocument.Open(document.ToBytes());
            var letters = independent.GetPage(1).Letters;
            var marker = Assert.Single(letters, letter => letter.Value == "•");
            var text = Assert.Single(letters, letter => letter.Value == "R");
            return text.GlyphRectangle.Left - marker.GlyphRectangle.Right;
        }
        double themed = Gap(null), explicitGap = Gap(new PdfListStyle { MarkerGap = OfficeLength.Points(1) });
        Assert.True(themed >= OfficeTheme.Default.TypeScale.Body.ToPoints() * .45D, "Bullet gap was " + themed + " points.");
        Assert.True(themed > explicitGap + 3D, "The explicit list gap did not override the theme default.");
    }
}
