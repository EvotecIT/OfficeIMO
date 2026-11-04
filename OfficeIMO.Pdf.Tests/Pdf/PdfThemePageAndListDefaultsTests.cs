using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfThemePageAndListDefaultsTests {
    [Fact]
    public void ThemeMarginsFollowTheSelectedThemeThroughCloningAndExplicitSidesKeepTheirValues() {
        var theme = OfficeTheme.Modern.WithPage(new OfficeThemePage(
            OfficeLength.Points(29), OfficeLength.Points(31), OfficeLength.Points(37), OfficeLength.Points(23)));
        var options = new PdfCreateOptions { Theme = theme, PageSetup = { Margins = { Right = OfficeLength.Points(9) } } };
        Assert.Equal(23D, options.PageSetup.Margins.Left.ToPoints());
        Assert.Equal(29D, options.PageSetup.Margins.Top.ToPoints());
        Assert.Equal(37D, options.PageSetup.Margins.Bottom.ToPoints());
        var clone = options.Clone();
        clone.Theme = OfficeTheme.Minimal;
        Assert.Equal(OfficeTheme.Minimal.Page.MarginLeft.ToPoints(), clone.PageSetup.Margins.Left.ToPoints());
        Assert.Equal(OfficeTheme.Minimal.Page.MarginTop.ToPoints(), clone.PageSetup.Margins.Top.ToPoints());
        Assert.Equal(OfficeTheme.Minimal.Page.MarginBottom.ToPoints(), clone.PageSetup.Margins.Bottom.ToPoints());
        Assert.Equal(9D, clone.PageSetup.Margins.Right.ToPoints());
        Assert.Equal(23D, options.PageSetup.Margins.Left.ToPoints());
        using var document = PdfDocument.Create(clone).Paragraph(paragraph => paragraph.AddText("Margin anchor"));
        using var independent = UglyToad.PdfPig.PdfDocument.Open(document.ToBytes());
        Assert.Equal(clone.PageSetup.Margins.Left.ToPoints(), independent.GetPage(1).Letters[0].StartBaseLine.X, 2);
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
