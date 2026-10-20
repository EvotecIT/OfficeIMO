using OfficeIMO.Html.Pdf;
using OfficeIMO.Tests.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlPdfAccessibilityReportTests {
    [Theory]
    [InlineData("<svg width='40' height='20' aria-labelledby='description'><rect width='40' height='20' fill='green'/></svg>")]
    [InlineData("<svg width='40' height='20'><title>Service status</title><rect width='40' height='20' fill='green'/></svg>")]
    [InlineData("<img width='40' height='20' alt='' aria-labelledby='description' src='data:image/svg+xml,%3Csvg xmlns=%22http://www.w3.org/2000/svg%22 width=%2240%22 height=%2220%22%3E%3Crect width=%2240%22 height=%2220%22 fill=%22green%22/%3E%3C/svg%3E'>")]
    public void ImageNamesUseTheSharedAccessibleNameResolution(string image) {
        HtmlPdfRenderRequestResult rendered = CreateReport("<p id='description'>Service status</p>" + image);

        var figures = PdfReadDocument.Open(rendered.ToBytes()).TaggedContent!.StructureElements
            .Where(element => element.StructureType == "Figure");

        Assert.Contains(figures, figure => figure.AlternateText == "Service status");
    }

    [Theory]
    [InlineData("role='presentation'")]
    [InlineData("role='none'")]
    [InlineData("aria-hidden='true'")]
    public void DecorativeLinkedSvgPreservesTheSurroundingLinkWithoutAddingAFigure(string attributes) {
        HtmlPdfRenderRequestResult rendered = CreateReport(
            "<p><a href='https://example.test/details'><svg width='40' height='20' " + attributes +
            "><rect width='40' height='20' fill='green'/></svg>Service details</a></p>");
        byte[] bytes = rendered.ToBytes();
        var tags = PdfReadDocument.Open(bytes).TaggedContent!.StructureElements;

        Assert.DoesNotContain(tags, element => element.StructureType == "Figure");
        Assert.Contains(tags, element => element.StructureType == "Link" && element.ObjectReferenceCount > 0);
        Assert.Contains("Service details", PdfReadDocument.Open(bytes).ExtractText());
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void DecorativeImageOnlyAnchorsRetainTheirNamedLink(bool raster, bool internalLink) {
        string image = raster
            ? "<img width='40' height='20' alt='' src='data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 2)) + "'>"
            : "<svg width='40' height='20' role='presentation'><rect width='40' height='20' fill='green'/></svg>";
        string href = internalLink ? "#target" : "https://example.test/home";
        var rendered = CreateReport("<p><a aria-label='Home' href='" + href + "'>" + image + "</a></p><p id='target'>Target</p>");
        var read = PdfReadDocument.Open(rendered.ToBytes());
        var tags = read.TaggedContent!.StructureElements;
        Assert.DoesNotContain(tags, element => element.StructureType == "Figure");
        Assert.Contains(tags, element => element.StructureType == "Link" && element.ObjectReferenceCount > 0);
        var link = Assert.Single(read.RawStructure().Objects, item => item.Value.Entries.TryGetValue("Subtype", out var type) && type.Text == "Link");
        Assert.Equal("Home", link.Value.Entries["Contents"].Text);
        Assert.Equal(internalLink ? "GoTo" : "URI", link.Value.Entries["A"].Entries["S"].Text);
    }

    private static HtmlPdfRenderRequestResult CreateReport(string body) =>
        OfficeIMO.Examples.Html.Html.CreateAccessibleReport(
            "<html lang='en'><head><title>Image semantics</title></head><body><h1>Image semantics</h1>" + body + "</body></html>",
            File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "Carlito-Regular.ttf")),
            File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "Carlito-Bold.ttf")));

    [Fact]
    public void CustomCssCountersRetainListMetadataAndPaintedLabels() {
        var rendered = CreateReport("<ol style='list-style-type:lower-greek'><li>First</li><li>Second</li></ol>");
        var read = PdfReadDocument.Open(rendered.ToBytes());
        var list = Assert.Single(read.RawStructure().Objects, item => item.Value.Entries.TryGetValue("S", out var role) && role.Text == "L");
        Assert.Equal("None", list.Value.Entries["A"].Entries["ListNumbering"].Text);
        Assert.Contains("α", read.ExtractText());
        Assert.Contains("β", read.ExtractText());
    }

    [Theory]
    [InlineData(PdfComplianceProfile.PdfUa1)]
    [InlineData(PdfComplianceProfile.PdfUa2)]
    public void AccessibleReportRetainsItsSemanticsAcrossPages(PdfComplianceProfile profile) {
        string html = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Reports", "accessibility-review.html"));
        byte[] regular = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "Carlito-Regular.ttf"));
        byte[] bold = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "Carlito-Bold.ttf"));

        HtmlPdfRenderRequestResult rendered = OfficeIMO.Examples.Html.Html.CreateAccessibleReport(html, regular, bold, profile);
        byte[] pdf = rendered.ToBytes();

        PdfDocumentInfo info = PdfInspector.Inspect(pdf);
        Assert.Equal(2, info.PageCount);
        Assert.Equal("en", info.CatalogLanguage);
        Assert.Equal("Service accessibility review", info.Metadata.Title);
        var tags = Assert.IsType<PdfTaggedContentInfo>(info.TaggedContent).StructureElements;
        Assert.Single(tags, element => element.StructureType == "H1");
        Assert.Equal(3, tags.Count(element => element.StructureType == "H2"));
        Assert.Equal(2, tags.Count(element => element.StructureType == "H3"));
        Assert.Equal(6, tags.Count(element => element.StructureType == "TH"));
        Assert.Equal(6, tags.Count(element => element.StructureType == "TD"));
        Assert.Equal(7, tags.Count(element => element.StructureType == "LI"));
        Assert.Contains(tags, element => element.StructureType == "Figure" && element.AlternateText == "All three service checks passed");
        Assert.Contains(tags, element => element.StructureType == "Link" && element.ObjectReferenceCount > 0);
        Assert.Contains(tags, element => element.StructureType == "Form" && element.ObjectReferenceCount > 0);
        var read = PdfReadDocument.Open(pdf);
        Assert.Equal(new int?[] { 1, 2 }, read.NamedDestinations.Select(destination => destination.PageNumber).OrderBy(page => page));
        var raw = read.RawStructure();
        foreach (var destination in read.NamedDestinations)
            Assert.Contains(raw.Objects, item => item.Value.Entries.TryGetValue("Subtype", out var subtype) && subtype.Text == "Link"
                && item.Value.Entries["A"].Entries["D"].Text == destination.Name);

        string text = PdfDocument.Load(pdf).Read().Text;
        Assert.Contains("Engineering", text);
        Assert.Contains("Publishing", text);
        Assert.Contains("Operations", text);
        int firstGroup = text.IndexOf("First reading group", StringComparison.Ordinal);
        int secondGroup = text.IndexOf("Second reading group", StringComparison.Ordinal);
        Assert.True(firstGroup >= 0 && secondGroup > firstGroup, text);
    }
}
