using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfSectionTableOfContentsTests {
    [Fact]
    public void TableOfContents_StabilizesSectionPagesAndBuildsInternalNavigation() {
        var introduction = new PdfSectionReference();
        var details = new PdfSectionReference();

        byte[] bytes = PdfDocument.Create()
            .TableOfContents()
            .Section(
                "Introduction",
                item => item.Paragraph(paragraph => paragraph.Text("Introduction body")),
                new PdfSectionOptions {
                    DestinationName = "intro",
                    StartOnNewPage = true,
                    Reference = introduction
                })
            .Section(
                "Details",
                item => item.Paragraph(paragraph => paragraph.Text("Details body")),
                new PdfSectionOptions {
                    DestinationName = "details",
                    StartOnNewPage = true,
                    Reference = details
                })
            .ToBytes();

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfLogicalDocument logical = PdfDocument.Load(bytes).Read.Logical();
        IReadOnlyList<string> textByPage = PdfDocument.Load(bytes).Read.TextByPage();

        Assert.Equal(3, info.PageCount);
        Assert.Equal(2, introduction.PageNumber);
        Assert.Equal(3, details.PageNumber);
        Assert.Equal("intro", introduction.DestinationName);
        Assert.Contains("intro", info.NamedDestinationNames);
        Assert.Contains("details", info.NamedDestinationNames);
        Assert.NotEmpty(logical.GetLinksByDestinationName("intro"));
        Assert.NotEmpty(logical.GetLinksByDestinationName("details"));
        Assert.Contains("Introduction", textByPage[0], StringComparison.Ordinal);
        Assert.Contains("2", textByPage[0], StringComparison.Ordinal);
        Assert.Contains("Details", textByPage[0], StringComparison.Ordinal);
        Assert.Contains("3", textByPage[0], StringComparison.Ordinal);
        Assert.Contains("Introduction body", textByPage[1], StringComparison.Ordinal);
        Assert.Contains("Details body", textByPage[2], StringComparison.Ordinal);
    }

    [Fact]
    public void TableOfContents_RespectsHierarchyAndCustomPageFormatting() {
        byte[] bytes = PdfDocument.Create()
            .TableOfContents(new PdfTableOfContentsOptions {
                MinimumLevel = 2,
                MaximumLevel = 2,
                PageNumberFormatter = page => "p" + page
            })
            .Section("Excluded", _ => { }, new PdfSectionOptions { Level = 1, StartOnNewPage = true })
            .Section("Included", _ => { }, new PdfSectionOptions { Level = 2, StartOnNewPage = true })
            .ToBytes();

        string firstPage = PdfDocument.Load(bytes).Read.TextByPage()[0];

        Assert.DoesNotContain("Excluded", firstPage, StringComparison.Ordinal);
        Assert.Contains("Included", firstPage, StringComparison.Ordinal);
        Assert.Contains("p3", firstPage, StringComparison.Ordinal);
    }
}
