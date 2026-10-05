using OfficeIMO.Pdf;
using System;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfHeadingHierarchyTests {
    [Theory]
    [InlineData(4, true)]
    [InlineData(7, true)]
    [InlineData(9, true)]
    [InlineData(4, false)]
    [InlineData(7, false)]
    [InlineData(9, false)]
    public void SparseHeadingLevels_KeepTaggedDepthAndUseOutlinesAsFallback(int level, bool tagged) {
        var document = PdfDocument.Create(new PdfOptions { CreateOutlineFromHeadings = true });
        if (tagged) document.TaggedPdfCatalogMarkers();
        document.Content.Heading(1, "ParentHeading").Heading(level, "SparseHeading");
        byte[] bytes = document.ToBytes();
        Assert.Equal(2, Assert.Single(Assert.Single(PdfInspector.Inspect(bytes).Outlines).Children).Level);
        var heading = Assert.Single(PdfDocumentReadResult.Load(bytes).Headings, item => item.Text == "SparseHeading");
        Assert.Equal(tagged ? level : 2, heading.Level);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AdapterHeading_RetainsStyledRunsWithoutChangingOutlineTitle(bool columns) {
        var document = PdfDocument.Create(new PdfOptions { CreateOutlineFromHeadings = true });
        document.TaggedPdfCatalogMarkers();
        PdfParagraphBuilder captured = null!;
        const string uri = "https://example.com/inline-heading";
        void AddHeading(PdfContentBuilder content) => content.Heading(9, "Outline title", builder => {
            captured = builder;
            builder.Text("PlainMarker ").Italic(true).Text("ItalicMarker ").Italic(false)
                .Link("LinkedMarker", uri).Baseline(PdfTextBaseline.Superscript).Text("2");
        }, PdfAlign.Left, null, new PdfHeadingStyle { FontSize = 12 });
        if (columns) document.Content.Columns(AddHeading);
        else AddHeading(document.Content);
        captured.Text("Late mutation");

        byte[] bytes = document.ToBytes();
        var spans = PdfReadDocument.Open(bytes).Pages.SelectMany(page => page.GetTextSpans()).ToArray();
        Assert.Contains(spans, span => span.Text.Contains("ItalicMarker") && span.IsItalic);
        Assert.DoesNotContain(spans, span => span.Text.Contains("Late mutation"));
        var logical = PdfDocumentReadResult.Load(bytes);
        Assert.Single(logical.GetLinksByUri(uri));
        Assert.Equal(9, Assert.Single(logical.Headings, heading => heading.Text.Contains("PlainMarker")).Level);
        Assert.Equal("Outline title", Assert.Single(PdfInspector.Inspect(bytes).Outlines).Title);
    }

    [Fact]
    public void Heading_PreservesNineAuthoredLevelsWithEqualFontSizes() {
        var document = PdfDocument.Create(new PdfOptions { CreateOutlineFromHeadings = true });
        document.TaggedPdfCatalogMarkers();
        for (int level = 1; level <= 9; level++) {
            document.Content.Heading(level, "HeadingLevel" + level, style: new PdfHeadingStyle { FontSize = 12 });
        }

        byte[] bytes = document.ToBytes();
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(bytes);
        PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(logical.TaggedContent);
        for (int level = 1; level <= 9; level++) {
            var heading = Assert.Single(logical.Headings, h => h.Text == "HeadingLevel" + level);
            Assert.Equal(level, heading.Level);
            Assert.Contains("H" + level, tagged.StructureTypes);
        }
        Assert.Equal(3, tagged.RoleMap.Count);
        foreach (string role in new[] { "H7", "H8", "H9" }) Assert.Equal("H6", tagged.RoleMap[role]);

        PdfOutlineItem outline = Assert.Single(PdfInspector.Inspect(bytes).Outlines);
        for (int level = 1; level <= 9; level++) {
            Assert.Equal("HeadingLevel" + level, outline.Title);
            Assert.Equal(level, outline.Level);
            if (level < 9) outline = Assert.Single(outline.Children);
            else Assert.Empty(outline.Children);
        }
    }

    [Fact]
    public void DeepSectionHeading_UsesTheSameRoleMappingAndSemanticDepth() {
        var document = PdfDocument.Create().TaggedPdfCatalogMarkers();
        document.Content.Section("DeepSectionHeading", _ => { }, new PdfSectionOptions { Level = 9 });
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(document.ToBytes());
        Assert.Equal(9, Assert.Single(logical.Headings, h => h.Text == "DeepSectionHeading").Level);
        Assert.Equal("H6", Assert.IsType<PdfTaggedContentInfo>(logical.TaggedContent).RoleMap["H9"]);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(10)]
    public void Heading_RejectsUnsupportedHierarchyLevels(int level) {
        var document = PdfDocument.Create();
        Assert.Throws<ArgumentOutOfRangeException>(() => document.Content.Heading(level, "InvalidHeading"));
    }
}
