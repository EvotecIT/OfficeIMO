using OfficeIMO.Pdf;
using System;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfHeadingHierarchyTests {
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
