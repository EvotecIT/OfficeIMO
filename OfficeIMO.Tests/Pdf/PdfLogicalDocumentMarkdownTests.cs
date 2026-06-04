using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfLogicalDocumentTests {
    [Fact]
    public void ToMarkdown_RendersLogicalHeadingsParagraphsListsTablesAndImages() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .H1("Logical Heading")
            .Paragraph(p => p.Text("Logical readback marker."))
            .Bullets(new[] { "Detected logical bullet" })
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha", "2" },
                new[] { "B-200", "Beta", "14" }
            }, style: new PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 70, 170, 60 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .Image(CreateMinimalRgbPng(), 18, 18)
            .ToBytes();

        PdfLogicalDocument logical = PdfLogicalDocument.Load(pdf, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        });

        string markdown = logical.ToMarkdown();
        string normalizedMarkdown = Normalize(markdown);

        Assert.Contains("# Logical Heading", markdown, StringComparison.Ordinal);
        Assert.Contains("Logicalreadbackmarker.", normalizedMarkdown, StringComparison.Ordinal);
        Assert.Contains("-Detectedlogicalbullet", normalizedMarkdown, StringComparison.Ordinal);
        Assert.Contains("| Code | Name | Qty |", markdown, StringComparison.Ordinal);
        Assert.Contains("| --- | --- | --- |", markdown, StringComparison.Ordinal);
        Assert.Contains("| A-100 | Alpha | 2 |", markdown, StringComparison.Ordinal);
        Assert.Contains("[Image: page 1", markdown, StringComparison.Ordinal);
        AssertContainsInOrder(normalizedMarkdown,
            "#LogicalHeading",
            "Logicalreadbackmarker.",
            "-Detectedlogicalbullet",
            "|Code|Name|Qty|",
            "[Image:page1");

        string withoutImages = logical.ToMarkdown(new PdfLogicalMarkdownOptions {
            IncludeImagePlaceholders = false
        });
        Assert.DoesNotContain("[Image:", withoutImages, StringComparison.Ordinal);
    }

    [Fact]
    public void ToMarkdown_EscapesMarkdownControlSyntaxFromPdfText() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 260,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Paragraph(p => p.Text("# Literal heading marker"))
            .Paragraph(p => p.Text("[not a link](https://example.test)"))
            .ToBytes();

        string markdown = PdfLogicalDocument.Load(pdf, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        }).ToMarkdown();

        string normalized = Normalize(markdown);
        Assert.Contains("\\#Literalheadingmarker", normalized, StringComparison.Ordinal);
        Assert.Contains("\\[notalink\\](https://example.test)", normalized, StringComparison.Ordinal);
    }

    [Fact]
    public void ToMarkdown_DoesNotRenderLeaderRowsTwiceWhenTableAlreadyContainsThem() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 260,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Paragraph(p => p.Text("Chapter One ........ 3"))
            .ToBytes();

        string markdown = PdfLogicalDocument.Load(pdf, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        }).ToMarkdown();

        Assert.Equal(1, CountOccurrences(markdown, "Chapter One"));
    }
}
