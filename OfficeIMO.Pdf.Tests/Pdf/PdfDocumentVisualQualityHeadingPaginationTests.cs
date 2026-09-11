using System;
using System.Linq;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Theory]
    [InlineData(1, false)]
    [InlineData(1, true)]
    [InlineData(2, false)]
    [InlineData(2, true)]
    [InlineData(3, false)]
    [InlineData(3, true)]
    public void ContinuedHeadingRetainsOneLogicalHeadingElement(int level, bool linked) {
        var options = new PdfOptions { PageWidth = 220, PageHeight = 160, MarginLeft = 25, MarginRight = 25, MarginTop = 25, MarginBottom = 25 };
        PdfDocument document = PdfDocument.Create(options).TaggedPdfCatalogMarkers();
        string text = string.Join(" ", Enumerable.Repeat("Heading content", 100)) + " FINAL-MARKER";
        string? link = linked ? "https://example.test/heading" : null;
        if (level == 1) document.H1(text, linkUri: link);
        else if (level == 2) document.H2(text, linkUri: link);
        else document.H3(text, linkUri: link);
        byte[] bytes = document.ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.True(pdf.NumberOfPages > 2);
        Assert.Contains("FINAL-MARKER", string.Join("", pdf.GetPages().Select(page => page.Text)), StringComparison.Ordinal);
        string content = System.Text.Encoding.ASCII.GetString(bytes);
        Assert.Equal(1, CountOccurrences(content, "/Type /StructElem /S /H" + level + " "));
        Assert.Equal(pdf.NumberOfPages, CountOccurrences(content, "/StructParents "));
        PdfTaggedContentInfo tagged = PdfReadDocument.Open(bytes).TaggedContent!;
        Assert.NotNull(tagged);
        PdfStructureElementInfo heading = Assert.Single(tagged.StructureElements.Where(element => element.StructureType == "H" + level));
        Assert.Null(heading.PageObjectNumber);
        PdfStructureElementInfo[] children = tagged.StructureElements.Where(element => heading.ChildElementObjectNumbers.Contains(element.ObjectNumber)).ToArray();
        Assert.Equal(pdf.NumberOfPages, children.Length);
        Assert.All(children, child => {
            Assert.Equal(heading.ObjectNumber, child.ParentObjectNumber);
            Assert.Equal(linked ? "Link" : "Span", child.StructureType);
            Assert.NotNull(Assert.Single(child.MarkedContentReferences).PageObjectNumber);
        });
        Assert.Equal(pdf.NumberOfPages, children.SelectMany(child => child.MarkedContentReferences).Select(reference => reference.PageObjectNumber).Distinct().Count());
    }

    [Theory]
    [InlineData(180, 220)]
    [InlineData(320, 160)]
    public void OversizedHeadingPaginatesTextAndLinksWithinThePageFrame(double pageWidth, double pageHeight) {
        var options = new PdfOptions {
            PageWidth = pageWidth, PageHeight = pageHeight,
            MarginLeft = 25, MarginRight = 25, MarginTop = 25, MarginBottom = 25,
            CreateOutlineFromHeadings = true
        };
        string heading = string.Join(" ", Enumerable.Repeat("Heading text", 100)) + " FINAL-MARKER";
        byte[] bytes = PdfDocument.Create(options).TaggedPdfCatalogMarkers().H1(heading, linkUri: "https://example.test/heading").Paragraph(p => p.Text("Following paragraph")).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.True(pdf.NumberOfPages > 2);
        Assert.Contains("FINAL-MARKER", string.Join("", pdf.GetPages().Select(page => page.Text)), StringComparison.Ordinal);
        foreach (var page in pdf.GetPages()) {
            Assert.NotEmpty(page.Letters);
            Assert.All(page.Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)), letter => {
                Assert.InRange(letter.BoundingBox.Bottom, options.MarginBottom - 1, options.PageHeight - options.MarginTop + 1);
                Assert.InRange(letter.BoundingBox.Top, options.MarginBottom - 1, options.PageHeight - options.MarginTop + 1);
            });
        }
        var rectangles = ExtractLinkRectangles(System.Text.Encoding.ASCII.GetString(bytes));
        Assert.True(rectangles.Count > pdf.NumberOfPages);
        string content = System.Text.Encoding.ASCII.GetString(bytes);
        Assert.Equal(pdf.NumberOfPages, CountOccurrences(content, "/StructParents "));
        Assert.Equal(rectangles.Count, CountOccurrences(content, "/Type /OBJR /Obj "));
        Assert.All(rectangles, rectangle => {
            Assert.InRange(rectangle.Y1, options.MarginBottom - 1, options.PageHeight - options.MarginTop + 1);
            Assert.InRange(rectangle.Y2, options.MarginBottom - 1, options.PageHeight - options.MarginTop + 1);
        });
        string? evidence = Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_PDF_EVIDENCE");
        if (!string.IsNullOrWhiteSpace(evidence)) {
            System.IO.Directory.CreateDirectory(evidence!);
            System.IO.File.WriteAllBytes(System.IO.Path.Combine(evidence!, "heading-" + pageWidth + "-" + pageHeight + ".pdf"), bytes);
        }
    }

    [Theory]
    [InlineData(70)]
    [InlineData(50.1)]
    [InlineData(83.5995)]
    public void ParagraphLineLargerThanTheFrameFailsWithoutUnboundedPagination(double pageHeight) {
        var options = new PdfOptions { PageHeight = pageHeight, MarginTop = 25, MarginBottom = 25, DefaultFontSize = 24 };
        Assert.Throws<ArgumentException>(() => PdfDocument.Create(options).Paragraph(p => p.Text("Too tall")).ToBytes());
    }

    [Fact]
    public void SmallPositiveFrameAcceptsALineThatFits() {
        var options = new PdfOptions { PageHeight = 50.2, MarginTop = 25, MarginBottom = 25, DefaultFontSize = 0.1 };
        byte[] bytes = PdfDocument.Create(options).Paragraph(p => p.Text("Fits")).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.Equal("Fits", pdf.GetPage(1).Text);
    }

    [Fact]
    public void HeadingRequiredTopSpacingCannotBeSilentlyDroppedToFitALine() {
        var options = new PdfOptions { PageHeight = 150, MarginTop = 25, MarginBottom = 25 };
        var style = new PdfHeadingStyle { FontSize = 24, LineHeight = 1.25, SpacingBefore = 90, SpacingAfter = 0, ApplySpacingBeforeAtTop = true };
        Assert.Throws<ArgumentException>(() => PdfDocument.Create(options).H1("Too tall", style: style).ToBytes());
    }

    [Fact]
    public void HeadingLineLargerThanTheFrameFailsWithoutUnboundedPagination() {
        var options = new PdfOptions { PageHeight = 70, MarginTop = 25, MarginBottom = 25 };
        Assert.Throws<ArgumentException>(() => PdfDocument.Create(options).H1("Too tall").ToBytes());
    }
}
