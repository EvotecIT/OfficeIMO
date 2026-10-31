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
        var options = new PdfCreateOptions { PageWidth = global::OfficeIMO.OfficeLength.Points(220), PageHeight = global::OfficeIMO.OfficeLength.Points(160), MarginLeft = global::OfficeIMO.OfficeLength.Points(25), MarginRight = global::OfficeIMO.OfficeLength.Points(25), MarginTop = global::OfficeIMO.OfficeLength.Points(25), MarginBottom = global::OfficeIMO.OfficeLength.Points(25) };
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
        PdfStructureElementInfo heading = Assert.Single(tagged.StructureElements, element => element.StructureType == "H" + level);
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
        var options = new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(pageWidth), PageHeight = global::OfficeIMO.OfficeLength.Points(pageHeight),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(25), MarginRight = global::OfficeIMO.OfficeLength.Points(25), MarginTop = global::OfficeIMO.OfficeLength.Points(25), MarginBottom = global::OfficeIMO.OfficeLength.Points(25),
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
                Assert.InRange(letter.BoundingBox.Bottom, options.MarginBottom.ToPoints() - 1, options.PageHeight.ToPoints() - options.MarginTop.ToPoints() + 1);
                Assert.InRange(letter.BoundingBox.Top, options.MarginBottom.ToPoints() - 1, options.PageHeight.ToPoints() - options.MarginTop.ToPoints() + 1);
            });
        }
        var rectangles = ExtractLinkRectangles(System.Text.Encoding.ASCII.GetString(bytes));
        Assert.True(rectangles.Count > pdf.NumberOfPages);
        string content = System.Text.Encoding.ASCII.GetString(bytes);
        Assert.Equal(pdf.NumberOfPages, CountOccurrences(content, "/StructParents "));
        Assert.Equal(rectangles.Count, CountOccurrences(content, "/Type /OBJR /Obj "));
        Assert.All(rectangles, rectangle => {
            Assert.InRange(rectangle.Y1, options.MarginBottom.ToPoints() - 1, options.PageHeight.ToPoints() - options.MarginTop.ToPoints() + 1);
            Assert.InRange(rectangle.Y2, options.MarginBottom.ToPoints() - 1, options.PageHeight.ToPoints() - options.MarginTop.ToPoints() + 1);
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
        // The boundary case leaves 33.5995 points for an explicitly 33.6-point line.
        var options = new PdfCreateOptions {
            PageHeight = global::OfficeIMO.OfficeLength.Points(pageHeight), MarginTop = global::OfficeIMO.OfficeLength.Points(25), MarginBottom = global::OfficeIMO.OfficeLength.Points(25),
            DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = global::OfficeIMO.OfficeLength.Points(24),
            DefaultParagraphStyle = new PdfParagraphStyle { LineHeight = 1.4, SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) }
        };
        Assert.Throws<ArgumentException>(() => PdfDocument.Create(options).Paragraph(p => p.Text("Too tall")).ToBytes());
    }

    [Fact]
    public void SmallPositiveFrameAcceptsALineThatFits() {
        var options = new PdfCreateOptions { PageHeight = global::OfficeIMO.OfficeLength.Points(50.2), MarginTop = global::OfficeIMO.OfficeLength.Points(25), MarginBottom = global::OfficeIMO.OfficeLength.Points(25), DefaultFontSize = global::OfficeIMO.OfficeLength.Points(0.1) };
        byte[] bytes = PdfDocument.Create(options).Paragraph(p => p.Text("Fits")).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.Equal("Fits", pdf.GetPage(1).Text);
    }

    [Fact]
    public void HeadingRequiredTopSpacingCannotBeSilentlyDroppedToFitALine() {
        var options = new PdfCreateOptions { PageHeight = global::OfficeIMO.OfficeLength.Points(150), MarginTop = global::OfficeIMO.OfficeLength.Points(25), MarginBottom = global::OfficeIMO.OfficeLength.Points(25) };
        var style = new PdfHeadingStyle { FontSize = global::OfficeIMO.OfficeLength.Points(24), LineHeight = 1.25, SpacingBefore = global::OfficeIMO.OfficeLength.Points(90), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), ApplySpacingBeforeAtTop = true };
        Assert.Throws<ArgumentException>(() => PdfDocument.Create(options).H1("Too tall", style: style).ToBytes());
    }

    [Fact]
    public void HeadingLineLargerThanTheFrameFailsWithoutUnboundedPagination() {
        var options = new PdfCreateOptions { PageHeight = global::OfficeIMO.OfficeLength.Points(70), MarginTop = global::OfficeIMO.OfficeLength.Points(25), MarginBottom = global::OfficeIMO.OfficeLength.Points(25) };
        Assert.Throws<ArgumentException>(() => PdfDocument.Create(options).H1("Too tall").ToBytes());
    }
}
