using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageExtractorTests {
    [Fact]
    public void ExtractPages_PreservesImageStreamsForSelectedPages() {
        byte[] source = PdfDocument.Create()
            .Paragraph(p => p.Text("Cover page"))
            .PageBreak()
            .Image(CreateMinimalRgbPng(), 24, 24)
            .Paragraph(p => p.Text("Image page marker"))
            .ToBytes();

        byte[] extracted = PdfPageExtractor.ExtractPages(source, 2);

        using var pdf = PdfPigDocument.Open(new MemoryStream(extracted));
        Assert.Equal(1, pdf.NumberOfPages);

        string pdfText = Encoding.ASCII.GetString(extracted);
        Assert.Contains("/Subtype /Image", pdfText);
        Assert.Contains("/Filter /FlateDecode", pdfText);
        Assert.Contains("/Width 1", pdfText);
        Assert.Contains("/Height 1", pdfText);
        string extractedText = NormalizeExtractedText(PdfReadDocument.Open(extracted).ExtractText());
        Assert.Contains("Imagepagemarker", extractedText);
        Assert.DoesNotContain("Coverpage", extractedText);
    }

    [Fact]
    public void ExtractPages_PreservesLinkAnnotationsForSelectedPages() {
        byte[] source = PdfDocument.Create()
            .Paragraph(p => p.Text("Cover page"))
            .PageBreak()
            .Paragraph(p => p.Link("OfficeIMO link", "https://evotec.xyz"))
            .ToBytes();

        byte[] extracted = PdfPageExtractor.ExtractPages(source, 2);

        PdfDocumentInfo info = PdfInspector.Inspect(extracted);
        Assert.True(info.HasAnnotations);

        string pdfText = Encoding.ASCII.GetString(extracted);
        Assert.Contains("/Annots [", pdfText, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Link", pdfText, StringComparison.Ordinal);
        Assert.Contains("/URI (https://evotec.xyz)", pdfText, StringComparison.Ordinal);

        string extractedText = NormalizeExtractedText(PdfReadDocument.Open(extracted).ExtractText());
        Assert.Contains("OfficeIMOlink", extractedText);
        Assert.DoesNotContain("Coverpage", extractedText);
    }

    [Fact]
    public void ExtractPages_ClonesAnnotationsForDuplicatePageSelections() {
        byte[] source = PdfDocument.Create()
            .Paragraph(p => p.Text("Cover page"))
            .PageBreak()
            .Paragraph(p => p.Link("OfficeIMO link", "https://evotec.xyz"))
            .ToBytes();

        byte[] extracted = PdfPageExtractor.ExtractPages(source, 2, 2);

        var document = PdfReadDocument.Open(extracted);
        var (objects, _) = PdfSyntax.ParseObjects(extracted);
        var annotationObjectNumbers = new HashSet<int>();

        for (int i = 0; i < document.Pages.Count; i++) {
            int pageObjectNumber = document.Pages[i].ObjectNumber;
            var page = Assert.IsType<PdfDictionary>(objects[pageObjectNumber].Value);
            var annotations = Assert.IsType<PdfArray>(page.Items["Annots"]);
            Assert.NotEmpty(annotations.Items);

            foreach (var annotationObject in annotations.Items) {
                var annotationReference = Assert.IsType<PdfReference>(annotationObject);
                annotationObjectNumbers.Add(annotationReference.ObjectNumber);

                var annotation = Assert.IsType<PdfDictionary>(objects[annotationReference.ObjectNumber].Value);
                if (annotation.Items.TryGetValue("P", out var annotationPage)) {
                    var pageReference = Assert.IsType<PdfReference>(annotationPage);
                    Assert.Equal(pageObjectNumber, pageReference.ObjectNumber);
                }
            }
        }

        Assert.Equal(2, annotationObjectNumbers.Count);
    }

    [Fact]
    public void ExtractPages_NormalizesRemappedReferencesToGenerationZero() {
        byte[] source = BuildSinglePagePdfWithGenerationOneContent();

        byte[] extracted = PdfPageExtractor.ExtractPages(source, 1);

        string extractedText = NormalizeExtractedText(PdfReadDocument.Open(extracted).ExtractText());
        Assert.Contains("Generationonecontent", extractedText);

        var document = PdfReadDocument.Open(extracted);
        var (objects, _) = PdfSyntax.ParseObjects(extracted);
        int pageObjectNumber = Assert.Single(document.Pages).ObjectNumber;
        var page = Assert.IsType<PdfDictionary>(objects[pageObjectNumber].Value);
        var contents = Assert.IsType<PdfReference>(page.Items["Contents"]);
        Assert.Equal(0, contents.Generation);
        Assert.True(objects.ContainsKey(contents.ObjectNumber));
    }

    [Fact]
    public void ExtractPages_NormalizesClonedAnnotationReferencesToGenerationZero() {
        byte[] source = BuildSinglePagePdfWithGenerationOneContent(includeAnnotation: true);

        byte[] extracted = PdfPageExtractor.ExtractPages(source, 1, 1);

        var document = PdfReadDocument.Open(extracted);
        var (objects, _) = PdfSyntax.ParseObjects(extracted);
        Assert.Equal(2, document.Pages.Count);

        foreach (var readPage in document.Pages) {
            var page = Assert.IsType<PdfDictionary>(objects[readPage.ObjectNumber].Value);
            var annotations = Assert.IsType<PdfArray>(page.Items["Annots"]);
            var annotationReference = Assert.IsType<PdfReference>(Assert.Single(annotations.Items));

            Assert.Equal(0, annotationReference.Generation);
            Assert.True(objects.ContainsKey(annotationReference.ObjectNumber));
        }
    }

    [Fact]
    public void ExtractPages_RejectsWrongGenerationReferencesBeforeRewrite() {
        byte[] source = BuildSinglePagePdfWithGenerationOneContent(contentObjectGeneration: 0, contentReferenceGeneration: 1);

        PdfMutationBlockedException exception = Assert.Throws<PdfMutationBlockedException>(() => PdfPageExtractor.ExtractPages(source, 1));

        Assert.Contains("FullRewrite.InvalidObjectReferences", exception.Plan.BlockerCodes);
        Assert.Contains(exception.Plan.Diagnostics, diagnostic => diagnostic.Contains("PDF object 4 1 R", StringComparison.Ordinal));
        Assert.Contains(exception.Plan.Diagnostics, diagnostic => diagnostic.Contains("active object generation is 0", StringComparison.Ordinal));
    }

    [Fact]
    public void ExtractPages_DropsBookmarkLinksWhenDestinationPageIsNotCopied() {
        byte[] source = PdfDocument.Create()
            .Paragraph(p => p.LinkToBookmark("Jump to details", "Details"))
            .PageBreak()
            .Bookmark("Details")
            .Paragraph(p => p.Text("Details marker"))
            .ToBytes();

        byte[] extracted = PdfPageExtractor.ExtractPages(source, 1);

        PdfDocumentInfo info = PdfInspector.Inspect(extracted);
        Assert.Single(info.Pages);
        Assert.Empty(info.NamedDestinationNames);
        Assert.Empty(info.LinkAnnotations);
        Assert.Empty(info.Pages[0].LinkAnnotations);

        string pdfText = Encoding.ASCII.GetString(extracted);
        Assert.DoesNotContain("/S /GoTo", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain("(Details)", pdfText, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractPages_PreservesBookmarkLinksWhenDestinationPageIsCopied() {
        byte[] source = PdfDocument.Create()
            .Paragraph(p => p.LinkToBookmark("Jump to details", "Details"))
            .PageBreak()
            .Bookmark("Details")
            .Paragraph(p => p.Text("Details marker"))
            .ToBytes();

        byte[] extracted = PdfPageExtractor.ExtractPages(source, 1, 2);

        PdfDocumentInfo info = PdfInspector.Inspect(extracted);
        Assert.Equal(2, info.PageCount);
        Assert.Equal(new[] { "Details" }, info.NamedDestinationNames);

        Assert.NotEmpty(info.LinkAnnotations);
        Assert.All(info.LinkAnnotations, link => {
            Assert.True(link.IsNamedDestinationLink);
            Assert.Equal("Details", link.DestinationName);
            Assert.Equal(1, link.PageNumber);
        });
    }

    [Fact]
    public void SplitPages_DropsBookmarkLinksWhoseDestinationsMoveToAnotherDocument() {
        byte[] source = PdfDocument.Create()
            .Paragraph(p => p.LinkToBookmark("Jump to details", "Details"))
            .PageBreak()
            .Bookmark("Details")
            .Paragraph(p => p.Text("Details marker"))
            .ToBytes();

        IReadOnlyList<byte[]> splitPages = PdfPageExtractor.SplitPages(source);

        Assert.Equal(2, splitPages.Count);
        PdfDocumentInfo first = PdfInspector.Inspect(splitPages[0]);
        PdfDocumentInfo second = PdfInspector.Inspect(splitPages[1]);

        Assert.Empty(first.LinkAnnotations);
        Assert.Empty(first.NamedDestinationNames);
        Assert.Empty(second.LinkAnnotations);
        Assert.Equal(new[] { "Details" }, second.NamedDestinationNames);
    }

    [Fact]
    public void ExtractPages_DropsDirectGoToLinksWhenDestinationPageIsNotCopied() {
        byte[] source = BuildTwoPageDirectGoToLinkAnnotationPdf();

        byte[] extracted = PdfPageExtractor.ExtractPages(source, 1);

        PdfDocumentInfo info = PdfInspector.Inspect(extracted);
        Assert.Single(info.Pages);
        Assert.Empty(info.LinkAnnotations);
        Assert.Empty(info.Pages[0].LinkAnnotations);

        string pdfText = Encoding.ASCII.GetString(extracted);
        Assert.DoesNotContain("/S /GoTo", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain("/FitH", pdfText, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractPages_DropsDirectGoToLinksFromDuplicatePagesWhenDestinationPageIsNotCopied() {
        byte[] source = BuildTwoPageDirectGoToLinkAnnotationPdf();

        byte[] extracted = PdfPageExtractor.ExtractPages(source, 1, 1);

        PdfDocumentInfo info = PdfInspector.Inspect(extracted);
        Assert.Equal(2, info.PageCount);
        Assert.Empty(info.LinkAnnotations);
        Assert.All(info.Pages, page => Assert.Empty(page.LinkAnnotations));

        string pdfText = Encoding.ASCII.GetString(extracted);
        Assert.DoesNotContain("/S /GoTo", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain("/FitH", pdfText, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractPages_PreservesDirectGoToLinksWhenDestinationPageIsCopied() {
        byte[] source = BuildTwoPageDirectGoToLinkAnnotationPdf();

        byte[] extracted = PdfPageExtractor.ExtractPages(source, 1, 2);

        PdfDocumentInfo info = PdfInspector.Inspect(extracted);
        Assert.Equal(2, info.PageCount);
        PdfLinkAnnotation link = Assert.Single(info.LinkAnnotations);
        Assert.Equal(1, link.PageNumber);
        Assert.Equal(2, link.DestinationPageNumber);
        Assert.Equal(144d, link.DestinationTop);
        Assert.Equal(PdfOpenActionDestinationMode.FitHorizontal, link.DestinationMode);
        Assert.True(link.IsInternalDestinationLink);
        Assert.False(link.IsNamedDestinationLink);
        Assert.False(link.IsUriLink);
    }

    [Fact]
    public void ExtractPages_PreservesIndirectDirectGoToLinksWhenDestinationPageIsCopied() {
        byte[] source = BuildTwoPageDirectGoToLinkAnnotationPdf(indirectDestination: true);

        byte[] extracted = PdfPageExtractor.ExtractPages(source, 1, 2);

        PdfDocumentInfo info = PdfInspector.Inspect(extracted);
        Assert.Equal(2, info.PageCount);
        PdfLinkAnnotation link = Assert.Single(info.LinkAnnotations);
        Assert.Equal(1, link.PageNumber);
        Assert.Equal(2, link.DestinationPageNumber);
        Assert.Equal(144d, link.DestinationTop);
        Assert.Equal(PdfOpenActionDestinationMode.FitHorizontal, link.DestinationMode);
    }

    private static byte[] BuildTwoPageDirectGoToLinkAnnotationPdf(bool indirectDestination = false) {
        string annotationDestination = indirectDestination ? "8 0 R" : "[5 0 R /FitH 144]";
        var lines = new List<string> {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 5 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /Annots [7 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 6 0 R >>",
            "endobj",
            "6 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /Type /Annot /Subtype /Link /Rect [10 20 90 42] /Contents (Jump to page two) /A << /S /GoTo /D " + annotationDestination + " >> >>",
            "endobj"
        };

        if (indirectDestination) {
            lines.AddRange(new[] {
                "8 0 obj",
                "[5 0 R /FitH 144]",
                "endobj"
            });
        }

        lines.Add("trailer");
        lines.Add(indirectDestination ? "<< /Root 1 0 R /Size 9 >>" : "<< /Root 1 0 R /Size 8 >>");
        lines.Add("%%EOF");

        string pdf = string.Join("\n", lines);
        return Encoding.ASCII.GetBytes(pdf);
    }

}
