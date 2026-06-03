using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentComplianceAssessmentTests {

    [Fact]
    public void ImageAlternativeTextEmitsMarkedFigureContent() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .Image(CreateMinimalRgbPng(), 24, 24, alternativeText: "Company logo")
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/Figure << /Alt <436F6D70616E79206C6F676F> >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("EMC", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedImageAlternativeTextEmitsFigureStructureReferences() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Image(CreateMinimalRgbPng(), 24, 24, alternativeText: "Company logo")
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/MarkInfo << /Marked true >>", content, StringComparison.Ordinal);
        Assert.Contains("/StructTreeRoot", content, StringComparison.Ordinal);
        Assert.Contains("/StructParents 0", content, StringComparison.Ordinal);
        Assert.Contains("/Figure << /Alt <436F6D70616E79206C6F676F> /MCID 0 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/ParentTree", content, StringComparison.Ordinal);
        Assert.Contains("/ParentTreeNextKey 1", content, StringComparison.Ordinal);
        Assert.Contains("/Nums [0 [", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Document", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Figure", content, StringComparison.Ordinal);
        Assert.Contains("/K << /Type /MCR", content, StringComparison.Ordinal);
        Assert.Contains("/MCID 0", content, StringComparison.Ordinal);
        Assert.Contains("/Alt <436F6D70616E79206C6F676F>", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedDocumentStructureRootEmitsLanguageMetadata() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false,
                Language = "en-US"
            })
            .TaggedPdfCatalogMarkers()
            .Paragraph(paragraph => paragraph.Text("Language metadata for generated document structure."))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/Lang <656E2D5553>", content, StringComparison.Ordinal);
        Assert.Matches(@"/Type /StructElem /S /Document /P \d+ 0 R /K \[[^\]]+\] /Lang <656E2D5553>", content);
    }

    [Fact]
    public void TaggedHeadingParagraphAndImageEmitStructureReferencesWithPageScopedMcids() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .H1("Quarterly summary")
            .Paragraph(paragraph => paragraph.Text("Revenue and risk notes."))
            .Image(CreateMinimalRgbPng(), 24, 24, alternativeText: "Company logo")
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/StructParents 0", content, StringComparison.Ordinal);
        Assert.Contains("/H1 << /MCID 0 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/P << /MCID 1 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Figure << /Alt <436F6D70616E79206C6F676F> /MCID 2 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Document", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /H1", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /P", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Figure", content, StringComparison.Ordinal);
        Assert.Contains("/ParentTree", content, StringComparison.Ordinal);
        Assert.Contains("/ParentTreeNextKey 1", content, StringComparison.Ordinal);
        Assert.Contains("/Nums [0 [", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedImageStructureFollowsFlowOrderBetweenParagraphs() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Paragraph(paragraph => paragraph.Text("Before image."))
            .Image(CreateMinimalRgbPng(), 24, 24, alternativeText: "Inline diagram")
            .Paragraph(paragraph => paragraph.Text("After image."))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/P << /MCID 0 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Figure << /Alt <496E6C696E65206469616772616D> /MCID 1 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/P << /MCID 2 >> BDC", content, StringComparison.Ordinal);
        Assert.Matches(@"/Nums \[0 \[\d+ 0 R \d+ 0 R \d+ 0 R\]\]", content);
    }

    [Fact]
    public void TaggedParagraphLinkEmitsStructureTabOrder() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Paragraph(paragraph => paragraph
                .Text("Read the ")
                .Link("project site", "https://officeimo.net/", contents: "Project site"))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/Annots [", content, StringComparison.Ordinal);
        Assert.Contains("/StructParents 0 /Tabs /S", content, StringComparison.Ordinal);
        Assert.Contains("/P << /MCID 0 >> BDC", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedParagraphLinkEmitsAnnotationStructureReferences() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Paragraph(paragraph => paragraph
                .Text("Read the ")
                .Link("project site", "https://officeimo.net/", contents: "Project site"))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/Subtype /Link", content, StringComparison.Ordinal);
        Assert.Contains("/StructParent 1", content, StringComparison.Ordinal);
        Assert.Contains("ET\nEMC\n/Link << /MCID", content, StringComparison.Ordinal);
        Assert.Contains("/Link << /MCID 1 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/P << /MCID 2 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Link << /MCID 3 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Document", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Link", content, StringComparison.Ordinal);
        Assert.Contains("/K [<< /Type /MCR /Pg ", content, StringComparison.Ordinal);
        Assert.Contains("/MCID 1 >> << /Type /MCR /Pg ", content, StringComparison.Ordinal);
        Assert.Contains("/MCID 3 >> << /Type /OBJR /Obj ", content, StringComparison.Ordinal);
        Assert.Contains("/ParentTreeNextKey 2", content, StringComparison.Ordinal);
        Assert.Contains("/Nums [0 [", content, StringComparison.Ordinal);
        Assert.Matches(@"/Nums \[0 \[[^\]]+\] 1 \d+ 0 R\]", content);
        Assert.Matches(@"/Nums \[0 \[\d+ 0 R (?<link>\d+) 0 R \d+ 0 R \k<link> 0 R\] 1 \k<link> 0 R\]", content);
    }

    [Fact]
    public void TaggedLinkedHeadingEmitsLinkStructureForVisibleText() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .H1("Embedded PDF fonts", linkUri: "https://officeimo.net/", linkContents: "OfficeIMO PDF")
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/Subtype /Link", content, StringComparison.Ordinal);
        Assert.Contains("/Link << /MCID", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /H1", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Link", content, StringComparison.Ordinal);
        Assert.Matches(@"/Type /StructElem /S /H1 /P \d+ 0 R /Pg \d+ 0 R /K \[\d+ 0 R\]", content);
        Assert.Matches(@"/Type /StructElem /S /Link /P \d+ 0 R /Pg \d+ 0 R /K \[<< /Type /MCR /Pg \d+ 0 R /MCID 0 >> << /Type /OBJR /Obj \d+ 0 R >>\]", content);
        Assert.Contains("/Type /OBJR /Obj", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedWrappedLinkedHeadingPreservesAllAnnotationStructureReferences() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false,
                PageWidth = 170,
                MarginLeft = 24,
                MarginRight = 24
            })
            .TaggedPdfCatalogMarkers()
            .H1("Embedded PDF fonts and compliance evidence", linkUri: "https://officeimo.net/", linkContents: "OfficeIMO PDF")
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.True(CountOccurrences(content, "/Subtype /Link") >= 2);
        Assert.Contains("/Type /StructElem /S /H1", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Link", content, StringComparison.Ordinal);
        Assert.Matches(@"/Type /StructElem /S /Link /P \d+ 0 R /Pg \d+ 0 R /K \[<< /Type /MCR /Pg \d+ 0 R /MCID 0 >> << /Type /OBJR /Obj \d+ 0 R >> << /Type /OBJR /Obj \d+ 0 R >>", content);
        Assert.Matches(@"/Nums \[0 \[(?<link>\d+) 0 R\] 1 \k<link> 0 R 2 \k<link> 0 R", content);
    }

    [Fact]
    public void TaggedRichTextDecorationsEmitArtifactMarkedContent() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Paragraph(paragraph => paragraph
                .Underlined("Underlined")
                .Text(" ")
                .Strikethrough("Strike"))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/P << /MCID 0 >> BDC", content, StringComparison.Ordinal);
        Assert.True(CountOccurrences(content, "/Artifact BMC") >= 2);
    }

    [Fact]
    public void TaggedRichTextBackgroundFillsEmitArtifactMarkedContent() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Paragraph(paragraph => paragraph
                .BackgroundColor(PdfColor.FromRgb(255, 255, 0))
                .Text("Highlighted text"))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/P << /MCID 0 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Artifact BMC", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedLinkedBaselineShiftResetsTextRiseBeforeNormalText() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Paragraph(paragraph => paragraph
                .Text("Value ")
                .Superscript()
                .Link("2", "https://officeimo.net/sup", underline: false)
                .Superscript(false)
                .Text(" normal"))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("0 Ts", content, StringComparison.Ordinal);
        Assert.Matches(@"/Link << /MCID \d+ >> BDC[\s\S]+[1-9]\d*(?:\.\d+)? Ts[\s\S]+EMC[\s\S]+0 Ts", content);
    }

    [Fact]
    public void UntaggedParagraphLinksDoNotEmitDanglingMarkedContentReferences() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .Paragraph(paragraph => paragraph
                .Text("Read the ")
                .Link("project site", "https://officeimo.net/", contents: "Project site"))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/Subtype /Link", content, StringComparison.Ordinal);
        Assert.DoesNotContain("/Link << /MCID", content, StringComparison.Ordinal);
        Assert.DoesNotContain("/StructParent", content, StringComparison.Ordinal);
        Assert.DoesNotContain("/StructTreeRoot", content, StringComparison.Ordinal);
    }

}
