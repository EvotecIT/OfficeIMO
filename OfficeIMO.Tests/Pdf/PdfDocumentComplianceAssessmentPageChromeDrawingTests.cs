using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentComplianceAssessmentTests {

    [Fact]
    public void HeaderFooterImageAlternativeTextEmitsMarkedFigureContent() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .Header(header => header.Image(CreateMinimalRgbPng(), 24, 12, alternativeText: "Header logo"))
            .Footer(footer => footer.Image(CreateMinimalRgbPng(), 24, 12, alternativeText: "Footer logo"))
            .Paragraph(paragraph => paragraph.Text("Header and footer image marked content."))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/Figure << /Alt <486561646572206C6F676F> >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Figure << /Alt <466F6F746572206C6F676F> >> BDC", content, StringComparison.Ordinal);
    }

    [Fact]
    public void DecorativeBackgroundAndWatermarkImagesEmitArtifactMarkedContent() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .BackgroundImage(CreateMinimalRgbPng(), OfficeIMO.Drawing.OfficeImageFit.Stretch, opacity: 0.2)
            .ImageWatermark(CreateMinimalRgbPng(), 80, 40, opacity: 0.2)
            .Paragraph(paragraph => paragraph.Text("Decorative image artifact marked content."))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.True(CountOccurrences(content, "/Artifact BMC") >= 2);
        Assert.DoesNotContain("/Figure << /Alt", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedHeaderFooterImagesWithoutAlternativeTextEmitArtifactMarkedContent() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Header(header => header.Image(CreateMinimalRgbPng(), 24, 12))
            .Footer(footer => footer.Image(CreateMinimalRgbPng(), 24, 12))
            .Paragraph(paragraph => paragraph.Text("Header and footer decorative images."))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.True(CountOccurrences(content, "/Artifact BMC") >= 2);
        Assert.DoesNotContain("/Figure << /Alt", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedDecorativePageChromeEmitsArtifactMarkedContent() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Background(PdfColor.White)
            .BackgroundImage(CreateMinimalRgbPng(), OfficeIMO.Drawing.OfficeImageFit.Stretch, opacity: 0.2)
            .ImageWatermark(CreateMinimalRgbPng(), 80, 40, opacity: 0.2)
            .Watermark("DRAFT", fontSize: 48, opacity: 0.18)
            .PageBorder(inset: 30, opacity: 0.4)
            .Paragraph(paragraph => paragraph.Text("Decorative page chrome artifact marked content."))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.True(CountOccurrences(content, "/Artifact BMC") >= 5);
        Assert.Contains("/Type /StructElem /S /Document", content, StringComparison.Ordinal);
        Assert.DoesNotContain("/Figure << /Alt", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedRunningHeaderFooterTextEmitsArtifactMarkedContent() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Header(header => header.Text("Running header"))
            .Footer(footer => footer.Text("Page {page} of {pages}"))
            .Paragraph(paragraph => paragraph.Text("Body text remains structured."))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.True(CountOccurrences(content, "/Artifact BMC") >= 2);
        Assert.Contains("/P << /MCID", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Document", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedHorizontalRulesEmitArtifactMarkedContent() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Paragraph(paragraph => paragraph.Text("Before top-level rule."))
            .HR(thickness: 1.2, color: PdfColor.FromRgb(148, 163, 184), spacingBefore: 2, spacingAfter: 2)
            .Paragraph(paragraph => paragraph.Text("After top-level rule."))
            .Compose(document => document.Page(page => page.Content(content => content.Row(row => row
                .Column(100, column => column
                    .Paragraph(paragraph => paragraph.Text("Before row rule."))
                    .HR(thickness: 0.8, color: PdfColor.FromRgb(203, 213, 225), spacingBefore: 2, spacingAfter: 2)
                    .Paragraph(paragraph => paragraph.Text("After row rule.")))))))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.True(CountOccurrences(content, "/Artifact BMC") >= 2);
        Assert.True(CountOccurrences(content, "/P << /MCID") >= 4);
        Assert.Contains("/Type /StructElem /S /Document", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedDecorativeLayoutChromeEmitsArtifactMarkedContent() {
        var panelStyle = new PanelStyle {
            Background = PdfColor.FromRgb(248, 250, 252),
            BorderColor = PdfColor.FromRgb(37, 99, 235),
            BorderWidth = 0.8,
            PaddingX = 10,
            PaddingY = 8
        };
        PdfTableStyle tableStyle = TableStyles.Minimal();
        tableStyle.HeaderRowCount = 1;
        tableStyle.HeaderFill = PdfColor.FromRgb(229, 231, 235);
        tableStyle.RowStripeFill = PdfColor.FromRgb(248, 250, 252);
        tableStyle.BorderColor = PdfColor.FromRgb(148, 163, 184);
        tableStyle.BorderWidth = 0.7;
        tableStyle.RowSeparatorColor = PdfColor.FromRgb(203, 213, 225);
        tableStyle.RowSeparatorWidth = 0.5;
        tableStyle.CellFills = new Dictionary<(int Row, int Column), PdfColor> {
            [(1, 1)] = PdfColor.FromRgb(220, 252, 231)
        };
        tableStyle.CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
            [(1, 1)] = new PdfCellBorder {
                Color = PdfColor.FromRgb(22, 163, 74),
                Width = 0.8
            }
        };

        byte[] pdf = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Panel(panel => panel.Paragraph(paragraph => paragraph.Text("Decorative panel chrome.")), panelStyle)
            .Table(new[] {
                new[] { "Name", "Status" },
                new[] { "Alpha", "Ready" },
                new[] { "Beta", "Done" }
            }, style: tableStyle)
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.True(CountOccurrences(content, "/Artifact BMC") >= 8);
        Assert.Contains("/P << /MCID", content, StringComparison.Ordinal);
        Assert.Contains("/TH << /MCID", content, StringComparison.Ordinal);
        Assert.Contains("/TD << /MCID", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Table", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /TD", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedShapeAndDrawingAlternativeTextEmitFigureStructureReferences() {
        OfficeShape shape = CreateComplianceShape();
        OfficeDrawing drawing = new OfficeDrawing(36, 18)
            .AddShape(CreateComplianceShape(36, 18), 0, 0);

        byte[] pdf = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Shape(shape, style: new PdfDrawingStyle {
                AlternativeText = "Risk status badge"
            }, linkUri: "https://officeimo.net/shape", linkContents: "Risk shape")
            .Drawing(drawing, style: new PdfDrawingStyle {
                AlternativeText = "Approval workflow diagram"
            }, linkUri: "https://officeimo.net/drawing", linkContents: "Approval drawing")
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/MarkInfo << /Marked true >>", content, StringComparison.Ordinal);
        Assert.Contains("/StructParents 0", content, StringComparison.Ordinal);
        Assert.Equal(2, CountOccurrences(content, "/Subtype /Link"));
        Assert.Contains("/Figure << /Alt <5269736B20737461747573206261646765> /MCID 0 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Figure << /Alt <417070726F76616C20776F726B666C6F77206469616772616D> /MCID 1 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Figure", content, StringComparison.Ordinal);
        Assert.Matches(@"/Type /StructElem /S /Figure /P \d+ 0 R /Pg \d+ 0 R /K \[<< /Type /MCR /Pg \d+ 0 R /MCID 0 >> << /Type /OBJR /Obj \d+ 0 R >>\] /Alt", content);
        Assert.Matches(@"/Type /StructElem /S /Figure /P \d+ 0 R /Pg \d+ 0 R /K \[<< /Type /MCR /Pg \d+ 0 R /MCID 1 >> << /Type /OBJR /Obj \d+ 0 R >>\] /Alt", content);
        Assert.Contains("/ParentTree", content, StringComparison.Ordinal);
        Assert.Matches(@"/Nums \[0 \[(?<shape>\d+) 0 R (?<drawing>\d+) 0 R\] 1 \k<shape> 0 R 2 \k<drawing> 0 R\]", content);
    }

    [Fact]
    public void TaggedDecorativeShapeAndDrawingEmitArtifactMarkedContent() {
        OfficeDrawing drawing = new OfficeDrawing(36, 18)
            .AddShape(CreateComplianceShape(36, 18), 0, 0);

        byte[] pdf = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Shape(CreateComplianceShape(), style: new PdfDrawingStyle {
                Decorative = true
            })
            .Drawing(drawing, style: new PdfDrawingStyle {
                Decorative = true
            })
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.True(CountOccurrences(content, "/Artifact BMC") >= 2);
        Assert.DoesNotContain("/Figure << /Alt", content, StringComparison.Ordinal);
        Assert.Contains("/StructTreeRoot", content, StringComparison.Ordinal);
    }

}
