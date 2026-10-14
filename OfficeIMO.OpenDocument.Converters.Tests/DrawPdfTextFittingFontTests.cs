using System;
using System.Linq;
using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Odg.Pdf;
using OfficeIMO.OpenDocument.Testing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class DrawPdfTextFittingFontTests {
    [Theory]
    [InlineData("margin")]
    [InlineData("frame")]
    [InlineData("line")]
    [InlineData("connector")]
    public void PdfStrictPolicyRejectsOmittedParagraphsAndCondensedGlyphPaint(string kind) {
        var source = OdgDocument.Create(); var page = source.AddPage();
        OdgShape shape = kind == "line" ? page.Shapes.AddLine(OdfLength.Points(20), OdfLength.Points(30),
            OdfLength.Points(160), OdfLength.Points(30)) : kind == "connector" ?
            page.Shapes.AddConnector(new OfficePoint(20, 30), new OfficePoint(160, 30)) :
            page.Shapes.AddTextBox(new OdfRect(OdfLength.Points(20), OdfLength.Points(30),
                OdfLength.Points(kind == "margin" ? 40 : 140), OdfLength.Points(kind == "margin" ? 30 : 4)), "END_MARKER");
        if (kind is "line" or "connector") shape.AddParagraph("END_MARKER");
        shape.Paragraphs[0].FontSize = OdfLength.Points(12); shape.TextVerticalAlignment = OdfTextAreaVerticalAlignment.Top;
        shape.WrapText = false; shape.TextPadding = new OdfInsets(OdfLength.Points(0), OdfLength.Points(0), OdfLength.Points(0), OdfLength.Points(0));
        if (kind == "margin") shape.Paragraphs[0].MarginLeft = OdfLength.Points(40);
        else shape.Paragraphs[0].LineHeight = OdfLength.Points(4);
        string[] before = XmlState(source); var result = source.ToPdfDocumentResult();
        Assert.True(IsClipped(Assert.IsType<OdfConversionReport>(Assert.Single(result.SourceConversionReports))));
        Assert.Throws<OdfConversionLossException>(() => source.ToPdfBytes(new OdgToPdfOptions {
            LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported
        }));
        Assert.Equal(before, XmlState(source));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PdfFontSelectionGovernsSourceClippingInBothDirections(bool wide) {
        var source = CreateDocument();
        // The ordinary drawing path proves the opposite font decision; PDF must use its own options.
        var opposite = source.ToDrawings(OdfTextFittingTestFonts.Profile(!wide));
        Assert.Equal(!wide, IsClipped(opposite.Report));
        var pdfOptions = new PdfOptions().RegisterNamedFontFamily(new PdfEmbeddedFontFamily(
            OdfTextFittingTestFonts.Family, OdfTextFittingTestFonts.Create(wide)));
        var options = new OdgToPdfOptions { PdfOptions = pdfOptions };
        string[] before = XmlState(source); var result = source.ToPdfDocumentResult(options);
        var report = Assert.IsType<OdfConversionReport>(Assert.Single(result.SourceConversionReports));
        Assert.Equal(wide, IsClipped(report));
        string rendered = PdfReadDocument.Open(result.ToBytes()).ExtractText();
        Assert.Equal(!wide, rendered.Contains(OdfTextFittingTestFonts.Body));
        options.LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported;
        if (wide) Assert.Throws<OdfConversionLossException>(() => source.ToPdfDocumentResult(options));
        else Assert.Contains(OdfTextFittingTestFonts.Body, PdfReadDocument.Open(source.ToPdfBytes(options)).ExtractText());
        Assert.Equal(before, XmlState(source));
    }

    [Theory]
    [InlineData("fixed")]
    [InlineData("height")]
    [InlineData("width")]
    public void PdfMetricsPreflightObservesEffectiveShapingCancellation(string growth) {
        var source = CreateDocument(); using var cancellation = new CancellationTokenSource();
        source.Pages[0].Shapes[0].AutoGrowHeight = growth == "height";
        source.Pages[0].Shapes[0].AutoGrowWidth = growth == "width";
        var provider = new CancellingProvider(cancellation);
        var pdfOptions = new PdfOptions {
            TextShapingProvider = provider
        }.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(OdfTextFittingTestFonts.Family, OdfTextFittingTestFonts.Create(false)));
        string[] before = XmlState(source);
        Assert.ThrowsAny<OperationCanceledException>(() => source.ToPdfDocumentResult(new OdgToPdfOptions { PdfOptions = pdfOptions }, cancellation.Token));
        Assert.Equal(cancellation.Token, provider.ObservedToken); Assert.Equal(before, XmlState(source));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WidthGrowthCapUsesFinalPdfFontAndStrictPolicyObservesOverflow(bool wide) {
        var source = CreateDocument(); var shape = source.Pages[0].Shapes[0];
        shape.AutoGrowWidth = true; shape.AutoGrowHeight = false; shape.WrapText = false;
        double Width(bool fontWide) => source.Pages[0].ToDrawing(OdfTextFittingTestFonts.Profile(fontWide),
            OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Value.Elements.OfType<OfficeDrawingRichText>().Single().Width;
        double narrowWidth = Width(false), wideWidth = Width(true); Assert.True(wideWidth > narrowWidth);
        shape.TextBoxMinimumWidth = OdfLength.Points(40); shape.TextBoxMaximumWidth = OdfLength.Points((narrowWidth + wideWidth) / 2);
        Assert.Equal(!wide, IsClipped(source.ToDrawings(OdfTextFittingTestFonts.Profile(!wide)).Report));
        var options = new OdgToPdfOptions { PdfOptions = new PdfOptions().RegisterNamedFontFamily(
            new PdfEmbeddedFontFamily(OdfTextFittingTestFonts.Family, OdfTextFittingTestFonts.Create(wide))) };
        string[] before = XmlState(source); var result = source.ToPdfDocumentResult(options);
        Assert.Equal(wide, IsClipped(Assert.IsType<OdfConversionReport>(Assert.Single(result.SourceConversionReports))));
        options.LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported;
        if (wide) Assert.Throws<OdfConversionLossException>(() => source.ToPdfBytes(options));
        else Assert.Equal(OdfTextFittingTestFonts.Body, string.Concat(PdfReadDocument.Open(source.ToPdfBytes(options)).ExtractText().Where(c => !char.IsWhiteSpace(c))));
        Assert.Equal(before, XmlState(source));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HeightGrowthUsesFinalPdfFontSelectionAndRetainsAllGlyphs(bool wide) {
        var source = CreateDocument(); var shape = source.Pages[0].Shapes[0];
        shape.AutoGrowHeight = true; shape.AutoGrowWidth = false;
        shape.FillColor = OdfColor.Parse("#e0f0ff"); shape.StrokeColor = OdfColor.Parse("#204060");
        var options = new OdgToPdfOptions {
            LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported,
            PdfOptions = new PdfOptions().RegisterNamedFontFamily(new PdfEmbeddedFontFamily(
                OdfTextFittingTestFonts.Family, OdfTextFittingTestFonts.Create(wide)))
        };
        string[] before = XmlState(source); var result = source.ToPdfDocumentResult(options);
        var report = Assert.IsType<OdfConversionReport>(Assert.Single(result.SourceConversionReports));
        Assert.False(IsClipped(report));
        Assert.Contains(report.Mappings, m => m.Feature.EndsWith(":text-auto-size", StringComparison.Ordinal) &&
            m.Status == OdfConversionMappingStatus.Approximated);
        string rendered = PdfReadDocument.Open(result.ToBytes()).ExtractText();
        Assert.Equal(OdfTextFittingTestFonts.Body, string.Concat(rendered.Where(c => !char.IsWhiteSpace(c))));
        Assert.Equal(before, XmlState(source));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void GrowthCapUsesFinalPdfFontAndStrictPolicyObservesClipping(bool wide) {
        var source = CreateDocument(); var shape = source.Pages[0].Shapes[0];
        shape.AutoGrowHeight = true; shape.AutoGrowWidth = false;
        shape.TextBoxMinimumHeight = OdfLength.Points(0);
        double Height(bool fontWide) => source.Pages[0].ToDrawing(OdfTextFittingTestFonts.Profile(fontWide),
            OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Value.Elements.OfType<OfficeDrawingRichText>().Single().Height;
        double narrowHeight = Height(false), wideHeight = Height(true);
        Assert.True(wideHeight > narrowHeight);
        shape.TextBoxMaximumHeight = OdfLength.Points((narrowHeight + wideHeight) / 2);
        Assert.Equal(!wide, IsClipped(source.ToDrawings(OdfTextFittingTestFonts.Profile(!wide)).Report));
        var options = new OdgToPdfOptions {
            PdfOptions = new PdfOptions().RegisterNamedFontFamily(new PdfEmbeddedFontFamily(
                OdfTextFittingTestFonts.Family, OdfTextFittingTestFonts.Create(wide)))
        };
        string[] before = XmlState(source);
        var result = source.ToPdfDocumentResult(options);
        Assert.Equal(wide, IsClipped(Assert.IsType<OdfConversionReport>(Assert.Single(result.SourceConversionReports))));
        options.LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported;
        if (wide) Assert.Throws<OdfConversionLossException>(() => source.ToPdfBytes(options));
        else Assert.Equal(OdfTextFittingTestFonts.Body, string.Concat(PdfReadDocument.Open(source.ToPdfBytes(options))
            .ExtractText().Where(c => !char.IsWhiteSpace(c))));
        Assert.Equal(before, XmlState(source));
    }

    private sealed class CancellingProvider(CancellationTokenSource source) : IOfficeTextShapingProvider {
        internal CancellationToken ObservedToken { get; private set; }
        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
            ObservedToken = request.CancellationToken; source.Cancel(); request.CancellationToken.ThrowIfCancellationRequested();
            return null;
        }
    }
    private static bool IsClipped(OdfConversionReport report) => report.Mappings.Any(m =>
        m.Feature.EndsWith(":text-clipped", StringComparison.Ordinal) && m.Status == OdfConversionMappingStatus.Unsupported);
    private static OdgDocument CreateDocument() {
        var document = OdgDocument.Create(); var page = document.AddPage("Font context");
        var shape = page.Shapes.AddTextBox(new OdfRect(OdfLength.Points(20), OdfLength.Points(30),
            OdfLength.Points(40), OdfLength.Points(20)), OdfTextFittingTestFonts.Body, "Body");
        shape.Paragraphs[0].FontFamily = OdfTextFittingTestFonts.Family; shape.Paragraphs[0].FontSize = OdfLength.Points(12);
        shape.WrapText = true; shape.TextPadding = new OdfInsets(OdfLength.Points(0), OdfLength.Points(0), OdfLength.Points(0), OdfLength.Points(0));
        return document;
    }
    private static string[] XmlState(OdfDocument document) => new[] { document.GetXml("content.xml").ToString(), document.GetXml("styles.xml").ToString() };
}
