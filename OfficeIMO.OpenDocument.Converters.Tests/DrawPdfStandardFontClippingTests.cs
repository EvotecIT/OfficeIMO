using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Odg.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class DrawPdfStandardFontClippingTests {
    public static IEnumerable<object[]> StandardFontFaces() {
        foreach (string family in new[] { "Helvetica", "Times New Roman", "Courier" })
            foreach (bool bold in new[] { false, true })
                foreach (bool italic in new[] { false, true })
                    yield return new object[] { family, bold, italic };
    }

    [Theory]
    [MemberData(nameof(StandardFontFaces))]
    public void CondensedCaptionsUsePaintedGlyphsForEveryBuiltInFace(string family, bool bold, bool italic) {
        // Absolute line spacing can be smaller than the ink: full captions fit a 20pt
        // frame, while the same caption really loses ink in a 4pt frame.
        Verify(Create("INVISIBLE_MARKER", 20, family, bold, italic), clipped: false, "INVISIBLE_MARKER");
        Verify(Create("INVISIBLE_MARKER", 4, family, bold, italic), clipped: true, "INVISIBLE_MARKER");
    }

    [Fact]
    public void InkFittingBelowOneEmDoesNotAcquireAnUnpaintedFontEnvelope() {
        // Adobe Helvetica metrics put caps and underscore within 11.58pt here.
        // Reader-substituted outlines can differ; this checks declared AFM bounds.
        // A full-font box or a seeded 12pt logical box wrongly rejects this frame.
        Verify(Create("AB_CD", 11.7, "Helvetica"), clipped: false, "AB_CD");
    }

    [Theory]
    [InlineData(0, true)]
    [InlineData(2, false)]
    public void AccentOverhangIsRejectedUntilParagraphMarginMovesItInside(double margin, bool clipped) {
        var source = Create("\u00C1", 20, "Helvetica");
        source.Pages[0].Shapes[0].Paragraphs[0].MarginTop = OdfLength.Points(margin);
        Verify(source, clipped, "\u00C1");
    }

    private static OdgDocument Create(string text, double height, string family, bool bold = false, bool italic = false) {
        var source = OdgDocument.Create();
        var shape = source.AddPage().Shapes.AddTextBox(new OdfRect(OdfLength.Points(20), OdfLength.Points(30),
            OdfLength.Points(180), OdfLength.Points(height)), text);
        var paragraph = shape.Paragraphs[0];
        paragraph.FontFamily = family; paragraph.Bold = bold; paragraph.Italic = italic;
        paragraph.FontSize = OdfLength.Points(12); paragraph.LineHeight = OdfLength.Points(4);
        shape.TextVerticalAlignment = OdfTextAreaVerticalAlignment.Top; shape.WrapText = false;
        shape.TextPadding = new OdfInsets(OdfLength.Points(0), OdfLength.Points(0), OdfLength.Points(0), OdfLength.Points(0));
        return source;
    }

    private static void Verify(OdgDocument source, bool clipped, string text) {
        string[] before = XmlState(source);
        var result = source.ToPdfDocumentResult();
        var report = Assert.IsType<OdfConversionReport>(Assert.Single(result.SourceConversionReports));
        Assert.Equal(clipped, report.Mappings.Any(m => m.Feature.EndsWith(":text-clipped", StringComparison.Ordinal)
            && m.Status == OdfConversionMappingStatus.Unsupported));
        var strict = new OdgToPdfOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported };
        if (clipped) Assert.Throws<OdfConversionLossException>(() => source.ToPdfBytes(strict));
        else Assert.Contains(text, PdfReadDocument.Open(source.ToPdfBytes(strict)).ExtractText());
        Assert.Equal(before, XmlState(source));
    }

    private static string[] XmlState(OdgDocument source) => new[] {
        source.GetXml("content.xml").ToString(), source.GetXml("styles.xml").ToString()
    };
}
