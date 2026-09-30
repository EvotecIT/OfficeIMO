using System;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfFontTrackingTests {
    [Theory]
    [InlineData(10D, 8D)]
    [InlineData(15D, 13.5D)]
    [InlineData(20D, 20D)]
    public void EmbeddedTrackingChangesRunWidthButPreservesNominalWidths(double size, double expected) {
        var font = PdfTrueTypeFontProgram.Parse(FontData());
        Assert.Equal(500, font.GetGlyphWidth1000(1));
        Assert.Equal(expected, font.MeasureTextWidth("AB", size), 9);
        Assert.Equal(expected, font.ForkForDocument().MeasureTextWidth("AB", size), 9);
    }

    [Fact]
    public void NamedAndStandardFontPaintingRetainTrackingAndText() {
        foreach (bool named in new[] { false, true }) {
            var options = new PdfOptions { CompressContentStreams = false };
            byte[] data = ManagedTextShapingTestAssets.AddTrackingTable(
                ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 'A', 'B'), Table());
            if (named) options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Tracked", data));
            else options.EmbedStandardFont(PdfStandardFont.Helvetica, data);
            byte[] pdf = named
                ? PdfDocument.Create(pdf => pdf.Content(c => c.Paragraph(p => p.FontFamily("Tracked").FontSize(15D).Text("AB"))), options).ToBytes()
                : PdfDocument.Create(pdf => pdf.Content(c => c.Paragraph(p => p.FontSize(15D).Text("AB"))), options).ToBytes();
            Assert.Contains("AB", PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
            Assert.Contains("<0002> 50] TJ", Encoding.ASCII.GetString(pdf), StringComparison.Ordinal);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AttachedGlyphsReceiveOneTrackingAdvanceAtTheirClusterEdge(bool negative) {
        var font = PdfTrueTypeFontProgram.Parse(FontData());
        var run = new PdfGlyphRun(new[] {
            new PdfGlyphInfo(1, "A", 0, 500, negative ? -500 : 500, 20, 0),
            new PdfGlyphInfo(1, "\u0301", 1, 500, 0, -30, 50),
            new PdfGlyphInfo(1, "B", 2, 500, negative ? -500 : 500, 0, 0)
        }, Array.Empty<PdfTextEncodingDiagnostic>(), "A\u0301B",
            negative ? OfficeTextDirection.RightToLeft : OfficeTextDirection.LeftToRight);
        var command = font.ToTextShowCommand("A\u0301B", run);
        Assert.Equal(new[] { false, true, true }, command.TrackingBoundaries);
        var builder = new StringBuilder();
        new ContentStreamBuilder(builder).ShowText(command, 15D);
        string painted = builder.ToString();
        Assert.Contains(negative ? "[<0001> 950] TJ" : "[<0001> 50] TJ", painted);
        // The base's normal advance/offset precedes the mark unchanged, even for a signed RTL run.
        Assert.Contains(negative ? "[-20 <0001> 1020] TJ" : "[-20 <0001> 20] TJ", painted);
        Assert.Contains("0 Ts", painted);
        Assert.Contains("/ActualText", painted);
    }

    [Fact]
    public void TrackingUsesFractionalDesignUnitsUntilThePaintSizeIsKnown() {
        var font = PdfTrueTypeFontProgram.Parse(FontData());
        var builder = new StringBuilder();
        new ContentStreamBuilder(builder).ShowText(font.EncodeTextShowCommand("AB"), 15.125D);
        Assert.Contains("<0001> 48.75] TJ", builder.ToString());
        Assert.Equal(13.6503125D, font.MeasureTextWidth("AB", 15.125D), 9);
    }

    [Fact]
    public void LigatureTrackingUsesThePaintedClusterRatherThanSourceLength() {
        byte[] data = ManagedTextShapingTestAssets.AddTrackingTable(
            ManagedTextShapingTestAssets.CreateFontWithLigature('A', 'B'), Table());
        var font = PdfTrueTypeFontProgram.Parse(data);
        Assert.Equal(6.75D, font.MeasureTextWidth("AB", 15D,
            shapingProvider: OfficeManagedTextShapingProvider.Instance,
            featureSettings: OfficeTextFeatureSettings.Default.With("liga", 1)), 9);
    }

    [Theory]
    [InlineData(false, OfficeTextDirection.LeftToRight)]
    [InlineData(false, OfficeTextDirection.RightToLeft)]
    [InlineData(true, OfficeTextDirection.LeftToRight)]
    [InlineData(true, OfficeTextDirection.RightToLeft)]
    public void ExplicitDirectionPositionedTextMeasuresTheTrackedPaintedWidth(bool named, OfficeTextDirection direction) {
        var options = new PdfOptions();
        if (named) options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Tracked", FontData()));
        else options.EmbedStandardFont(PdfStandardFont.Helvetica, FontData());
        var run = PdfTextRun.Normal("AB", fontSize: 15D, fontFamily: named ? "Tracked" : null)
            .WithTextDirection(direction);
        Assert.Equal(13.5D, PdfWriter.MeasurePositionedText(run, options)!.Value, 9);
    }

    [Theory]
    [InlineData(1D, 23.5D)]
    [InlineData(.75D, 25D)]
    public void IsolatedLogicalClustersRetainTrackingAndTheFollowingTextOrigin(double metricScale, double finalX) {
        var font = PdfTrueTypeFontProgram.Parse(FontData());
        var run = new PdfGlyphRun(new[] {
            new PdfGlyphInfo(1, "A", 0, 500),
            new PdfGlyphInfo(1, "B", 1, 500)
        }, Array.Empty<PdfTextEncodingDiagnostic>(), preserveGlyphUnicode: true);
        var content = new StringBuilder();
        new ContentStreamBuilder(content).TextMatrix(10D, 20D)
            .ShowText(font.ToTextShowCommand("AB", run, metricScale), 15D);

        string painted = content.ToString();
        Assert.Contains("/ActualText <41>", painted, StringComparison.Ordinal);
        Assert.Contains("/ActualText <42>", painted, StringComparison.Ordinal);
        Assert.Contains("1 0 0 1 " + finalX.ToString(System.Globalization.CultureInfo.InvariantCulture)
            + " 20 Tm", painted, StringComparison.Ordinal);
    }

    private static byte[] FontData() => ManagedTextShapingTestAssets.CreateTrackingFont(
        Table());

    internal static byte[] Table() => ManagedTextShapingTestAssets.CreateTrackingTable(
        new[] { 10D, 20D }, new[] { -1D, 1D },
        new[] { new short[] { -120, -40 }, new short[] { -80, 40 } });
}
