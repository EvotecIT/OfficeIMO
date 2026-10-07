using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfLogicalTextCommandTests {
    [Theory]
    [InlineData(10, 100)]
    [InlineData(15, 50)]
    [InlineData(20, 0)]
    public void IsolatedLogicalGlyphsPreserveTrackingAdvanceAtConvertedFontSize(int cssSize, int adjustment) {
        byte[] data = ManagedTextShapingTestAssets.AddTrackingTable(
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A', 'B'),
            ManagedTextShapingTestAssets.CreateTrackingTable(new[] { 10D, 20D }, new[] { 0D },
                new[] { new short[] { -100, 0 } }));
        var font = PdfTrueTypeFontProgram.Parse(data);
        var run = new PdfGlyphRun(new[] { new PdfGlyphInfo(1, "A", 0, 500), new PdfGlyphInfo(2, "B", 1, 500) },
            Array.Empty<PdfTextEncodingDiagnostic>(), preserveGlyphUnicode: true);
        var output = new System.Text.StringBuilder();
        new ContentStreamBuilder(output).BeginText().TextMatrix(40D, 400D)
            .ShowText(font.ToTextShowCommand("AB", run, fontMetricScale: 0.75D), cssSize * 0.75D).EndText();
        var origins = new List<double>();
        PdfContentStreamInterpreter.Interpret(output.ToString(), 100, operation => {
            if (operation.Name == "Tm") origins.Add(Convert.ToDouble(operation.Operands[4], System.Globalization.CultureInfo.InvariantCulture));
        });
        double expected = 40D + (500D - adjustment) * cssSize * 0.75D / 1000D;
        // Ordinary paint matrices serialize coordinates to three decimal places.
        Assert.InRange(origins[2], expected - 0.0005D, expected + 0.0005D);
    }

    [Fact]
    public void LogicalCommandAlsoPublishesGlyphBytesWhenActualTextIsSuppressed() {
        var glyphs = new[] { new PdfGlyphInfo(1, "A", 0, 600), new PdfGlyphInfo(2, "B", 1, 600) };
        var command = new PdfGlyphRun(glyphs, Array.Empty<PdfTextEncodingDiagnostic>(), preserveGlyphUnicode: true).ToTextShowCommand();
        var output = new System.Text.StringBuilder();
        new ContentStreamBuilder(output).BeginText().TextMatrix(40D, 400D)
            .ShowText(command, 10D).ShowText(command, 10D, suppressActualText: true).EndText();
        var shows = new List<byte[]>();
        PdfContentStreamInterpreter.Interpret(output.ToString(), 100, operation => {
            if (operation.Name == "Tj") shows.Add((byte[])operation.Operands[0]);
        });
        Assert.Equal(3, shows.Count);
        Assert.Equal(new byte[] { 0, 1 }, shows[0]);
        Assert.Equal(new byte[] { 0, 2 }, shows[1]);
        Assert.Equal(new byte[] { 0, 1, 0, 2 }, shows[2]);
    }
}
