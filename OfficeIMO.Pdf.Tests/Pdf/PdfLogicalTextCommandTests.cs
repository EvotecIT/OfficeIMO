using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfLogicalTextCommandTests {
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
