using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfFontEmbeddingPermissionTests {
    [Theory]
    [InlineData(0x0000, "subset")]
    [InlineData(0x0006, "subset")]
    [InlineData(0x000A, "subset")]
    [InlineData(0x0100, "whole")]
    [InlineData(0x0002, "refused")]
    [InlineData(0x0200, "refused")]
    public void GeneratedPdfHonoursTheFontEmbeddingPermission(int fsType, string expected) {
        string? fontPath = PdfComplianceTestFonts.FindBundledTrueTypeFont();
        Assert.NotNull(fontPath);
        byte[] fontData = File.ReadAllBytes(fontPath!);
        WriteFsType(fontData, (ushort)fsType);
        byte[] Generate() => PdfDocument.Create(new PdfCreateOptions { CompressContentStreams = false })
            .UseFontFamily(new PdfEmbeddedFontFamily("OfficeIMO Licence", fontData))
            .Paragraph(paragraph => paragraph.AddText("Licence text"))
            .ToBytes();
        if (expected == "refused") {
            // A family the caller chose is refused when the document is written; installed fallback faces are skipped instead.
            var exception = Assert.Throws<NotSupportedException>(() => Generate());
            Assert.Contains("fsType", exception.Message, StringComparison.Ordinal);
            return;
        }

        byte[] bytes = Generate();

        string raw = Encoding.ASCII.GetString(bytes);
        int[] lengths = Regex.Matches(raw, @"/Length1\s+(\d+)").Cast<Match>().Select(match => int.Parse(match.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture)).ToArray();
        Assert.Contains("Licence text", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
        if (expected == "whole") {
            Assert.Equal(fontData.Length, Assert.Single(lengths));
            Assert.Matches("/BaseFont /OfficeIMOLicence-Regular", raw);
        } else {
            Assert.InRange(Assert.Single(lengths), 1, fontData.Length - 1);
            Assert.Matches("/BaseFont /[A-Z]{6}\\+OfficeIMOLicence-Regular", raw);
        }
    }

    private static void WriteFsType(byte[] font, ushort fsType) {
        int count = (font[4] << 8) | font[5];
        for (int index = 0; index < count; index++) {
            int record = 12 + index * 16;
            if (Encoding.ASCII.GetString(font, record, 4) != "OS/2") continue;
            int offset = (font[record + 8] << 24) | (font[record + 9] << 16) | (font[record + 10] << 8) | font[record + 11];
            font[offset + 8] = (byte)(fsType >> 8);
            font[offset + 9] = (byte)fsType;
            return;
        }
        throw new InvalidOperationException("The test font has no OS/2 table.");
    }
}
