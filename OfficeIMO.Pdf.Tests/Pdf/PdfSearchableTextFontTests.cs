using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfSearchableTextFontTests {
    [Theory]
    [InlineData(0, false)]
    [InlineData(1, false)]
    [InlineData(2, false)]
    [InlineData(0, true)]
    [InlineData(1, true)]
    [InlineData(2, true)]
    public void SearchableUnicodeDoesNotDependOnActualTextReplacement(int fontKind, bool crossFontBank) {
        var options = new PdfOptions { CompressContentStreams = false };
        if (crossFontBank) {
            options.FileVersion = PdfFileVersion.Pdf17;
            options.ObjectSerializationMode = PdfObjectSerializationMode.ForwardOnly;
        }
        if (fontKind != 0) {
            string? path = fontKind == 1 ? PdfComplianceTestFonts.FindBundledTrueTypeFont() : PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
            Assert.NotNull(path);
            options.EmbedStandardFont(PdfStandardFont.Helvetica, File.ReadAllBytes(path!));
        }
        string text = crossFontBank
            ? string.Concat(Enumerable.Range(0x400, 270).Select(value => ((char)value).ToString())) + "😀fi A"
            : "Searchable XPS fi A😀";
        byte[] bytes = PdfDocument.Create(document => document.Content(content => content.Canvas(canvas =>
            canvas.SearchableText(text, 20, 30, 400, 20))), options).ToBytes();
        Assert.Equal(text, PdfReadDocument.Open(bytes).ExtractText().Trim());

        // Preserve offsets while disabling replacement semantics, as readers that
        // ignore multi-character ActualText still need the font's Unicode contract.
        byte[] withoutReplacement = Encoding.GetEncoding(28591).GetBytes(Encoding.GetEncoding(28591).GetString(bytes).Replace("/ActualText", "/UnusedText"));
        Assert.Equal(text, PdfReadDocument.Open(withoutReplacement).ExtractText().Trim());
    }
}
