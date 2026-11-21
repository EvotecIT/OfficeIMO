using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfFontFamilyTests {
    [Fact]
    public void SyntheticStylesOfOneFaceShareItsEmbeddedFontProgram() {
        byte[] regular = File.ReadAllBytes(Assert.IsType<string>(PdfComplianceTestFonts.FindBundledTrueTypeFont()));
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions { CompressContentStreams = false })
            .UseFontFamily(new PdfEmbeddedFontFamily("Single Face", regular))
            .Paragraph(paragraph => {
                paragraph.AddText("Plain ");
                paragraph.AddText("Slanted ").SetItalic();
                paragraph.AddText("Heavy ").SetBold();
                paragraph.AddText("Both").SetBold().SetItalic();
            })
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        // One embedded program; each style keeps its own font dictionary and name.
        Assert.Single(Regex.Matches(raw, @"/FontFile2 (\d+) 0 R").Cast<Match>().Select(match => match.Groups[1].Value).Distinct());
        foreach (string face in new[] { "Regular", "Bold", "Italic", "BoldItalic" }) {
            Assert.Matches(PdfTestContentText.BaseFontPattern("SingleFace-" + face), raw);
        }
        Assert.Contains("Plain Slanted Heavy Both", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
    }
}
