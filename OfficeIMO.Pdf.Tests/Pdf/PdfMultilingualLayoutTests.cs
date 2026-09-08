using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfMultilingualLayoutTests {
    [Fact]
    public void SearchableSelectionBoxesWithDifferentGlyphHeightsRemainOneLine() {
        byte[] bytes = PdfDocument.Create().Canvas(canvas => canvas
            .SearchableText("Raport", 47D, 55.429D, 84.45D, 21.356D)
            .SearchableText("jakosci", 138.913D, 53.509D, 86.61D, 23.276D)
            .SearchableText("dokumentu", 236.321D, 54.709D, 142.03D, 17.517D)
            .SearchableText("Separate row", 47D, 90D, 160D, 21D))
            .ToBytes();
        PdfLogicalPage page = Assert.Single(PdfDocument.Load(bytes)
            .Read(new PdfReadOptions { Profile = PdfReadProfile.Structured }).Pages);
        Assert.Equal(new[] { "Raport jakosci dokumentu", "Separate row" },
            page.Analysis.Lines.Select(static line => line.Text));
        PdfUnderstandingLine heading = page.Analysis.Lines[0];
        Assert.Equal(3, heading.Words.Count);
        // The layout correction must not rewrite the distinct baselines used by selection boxes.
        Assert.True(heading.Words.Max(static word => word.BaselineY) -
            heading.Words.Min(static word => word.BaselineY) > 4D);
        Assert.All(heading.Words.SelectMany(static word => word.SourceRuns),
            run => Assert.Contains(run, page.Analysis.DecodedRuns));
    }

    [Theory]
    [InlineData("latin-0")]
    [InlineData("latin-90")]
    [InlineData("latin-180")]
    [InlineData("latin-270")]
    [InlineData("rtl-0")]
    [InlineData("rtl-90")]
    [InlineData("rtl-180")]
    [InlineData("rtl-270")]
    public void IndependentNativeLayoutPreservesTextOrderTablesCaptionsAndSourceGeometry(string id) {
        string root = Path.Combine(AppContext.BaseDirectory, "MultilingualLayout");
        using JsonDocument manifest = JsonDocument.Parse(File.ReadAllBytes(Path.Combine(root, "manifest.json")));
        JsonElement fixture = manifest.RootElement.GetProperty("cases").EnumerateArray()
            .Single(value => value.GetProperty("id").GetString() == id);
        byte[] source = File.ReadAllBytes(Path.Combine(root, id + "-native.pdf"));
        using (SHA256 digest = SHA256.Create()) {
            string hash = BitConverter.ToString(digest.ComputeHash(source)).Replace("-", string.Empty);
            Assert.Equal(fixture.GetProperty("files").GetProperty("native.pdf").GetString()!.ToUpperInvariant(), hash);
        }
        string[] expected = fixture.GetProperty("readingOrder").EnumerateArray().Select(value => value.GetString()!).ToArray();
        PdfDocumentReadResult read = PdfDocument.Load(source).Read(new PdfReadOptions { Profile = PdfReadProfile.Structured });
        Assert.Equal(Normalize(string.Join(" ", expected)), Normalize(read.Text));
        PdfLogicalPage page = Assert.Single(read.Pages);
        PdfLogicalTable table = Assert.Single(page.Tables);
        string[][] expectedRows = fixture.GetProperty("table").EnumerateArray()
            .Select(row => row.EnumerateArray().Select(value => value.GetString()!).ToArray()).ToArray();
        Assert.Equal(expectedRows.Length, table.Rows.Count);
        for (int index = 0; index < expectedRows.Length; index++) Assert.Equal(expectedRows[index], table.Rows[index]);
        Assert.Equal(fixture.GetProperty("caption").GetString(), Assert.Single(page.Captions).Text);
        Assert.Contains(page.Headings, heading => heading.Text == expected[0]);

        double expectedAngle = fixture.GetProperty("clockwiseDegrees").GetInt32() switch { 90 => -90D, 180 => 180D, 270 => 90D, _ => 0D };
        Assert.All(page.Analysis.Lines, line => {
            Assert.InRange(Math.Abs(Math.Abs(line.RotationDegrees) - Math.Abs(expectedAngle)), 0D, 0.001D);
            Assert.All(line.Words.SelectMany(word => word.SourceRuns), run => Assert.Contains(run, page.Analysis.DecodedRuns));
            if (expectedAngle != 0D) {
                Assert.NotNull(line.VisualBounds);
                Assert.True(line.VisualBounds!.Right > line.VisualBounds.Left);
                Assert.True(line.VisualBounds.Bottom > line.VisualBounds.Top);
            }
        });
    }

    private static string Normalize(string value) => Regex.Replace(value.Normalize(NormalizationForm.FormC), @"\s+", " ").Trim();
}
