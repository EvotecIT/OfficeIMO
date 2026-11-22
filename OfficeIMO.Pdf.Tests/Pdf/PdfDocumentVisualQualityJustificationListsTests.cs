using OfficeIMO.Pdf;
using Letter = UglyToad.PdfPig.Content.Letter;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    // A justified list fills every line of an item to the right edge and leaves the last line natural,
    // unless the list style asks for the last line to be stretched as well.
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void JustifiedListFillsItemLinesAndFollowsTheLastLinePolicy(bool numbered, bool justifyLastLine) {
        const string text = "alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi rho sigma tau LAST";
        var style = new PdfListStyle { JustifyLastLine = justifyLastLine };
        PdfDocument document = PdfDocument.Create(JustificationPolicyOptions());
        document = numbered
            ? document.Numbered(new[] { text }, OfficeAlignment.Justify, style: style)
            : document.Bullets(new[] { text }, OfficeAlignment.Justify, style: style);
        Letter[] letters = ReadJustificationPolicyLetters(document.ToBytes());

        var lines = letters.GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .OrderByDescending(line => line.Key).Select(line => line.OrderBy(letter => letter.StartBaseLine.X).ToList()).ToList();
        Assert.True(lines.Count >= 3, $"The item wraps to {lines.Count} lines.");
        foreach (List<Letter> line in lines.Take(lines.Count - 1))
            Assert.InRange(line.Last().EndBaseLine.X, 269.95D, 270.05D);
        double lastEnd = lines.Last().Last().EndBaseLine.X;
        if (justifyLastLine) Assert.InRange(lastEnd, 269.95D, 270.05D);
        else Assert.True(lastEnd < 260D, $"The last line ends at {lastEnd:F2}.");
    }
}
