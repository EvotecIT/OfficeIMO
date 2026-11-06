using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfBrowserColumnReadingTests {
    [Theory]
    [InlineData("two-columns-narrow-gutter.pdf", 6)]
    [InlineData("three-columns-wrapped-prose.pdf", 9)]
    [InlineData("three-columns-ragged-gutter.pdf", 9)]
    public void ReadKeepsCompleteWrappedParagraphsInColumnOrder(string file, int paragraphCount) {
        string path = Path.Combine(AppContext.BaseDirectory, "Pdf", "Fixtures", "Interoperability", "BrowserColumns", file);
        PdfDocumentReadResult result = PdfDocument.Load(path).Read();
        string expected = string.Join(" ", Enumerable.Range(1, paragraphCount).Select(number =>
            $"Para{number} begins here and continues with enough ordinary words so that it wraps over several lines inside its column for testing."));

        Assert.Empty(result.Tables);
        Assert.Equal(expected, Regex.Replace(result.Text, @"\s+", " ").Trim());
        Assert.Equal(expected, Regex.Replace(result.ExtractMarkdown(), @"\s+", " ").Trim());
    }
}
