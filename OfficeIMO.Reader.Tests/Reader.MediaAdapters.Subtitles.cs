using OfficeIMO.Reader;
using OfficeIMO.Reader.Subtitles;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class ReaderMediaAdapterTests {
    private static string EscapeSubtitleMarkdown(string value) => value
        .Replace("&", "&amp;")
        .Replace("<", "&lt;")
        .Replace(">", "&gt;");

    [Theory]
    [InlineData("first<br>second")]
    [InlineData("first<BR/>second")]
    public void SubtitleAdapter_PreservesLineBreaksWhenStrippingBreakMarkup(string cueText) {
        string srt = "1\n00:00:00,000 --> 00:00:01,000\n" + cueText + "\n";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddSubtitleHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes(srt), "break.srt");

        ReaderChunk chunk = Assert.Single(result.Chunks);
        Assert.Equal("first\nsecond", chunk.Text);
        Assert.EndsWith("first\nsecond", chunk.Markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void SubtitleAdapter_EncodedMarkupRemainsLiteralInMarkdown() {
        const string srt = "1\n00:00:00,000 --> 00:00:01,000\n&lt;img src=x onerror=alert(1)&gt;\n";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddSubtitleHandler().Build();

        ReaderChunk chunk = Assert.Single(reader.ReadDocument(
            Encoding.UTF8.GetBytes(srt), "encoded.srt").Chunks);

        Assert.Equal("<img src=x onerror=alert(1)>", chunk.Text);
        Assert.DoesNotContain("<img", chunk.Markdown, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("&lt;img", chunk.Markdown, StringComparison.Ordinal);
    }
}
