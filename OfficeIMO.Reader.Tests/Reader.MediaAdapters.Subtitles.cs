using OfficeIMO.Reader;
using OfficeIMO.Reader.Subtitles;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class ReaderMediaAdapterTests {
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
}
