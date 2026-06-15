using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Html;
using System.IO;
using System.Linq;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlIoTests {
    [Fact]
    public void RtfHtml_Output_Provides_Encoded_Bytes_And_Memory_Stream() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Clinical ż");

        string html = document.ToHtml();
        byte[] bytes = document.ToHtmlBytes();

        Assert.Equal(html, Encoding.UTF8.GetString(bytes));

        using MemoryStream memoryStream = document.ToHtmlMemoryStream();
        Assert.Equal(bytes, memoryStream.ToArray());

        RtfDocument roundTrip = bytes.ToRtfDocumentFromHtml();
        Assert.Equal("Clinical ż", Assert.Single(roundTrip.Paragraphs).ToPlainText());
    }

    [Fact]
    public void Html_To_Rtf_Provides_Text_Bytes_Memory_Stream_And_Stream_Save() {
        const string html = "<p>Clinical ż</p>";
        var writeOptions = new RtfWriteOptions { IncludeGenerator = false };

        string rtf = html.ToRtfFromHtml(writeOptions: writeOptions);
        byte[] bytes = html.ToRtfBytesFromHtml(writeOptions: writeOptions);

        Assert.Equal(rtf, Encoding.UTF8.GetString(bytes));
        Assert.Contains(@"\u380?", rtf, StringComparison.Ordinal);

        using MemoryStream memoryStream = html.ToRtfMemoryStreamFromHtml(writeOptions: writeOptions);
        Assert.Equal(bytes, memoryStream.ToArray());

        using var output = new MemoryStream();
        output.WriteByte(0x2A);
        html.SaveAsRtfFromHtml(output, writeOptions: writeOptions);
        byte[] saved = output.ToArray();

        Assert.Equal(saved.Length, output.Position);
        Assert.Equal(0x2A, saved[0]);
        Assert.Equal(rtf, Encoding.UTF8.GetString(saved, 1, saved.Length - 1));
        Assert.Equal("Clinical ż", Assert.Single(RtfDocument.Load(saved.Skip(1).ToArray()).Document.Paragraphs).ToPlainText());
    }
}
