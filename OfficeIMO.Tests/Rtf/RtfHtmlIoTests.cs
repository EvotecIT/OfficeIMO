using OfficeIMO.Rtf;
using OfficeIMO.Html;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
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

        RtfDocument roundTrip = bytes.LoadRtfFromHtml();
        Assert.Equal("Clinical ż", Assert.Single(roundTrip.Paragraphs).ToPlainText());
    }

    [Fact]
    public void Html_To_Rtf_Provides_Text_Bytes_Memory_Stream_And_Stream_Save() {
        const string html = "<p>Clinical ż</p>";
        var writeOptions = new RtfWriteOptions { IncludeGenerator = false };

        string rtf = html.ToRtf(writeOptions: writeOptions);
        byte[] bytes = html.ToRtfBytes(writeOptions: writeOptions);

        Assert.Equal(rtf, Encoding.UTF8.GetString(bytes));
        Assert.Contains(@"\u380?", rtf, StringComparison.Ordinal);

        using MemoryStream memoryStream = html.ToRtfMemoryStream(writeOptions: writeOptions);
        Assert.Equal(bytes, memoryStream.ToArray());

        using var output = new MemoryStream();
        output.WriteByte(0x2A);
        html.SaveAsRtf(output, writeOptions: writeOptions);
        byte[] saved = output.ToArray();

        Assert.Equal(saved.Length, output.Position);
        Assert.Equal(0x2A, saved[0]);
        Assert.Equal(rtf, Encoding.UTF8.GetString(saved, 1, saved.Length - 1));
        Assert.Equal("Clinical ż", Assert.Single(RtfDocument.Load(saved.Skip(1).ToArray()).Document.Paragraphs).ToPlainText());
    }

    [Fact]
    public void RtfHtml_Fluent_Byte_And_Stream_IO_Replaces_Static_Facade() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Fluent ż");

        string html = document.ToHtml();
        byte[] htmlBytes = document.ToHtmlBytes();

        Assert.Equal(html, Encoding.UTF8.GetString(htmlBytes));

        using MemoryStream htmlStream = document.ToHtmlMemoryStream();
        Assert.Equal(htmlBytes, htmlStream.ToArray());

        RtfDocument fromBytes = htmlBytes.LoadRtfFromHtml();
        Assert.Equal("Fluent ż", Assert.Single(fromBytes.Paragraphs).ToPlainText());

        byte[] prefixedHtml = Encoding.UTF8.GetBytes("*" + html);
        using var source = new MemoryStream(prefixedHtml);
        source.Position = 1;
        RtfDocument fromStream = source.LoadRtfFromHtml();

        Assert.Equal(source.Length, source.Position);
        Assert.Equal("Fluent ż", Assert.Single(fromStream.Paragraphs).ToPlainText());

        var writeOptions = new RtfWriteOptions { IncludeGenerator = false };
        byte[] rtfBytes = htmlBytes.ToRtfBytes(writeOptions: writeOptions);
        string rtf = Encoding.UTF8.GetString(rtfBytes);

        Assert.Contains(@"\u380?", rtf, StringComparison.Ordinal);
        Assert.Equal("Fluent ż", Assert.Single(RtfDocument.Load(rtfBytes).Document.Paragraphs).ToPlainText());

        using MemoryStream rtfMemoryStream = htmlBytes.ToRtfMemoryStream(writeOptions: writeOptions);
        Assert.Equal(rtfBytes, rtfMemoryStream.ToArray());

        using var secondSource = new MemoryStream(prefixedHtml);
        secondSource.Position = 1;
        using var output = new MemoryStream();
        output.WriteByte(0x2A);

        secondSource.SaveAsRtf(output, writeOptions: writeOptions);
        byte[] saved = output.ToArray();

        Assert.Equal(secondSource.Length, secondSource.Position);
        Assert.Equal(saved.Length, output.Position);
        Assert.Equal(0x2A, saved[0]);
        Assert.Equal("Fluent ż", Assert.Single(RtfDocument.Load(saved.Skip(1).ToArray()).Document.Paragraphs).ToPlainText());
    }

    [Fact]
    public async Task RtfHtml_File_Loading_Matches_Text_Stream_And_Async_IO() {
        const string html = "<p>File ż</p>";
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".html");

        try {
            File.WriteAllText(path, html, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

            RtfDocument fromFile = HtmlRtfConverterExtensions.LoadRtfFromHtmlFile(path);
            Assert.Equal("File ż", Assert.Single(fromFile.Paragraphs).ToPlainText());

            RtfDocument fromEncodedFile = HtmlRtfConverterExtensions.LoadRtfFromHtmlFile(path, encoding: Encoding.UTF8);
            Assert.Equal("File ż", Assert.Single(fromEncodedFile.Paragraphs).ToPlainText());

            RtfDocument fromAsyncFile = await HtmlRtfConverterExtensions.LoadRtfFromHtmlFileAsync(path);
            Assert.Equal("File ż", Assert.Single(fromAsyncFile.Paragraphs).ToPlainText());
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public async Task RtfHtml_Async_IO_Matches_Fluent_Sync_IO() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Async ż");

        string html = await document.ToHtmlAsync();
        byte[] htmlBytes = await document.ToHtmlBytesAsync();

        Assert.Equal(html, Encoding.UTF8.GetString(htmlBytes));

        using MemoryStream htmlMemoryStream = await document.ToHtmlMemoryStreamAsync();
        Assert.Equal(htmlBytes, htmlMemoryStream.ToArray());

        using var htmlOutput = new MemoryStream();
        htmlOutput.WriteByte(0x2A);
        await document.SaveAsHtmlAsync(htmlOutput);
        byte[] savedHtml = htmlOutput.ToArray();

        Assert.Equal(savedHtml.Length, htmlOutput.Position);
        Assert.Equal(0x2A, savedHtml[0]);
        Assert.Equal(html, Encoding.UTF8.GetString(savedHtml, 1, savedHtml.Length - 1));

        RtfDocument fromBytes = await htmlBytes.LoadRtfFromHtmlAsync();
        Assert.Equal("Async ż", Assert.Single(fromBytes.Paragraphs).ToPlainText());

        byte[] prefixedHtml = Encoding.UTF8.GetBytes("*" + html);
        using var source = new MemoryStream(prefixedHtml);
        source.Position = 1;
        RtfDocument fromStream = await source.LoadRtfFromHtmlAsync();

        Assert.Equal(source.Length, source.Position);
        Assert.Equal("Async ż", Assert.Single(fromStream.Paragraphs).ToPlainText());

        var writeOptions = new RtfWriteOptions { IncludeGenerator = false };
        string rtf = await html.ToRtfAsync(writeOptions: writeOptions);
        byte[] rtfBytes = await htmlBytes.ToRtfBytesAsync(writeOptions: writeOptions);

        Assert.Equal(rtf, Encoding.UTF8.GetString(rtfBytes));
        Assert.Contains(@"\u380?", rtf, StringComparison.Ordinal);

        using MemoryStream rtfMemoryStream = await htmlBytes.ToRtfMemoryStreamAsync(writeOptions: writeOptions);
        Assert.Equal(rtfBytes, rtfMemoryStream.ToArray());

        using var secondSource = new MemoryStream(prefixedHtml);
        secondSource.Position = 1;
        using var output = new MemoryStream();
        output.WriteByte(0x2A);

        await secondSource.SaveAsRtfAsync(output, writeOptions: writeOptions);
        byte[] savedRtf = output.ToArray();

        Assert.Equal(secondSource.Length, secondSource.Position);
        Assert.Equal(savedRtf.Length, output.Position);
        Assert.Equal(0x2A, savedRtf[0]);
        Assert.Equal("Async ż", Assert.Single(RtfDocument.Load(savedRtf.Skip(1).ToArray()).Document.Paragraphs).ToPlainText());
    }

    [Fact]
    public async Task RtfHtml_Async_IO_Honors_Cancellation() {
        using var cts = new CancellationTokenSource();
        cts.Cancel();

        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            "<p>Cancelled</p>".LoadRtfFromHtmlAsync(cancellationToken: cts.Token));
        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            HtmlRtfConverterExtensions.LoadRtfFromHtmlFileAsync("ignored.html", cancellationToken: cts.Token));
    }
}
