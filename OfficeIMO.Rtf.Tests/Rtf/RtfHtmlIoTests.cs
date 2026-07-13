using OfficeIMO.Html;
using OfficeIMO.Rtf;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlIoTests {
    [Fact]
    public void RtfHtml_Output_Provides_Text_Bytes_And_Stream() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Clinical ż");

        string html = document.ToHtml(RtfToHtmlOptions.CreateRoundTripProfile());
        byte[] bytes = document.ToHtmlBytes(RtfToHtmlOptions.CreateRoundTripProfile());

        Assert.Equal(html, Encoding.UTF8.GetString(bytes));
        using MemoryStream stream = document.ToHtmlStream(RtfToHtmlOptions.CreateRoundTripProfile());
        Assert.Equal(bytes, stream.ToArray());

        RtfDocument roundTrip = HtmlConversionDocument.Parse(html).ToRtfDocument();
        Assert.Equal("Clinical ż", Assert.Single(roundTrip.Paragraphs).ToPlainText());
    }

    [Fact]
    public void Html_To_Rtf_Uses_One_Prepared_Source_For_All_Target_Forms() {
        HtmlConversionDocument source = HtmlConversionDocument.Parse("<p>Clinical ż</p>");
        var writeOptions = new RtfWriteOptions { IncludeGenerator = false };

        string rtf = source.ToRtf(writeOptions: writeOptions);
        byte[] bytes = source.ToRtfBytes(writeOptions: writeOptions);
        using MemoryStream memoryStream = source.ToRtfStream(writeOptions: writeOptions);

        Assert.Equal(rtf, Encoding.UTF8.GetString(bytes));
        Assert.Equal(bytes, memoryStream.ToArray());
        Assert.Contains(@"\u380?", rtf, StringComparison.Ordinal);

        using var output = new MemoryStream();
        output.WriteByte(0x2A);
        source.SaveAsRtf(output, writeOptions: writeOptions);

        Assert.Equal(0, output.Position);
        Assert.Equal(bytes, output.ToArray());
        Assert.Equal("Clinical ż", Assert.Single(RtfDocument.Load(bytes).Document.Paragraphs).ToPlainText());
    }

    [Fact]
    public async Task Html_Lifecycle_Owns_File_And_Stream_Loading_Before_Rtf_Conversion() {
        const string html = "<p>File ż</p>";
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".html");

        try {
            File.WriteAllText(path, html, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

            HtmlConversionDocument fromFile = HtmlConversionDocument.Load(path);
            HtmlConversionDocument fromAsyncFile = await HtmlConversionDocument.LoadAsync(path);
            Assert.Equal("File ż", Assert.Single(fromFile.ToRtfDocument().Paragraphs).ToPlainText());
            Assert.Equal("File ż", Assert.Single(fromAsyncFile.ToRtfDocument().Paragraphs).ToPlainText());

            byte[] prefixedHtml = Encoding.UTF8.GetBytes("*" + html);
            using var sourceStream = new MemoryStream(prefixedHtml);
            sourceStream.Position = 1;
            HtmlConversionDocument fromStream = HtmlConversionDocument.Load(sourceStream);

            Assert.Equal(1, sourceStream.Position);
            Assert.Equal("*<p>File ż</p>", fromStream.SourceHtml);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task Html_To_Rtf_Async_APIs_Are_Target_IO_Only() {
        HtmlConversionDocument source = HtmlConversionDocument.Parse("<p>Async ż</p>");
        var writeOptions = new RtfWriteOptions { IncludeGenerator = false };
        byte[] expected = source.ToRtfBytes(writeOptions: writeOptions);

        using var output = new MemoryStream();
        output.WriteByte(0x2A);
        await source.SaveAsRtfAsync(output, writeOptions: writeOptions);

        Assert.Equal(0, output.Position);
        Assert.Equal(expected, output.ToArray());

        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            source.SaveAsRtfAsync(new MemoryStream(), cancellationToken: cancelled.Token));
    }

    [Fact]
    public void Html_To_Rtf_Public_Contract_Rejects_Raw_Input_Overload_Forests() {
        MethodInfo[] methods = typeof(HtmlRtfConverterExtensions).GetMethods(BindingFlags.Public | BindingFlags.Static);
        MethodInfo[] htmlToRtf = methods
            .Where(method => method.Name.Contains("Rtf", StringComparison.Ordinal) &&
                             method.GetParameters().Length > 0 &&
                             method.GetParameters()[0].ParameterType != typeof(RtfDocument))
            .ToArray();

        Assert.NotEmpty(htmlToRtf);
        Assert.All(htmlToRtf, method =>
            Assert.Equal(typeof(HtmlConversionDocument), method.GetParameters()[0].ParameterType));
        Assert.DoesNotContain(methods, method => method.Name.Contains("FromHtmlFile", StringComparison.Ordinal));
    }
}
