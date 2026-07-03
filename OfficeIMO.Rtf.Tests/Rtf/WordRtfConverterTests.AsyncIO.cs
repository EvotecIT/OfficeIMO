using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public partial class WordRtfConverterTests {
    [Fact]
    public async Task Word_Rtf_Async_IO_Provides_Byte_Stream_And_File_Loading_Surface() {
        using WordDocument word = WordDocument.Create();
        word.AddParagraph("Async clinical ż");
        var options = new RtfWriteOptions { IncludeGenerator = false };

        string rtf = await word.ToRtfAsync(options);
        byte[] bytes = await word.ToRtfBytesAsync(options);

        Assert.Equal(rtf, Encoding.UTF8.GetString(bytes));

        using MemoryStream memoryStream = await word.ToRtfMemoryStreamAsync(options);
        Assert.Equal(bytes, memoryStream.ToArray());

        using var output = new MemoryStream();
        output.WriteByte(0x2A);
        await word.SaveAsRtfAsync(output, options);
        byte[] saved = output.ToArray();

        Assert.Equal(saved.Length, output.Position);
        Assert.Equal(0x2A, saved[0]);
        Assert.Equal(rtf, Encoding.UTF8.GetString(saved, 1, saved.Length - 1));

        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".rtf");
        try {
            await word.SaveAsRtfAsync(path, options, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            Assert.Equal(bytes, File.ReadAllBytes(path));

            using WordDocument fromFile = await WordRtfConverterExtensions.LoadFromRtfFileAsync(path);
            Assert.Contains("Async clinical ż", string.Concat(fromFile.Paragraphs.Select(paragraph => paragraph.Text)), StringComparison.Ordinal);
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }

        using WordDocument fromText = await rtf.LoadFromRtfAsync();
        Assert.Contains("Async clinical ż", string.Concat(fromText.Paragraphs.Select(paragraph => paragraph.Text)), StringComparison.Ordinal);

        using WordDocument fromBytes = await bytes.LoadFromRtfAsync();
        Assert.Contains("Async clinical ż", string.Concat(fromBytes.Paragraphs.Select(paragraph => paragraph.Text)), StringComparison.Ordinal);

        using var input = new MemoryStream();
        input.WriteByte(0x2A);
        input.Write(bytes, 0, bytes.Length);
        input.Position = 1;

        using WordDocument fromStream = await input.LoadFromRtfAsync();
        Assert.Equal(input.Length, input.Position);
        Assert.Contains("Async clinical ż", string.Concat(fromStream.Paragraphs.Select(paragraph => paragraph.Text)), StringComparison.Ordinal);
    }

    [Fact]
    public async Task Word_Rtf_Async_IO_Honors_Cancellation_Before_Work_Starts() {
        using WordDocument word = WordDocument.Create();
        word.AddParagraph("Cancelled");

        using var cts = new CancellationTokenSource();
        cts.Cancel();

        await Assert.ThrowsAsync<OperationCanceledException>(() => word.ToRtfAsync(cancellationToken: cts.Token));
        await Assert.ThrowsAsync<OperationCanceledException>(() => word.ToRtfBytesAsync(cancellationToken: cts.Token));
        await Assert.ThrowsAsync<OperationCanceledException>(() => "ignored".LoadFromRtfAsync(cancellationToken: cts.Token));
        await Assert.ThrowsAsync<OperationCanceledException>(() => new byte[] { 123, 125 }.LoadFromRtfAsync(cancellationToken: cts.Token));
    }
}
