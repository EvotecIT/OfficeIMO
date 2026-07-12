using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using System.IO;
using System.Linq;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public partial class WordRtfConverterTests {
    [Fact]
    public void Word_Rtf_IO_Provides_Byte_Stream_And_File_Loading_Surface() {
        using WordDocument word = WordDocument.Create();
        word.AddParagraph("Clinical ż");
        var options = new RtfWriteOptions { IncludeGenerator = false };

        byte[] bytes = word.ToRtfBytes(options);
        string rtf = word.ToRtf(options);

        Assert.Equal(rtf, Encoding.UTF8.GetString(bytes));

        using MemoryStream memoryStream = word.ToRtfMemoryStream(options);
        Assert.Equal(bytes, memoryStream.ToArray());

        using var output = new MemoryStream();
        output.WriteByte(0x2A);
        word.SaveAsRtf(output, options);
        byte[] saved = output.ToArray();

        Assert.Equal(0, output.Position);
        Assert.Equal(bytes, saved);
        Assert.Equal(rtf, Encoding.UTF8.GetString(saved));

        using WordDocument fromBytes = bytes.LoadFromRtf();
        Assert.Contains("Clinical ż", string.Concat(fromBytes.Paragraphs.Select(paragraph => paragraph.Text)), StringComparison.Ordinal);

        using var input = new MemoryStream();
        input.WriteByte(0x2A);
        input.Write(bytes, 0, bytes.Length);
        input.Position = 1;
        using WordDocument fromStream = input.LoadFromRtf();
        Assert.Equal(input.Length, input.Position);
        Assert.Contains("Clinical ż", string.Concat(fromStream.Paragraphs.Select(paragraph => paragraph.Text)), StringComparison.Ordinal);

        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".rtf");
        try {
            File.WriteAllBytes(path, bytes);
            using WordDocument fromFile = WordRtfConverterExtensions.LoadFromRtfFile(path);
            Assert.Contains("Clinical ż", string.Concat(fromFile.Paragraphs.Select(paragraph => paragraph.Text)), StringComparison.Ordinal);
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }
}
