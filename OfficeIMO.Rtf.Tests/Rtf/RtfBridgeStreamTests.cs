using OfficeIMO.Rtf;
using OfficeIMO.Html;
using OfficeIMO.Rtf.Pdf;
using PdfCore = OfficeIMO.Pdf;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfBridgeStreamTests {
    [Fact]
    public void SaveAsHtml_Overwrites_And_Rewinds_Seekable_Stream() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Clinical note");

        using var stream = new MemoryStream();
        stream.WriteByte(0x2A);

        document.SaveAsHtml(stream);

        byte[] bytes = stream.ToArray();
        Assert.Equal(0, stream.Position);
        Assert.Contains("<p>Clinical note</p>", Encoding.UTF8.GetString(bytes), StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_Overwrites_And_Rewinds_Seekable_Stream() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Clinical note");

        using var stream = new MemoryStream();
        stream.WriteByte(0x2A);

        document.SaveAsPdf(stream);

        byte[] bytes = stream.ToArray();
        Assert.Equal(0, stream.Position);
        Assert.Equal("%PDF-", Encoding.ASCII.GetString(bytes, 0, 5));
    }

    [Fact]
    public void TrySaveAsPdf_Overwrites_And_Rewinds_Seekable_Stream() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Clinical note");

        using var stream = new MemoryStream();
        stream.WriteByte(0x2A);

        PdfCore.PdfSaveResult result = document.TrySaveAsPdf(stream);

        byte[] bytes = stream.ToArray();
        Assert.True(result.Succeeded, result.Exception?.Message);
        Assert.Equal(0, stream.Position);
        Assert.Equal("%PDF-", Encoding.ASCII.GetString(bytes, 0, 5));
    }

    [Fact]
    public async Task SaveAsPdfAsync_Overwrites_And_Rewinds_Seekable_Stream() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Clinical note");

        using var stream = new MemoryStream();
        stream.WriteByte(0x2A);

        await document.SaveAsPdfAsync(stream);

        byte[] bytes = stream.ToArray();
        Assert.Equal(0, stream.Position);
        Assert.Equal("%PDF-", Encoding.ASCII.GetString(bytes, 0, 5));
    }
}
