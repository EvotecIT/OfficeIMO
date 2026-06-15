using OfficeIMO.Rtf.Pdf;
using PdfCore = OfficeIMO.Pdf;
using System.IO;
using System.Linq;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfPdfIoTests {
    [Fact]
    public void RtfPdf_Converts_Source_Bytes_To_Pdf_Bytes_Stream_And_File() {
        byte[] rtfBytes = Encoding.ASCII.GetBytes(@"{\rtf1\ansi\pard Byte PDF\par}");

        byte[] pdf = rtfBytes.SaveAsPdf();
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();

        Assert.Equal("%PDF-", Encoding.ASCII.GetString(pdf, 0, 5));
        Assert.Contains("Byte PDF", text, StringComparison.Ordinal);

        using var output = new MemoryStream();
        output.WriteByte(0x2A);
        rtfBytes.SaveAsPdf(output);
        byte[] streamed = output.ToArray();

        Assert.Equal(streamed.Length, output.Position);
        Assert.Equal(0x2A, streamed[0]);
        Assert.Equal("%PDF-", Encoding.ASCII.GetString(streamed, 1, 5));

        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pdf");
        try {
            rtfBytes.SaveAsPdf(path);
            Assert.Contains("Byte PDF", PdfCore.PdfReadDocument.Load(path).ExtractText(), StringComparison.Ordinal);
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void RtfPdf_Converts_Source_Stream_From_Current_Position() {
        byte[] rtfBytes = Encoding.ASCII.GetBytes(@"{\rtf1\ansi\pard Stream PDF\par}");
        using var source = new MemoryStream();
        source.WriteByte(0x2A);
        source.Write(rtfBytes, 0, rtfBytes.Length);
        source.Position = 1;

        byte[] pdf = source.SaveAsPdf();

        Assert.Equal(source.Length, source.Position);
        Assert.Equal("%PDF-", Encoding.ASCII.GetString(pdf, 0, 5));
        Assert.Contains("Stream PDF", PdfCore.PdfReadDocument.Load(pdf).ExtractText(), StringComparison.Ordinal);

        using var secondSource = new MemoryStream();
        secondSource.WriteByte(0x2A);
        secondSource.Write(rtfBytes, 0, rtfBytes.Length);
        secondSource.Position = 1;
        using var output = new MemoryStream();
        output.WriteByte(0x2A);

        secondSource.SaveAsPdf(output);
        byte[] streamed = output.ToArray();

        Assert.Equal(secondSource.Length, secondSource.Position);
        Assert.Equal(streamed.Length, output.Position);
        Assert.Equal("%PDF-", Encoding.ASCII.GetString(streamed, 1, 5));
        Assert.Contains("Stream PDF", PdfCore.PdfReadDocument.Load(streamed.Skip(1).ToArray()).ExtractText(), StringComparison.Ordinal);
    }
}
