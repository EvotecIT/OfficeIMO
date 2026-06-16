using OfficeIMO.Rtf.Pdf;
using PdfCore = OfficeIMO.Pdf;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
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

    [Fact]
    public async Task RtfPdf_Converts_Source_File_To_Document_Bytes_File_And_Stream() {
        const string rtf = @"{\rtf1\ansi\pard File PDF\par}";
        string rtfPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".rtf");
        string pdfPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pdf");

        try {
            File.WriteAllText(rtfPath, rtf, Encoding.ASCII);

            PdfCore.PdfDocument pdfDocument = rtfPath.ToPdfDocumentFromRtfFile(encoding: Encoding.ASCII);
            Assert.Contains("File PDF", PdfCore.PdfReadDocument.Load(pdfDocument.ToBytes()).ExtractText(), StringComparison.Ordinal);

            byte[] pdfBytes = RtfPdfConverterExtensions.SaveRtfFileAsPdf(rtfPath, encoding: Encoding.ASCII);
            Assert.Equal("%PDF-", Encoding.ASCII.GetString(pdfBytes, 0, 5));
            Assert.Contains("File PDF", PdfCore.PdfReadDocument.Load(pdfBytes).ExtractText(), StringComparison.Ordinal);

            RtfPdfConverterExtensions.SaveRtfFileAsPdf(rtfPath, pdfPath, encoding: Encoding.ASCII);
            Assert.Contains("File PDF", PdfCore.PdfReadDocument.Load(pdfPath).ExtractText(), StringComparison.Ordinal);

            using var output = new MemoryStream();
            output.WriteByte(0x2A);
            RtfPdfConverterExtensions.SaveRtfFileAsPdf(rtfPath, output, encoding: Encoding.ASCII);
            byte[] streamed = output.ToArray();

            Assert.Equal(streamed.Length, output.Position);
            Assert.Equal(0x2A, streamed[0]);
            Assert.Equal("%PDF-", Encoding.ASCII.GetString(streamed, 1, 5));
            Assert.Contains("File PDF", PdfCore.PdfReadDocument.Load(streamed.Skip(1).ToArray()).ExtractText(), StringComparison.Ordinal);

            PdfCore.PdfSaveResult tryResult = RtfPdfConverterExtensions.TrySaveRtfFileAsPdf(rtfPath, output, encoding: Encoding.ASCII);
            Assert.True(tryResult.Succeeded);

            PdfCore.PdfDocument asyncDocument = await rtfPath.ToPdfDocumentFromRtfFileAsync(encoding: Encoding.ASCII);
            Assert.Contains("File PDF", PdfCore.PdfReadDocument.Load(asyncDocument.ToBytes()).ExtractText(), StringComparison.Ordinal);

            byte[] asyncBytes = await RtfPdfConverterExtensions.SaveRtfFileAsPdfAsync(rtfPath, encoding: Encoding.ASCII);
            Assert.Contains("File PDF", PdfCore.PdfReadDocument.Load(asyncBytes).ExtractText(), StringComparison.Ordinal);
        } finally {
            if (File.Exists(rtfPath)) {
                File.Delete(rtfPath);
            }

            if (File.Exists(pdfPath)) {
                File.Delete(pdfPath);
            }
        }
    }

    [Fact]
    public async Task RtfPdf_Async_Converts_String_Bytes_And_Source_Stream() {
        const string rtf = @"{\rtf1\ansi\pard Async PDF\par}";
        byte[] rtfBytes = Encoding.ASCII.GetBytes(rtf);

        byte[] fromString = await rtf.SaveAsPdfAsync();
        byte[] fromBytes = await rtfBytes.SaveAsPdfAsync();

        Assert.Equal("%PDF-", Encoding.ASCII.GetString(fromString, 0, 5));
        Assert.Equal("%PDF-", Encoding.ASCII.GetString(fromBytes, 0, 5));
        Assert.Contains("Async PDF", PdfCore.PdfReadDocument.Load(fromString).ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Async PDF", PdfCore.PdfReadDocument.Load(fromBytes).ExtractText(), StringComparison.Ordinal);

        using var source = new MemoryStream();
        source.WriteByte(0x2A);
        source.Write(rtfBytes, 0, rtfBytes.Length);
        source.Position = 1;
        byte[] fromStream = await source.SaveAsPdfAsync(encoding: Encoding.ASCII);

        Assert.Equal(source.Length, source.Position);
        Assert.Equal("%PDF-", Encoding.ASCII.GetString(fromStream, 0, 5));
        Assert.Contains("Async PDF", PdfCore.PdfReadDocument.Load(fromStream).ExtractText(), StringComparison.Ordinal);

        using var secondSource = new MemoryStream();
        secondSource.WriteByte(0x2A);
        secondSource.Write(rtfBytes, 0, rtfBytes.Length);
        secondSource.Position = 1;
        using var output = new MemoryStream();
        output.WriteByte(0x2A);

        await secondSource.SaveAsPdfAsync(output, encoding: Encoding.ASCII);
        byte[] streamed = output.ToArray();

        Assert.Equal(secondSource.Length, secondSource.Position);
        Assert.Equal(streamed.Length, output.Position);
        Assert.Equal("%PDF-", Encoding.ASCII.GetString(streamed, 1, 5));
        Assert.Contains("Async PDF", PdfCore.PdfReadDocument.Load(streamed.Skip(1).ToArray()).ExtractText(), StringComparison.Ordinal);

        using var cts = new CancellationTokenSource();
        cts.Cancel();

        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            RtfPdfConverterExtensions.SaveRtfFileAsPdfAsync("ignored.rtf", encoding: Encoding.ASCII, cancellationToken: cts.Token));
    }

    [Fact]
    public async Task RtfPdf_TrySave_Source_Overloads_Return_Diagnostics() {
        const string rtf = @"{\rtf1\ansi\pard Try PDF\par}";
        byte[] rtfBytes = Encoding.ASCII.GetBytes(rtf);
        using var output = new MemoryStream();

        PdfCore.PdfSaveResult streamResult = rtf.TrySaveAsPdf(output);

        Assert.True(streamResult.Succeeded);
        Assert.Equal(output.Length, streamResult.BytesWritten);
        Assert.Equal("%PDF-", Encoding.ASCII.GetString(output.ToArray(), 0, 5));
        Assert.Contains("Try PDF", PdfCore.PdfReadDocument.Load(output.ToArray()).ExtractText(), StringComparison.Ordinal);

        using var source = new MemoryStream();
        source.WriteByte(0x2A);
        source.Write(rtfBytes, 0, rtfBytes.Length);
        source.Position = 1;
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pdf");
        try {
            PdfCore.PdfSaveResult pathResult = source.TrySaveAsPdf(path, encoding: Encoding.ASCII);

            Assert.True(pathResult.Succeeded);
            Assert.Equal(Path.GetFullPath(path), pathResult.OutputPath);
            Assert.Equal(source.Length, source.Position);
            Assert.Contains("Try PDF", PdfCore.PdfReadDocument.Load(path).ExtractText(), StringComparison.Ordinal);
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }

        using var asyncOutput = new MemoryStream();
        PdfCore.PdfSaveResult asyncResult = await rtfBytes.TrySaveAsPdfAsync(asyncOutput);

        Assert.True(asyncResult.Succeeded);
        Assert.Equal(asyncOutput.Length, asyncResult.BytesWritten);
        Assert.Contains("Try PDF", PdfCore.PdfReadDocument.Load(asyncOutput.ToArray()).ExtractText(), StringComparison.Ordinal);

        using var readOnlyStream = new MemoryStream(Array.Empty<byte>(), writable: false);
        PdfCore.PdfSaveResult failure = rtf.TrySaveAsPdf(readOnlyStream);

        Assert.False(failure.Succeeded);
        Assert.NotEmpty(failure.Diagnostics);
    }
}
