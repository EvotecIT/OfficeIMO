using System.Text;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Mhtml;
using OfficeIMO.Pdf;

internal static class PackedMhtmlContract {
    // Exercise MIME ownership and CID resource delivery through actual packages,
    // rather than letting project references hide missing packaged dependencies.
    public static void Verify(byte[] png) {
        string mime = "MIME-Version: 1.0\r\n" +
            "Content-Type: multipart/related; boundary=archive; type=\"text/html\"; start=\"<root>\"\r\n\r\n" +
            "--archive\r\nContent-Type: image/png\r\nContent-ID: <logo>\r\n" +
            "Content-Transfer-Encoding: base64\r\n\r\n" + Convert.ToBase64String(png) + "\r\n" +
            "--archive\r\nContent-Type: text/html; charset=utf-8\r\nContent-ID: <root>\r\n" +
            "Content-Location: https://example.test/archive.html\r\n\r\n" +
            "<html><body><p>Packed archive</p><img src='cid:logo' width='40' height='40'></body></html>\r\n" +
            "--archive--\r\n";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(mime));
        MhtmlDocument archive = MhtmlDocument.Load(stream);
        if (archive.RootContentId != "root" || archive.Resources.Count != 1 ||
            archive.Resources[0].ContentId != "logo")
            throw new InvalidOperationException("Packed MHTML root selection or CID decoding failed.");
        byte[] pdf = archive.ToPdfDocumentResult().ToBytes();
        if (!PdfReadDocument.Open(pdf).ExtractText().Contains("Packed archive") ||
            PdfReadDocument.Open(pdf).ExtractImages().Count(image => image.IsImageFile) != 1)
            throw new InvalidOperationException("Packed MHTML PDF lost text or its embedded image.");
    }
}
