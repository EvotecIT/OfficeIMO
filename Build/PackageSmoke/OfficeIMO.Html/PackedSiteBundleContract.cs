using System.IO.Compression;
using System.Text;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;

internal static class PackedSiteBundleContract {
    public static void Verify(byte[] png) {
        using var stream = new MemoryStream();
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
            Add(zip, "./index.html", Encoding.UTF8.GetBytes(
                "<link rel='stylesheet' href='styles/site.css'><p>Packed site</p>"
                + "<img src='image.png?v=1' width='24' height='24'>"));
            Add(zip, "styles/site.css", Encoding.UTF8.GetBytes("p{color:#123456}"));
            Add(zip, "image.png", png);
        }
        stream.Position = 7;
        HtmlSiteBundle bundle = HtmlSiteBundle.Load(stream);
        if (stream.Position != 7 || !stream.CanRead || bundle.EntryPath != "index.html")
            throw new InvalidOperationException("Packed site-bundle input ownership failed.");

        HtmlRenderRequest svgRequest = bundle.CreateRenderRequest(HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Svg));
        foreach (HtmlRenderResult result in new[] {
            HtmlRenderEngine.Execute(bundle.HtmlDocument, svgRequest),
            HtmlRenderEngine.ExecuteAsync(bundle.HtmlDocument, svgRequest).GetAwaiter().GetResult()
        }) {
            string svg = Encoding.UTF8.GetString(result.ExportImage().Bytes);
            if (!svg.Contains("<image") || !svg.Contains("#123456"))
                throw new InvalidOperationException("Packed site-bundle SVG lost archive images or CSS.");
        }

        HtmlRenderRequest pdfRequest = HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged,
            HtmlRenderEncoder.Pdf, new HtmlToPdfOptions {
                ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
            });
        foreach (HtmlPdfRenderRequestResult result in new[] {
            bundle.RenderToPdfResult(pdfRequest),
            bundle.RenderToPdfResultAsync(pdfRequest).GetAwaiter().GetResult()
        }) {
            PdfReadDocument pdf = PdfReadDocument.Open(result.ToBytes());
            if (!pdf.ExtractText().Contains("Packed site") || !pdf.ExtractImages().Any(image => image.IsImageFile))
                throw new InvalidOperationException("Packed site-bundle PDF lost text or archive images.");
        }
    }

    private static void Add(ZipArchive archive, string path, byte[] bytes) {
        using Stream payload = archive.CreateEntry(path, CompressionLevel.Optimal).Open();
        payload.Write(bytes, 0, bytes.Length);
    }
}
