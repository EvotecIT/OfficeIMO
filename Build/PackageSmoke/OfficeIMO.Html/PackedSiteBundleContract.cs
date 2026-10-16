using System.IO.Compression;
using System.Text;
using OfficeIMO.Drawing;
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
            HtmlRenderPage page = result.Document.Pages[0];
            HtmlRenderImage image = page.Visuals.OfType<HtmlRenderImage>().Single();
            if (!image.Bytes.SequenceEqual(png) || image.Width != 24D || image.Height != 24D
                || !page.Visuals.OfType<HtmlRenderText>().Any(text => text.Text == "Packed site" && text.Color == OfficeColor.FromRgb(0x12, 0x34, 0x56)))
                throw new InvalidOperationException("Packed site-bundle rendering lost archive image bytes or stylesheet text color.");
            if (!OfficeSvgDrawingReader.TryRead(result.ExportImage().Bytes, out OfficeDrawing? drawing, out _)
                || !HasImagePaint(drawing!, page, image))
                throw new InvalidOperationException("Packed site-bundle SVG lost visible archive-image paint.");
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

    private static bool HasImagePaint(OfficeDrawing drawing, HtmlRenderPage page, HtmlRenderImage image) {
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing, 2D, OfficeColor.White);
        double scale = raster.Width / page.Width;
        int left = Math.Max(0, (int)Math.Floor(image.X * scale));
        int top = Math.Max(0, (int)Math.Floor(image.Y * scale));
        int right = Math.Min(raster.Width, (int)Math.Ceiling((image.X + image.Width) * scale));
        int bottom = Math.Min(raster.Height, (int)Math.Ceiling((image.Y + image.Height) * scale));
        for (int y = top; y < bottom; y++)
            for (int x = left; x < right; x++) {
                OfficeColor pixel = raster.GetPixel(x, y);
                if (pixel.A > 128 && pixel.R < 250 && pixel.G < 250 && pixel.B < 250) return true;
            }
        return false;
    }

    private static void Add(ZipArchive archive, string path, byte[] bytes) {
        using Stream payload = archive.CreateEntry(path, CompressionLevel.Optimal).Open();
        payload.Write(bytes, 0, bytes.Length);
    }
}
