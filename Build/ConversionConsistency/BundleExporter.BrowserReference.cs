using System.Globalization;
using HtmlTinkerX;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.ConversionConsistency;

internal static partial class BundleExporter {
    private static async Task<BrowserReferenceArtifact> CaptureBrowserReferenceAsync(
        HtmlBrowserPdfRenderer renderer,
        ConsistencyCase contract,
        string source,
        ConsistencySuite suite,
        byte[] fontBytes,
        string output,
        CancellationToken cancellationToken) {
        if (!string.Equals(contract.Format, "html", StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("Source-browser references are supported only for HTML cases: " + contract.Id);
        if (contract.Pages.Any(page => page.TextPositions.Count > 0 && page.BrowserTextPositions.Count == 0 ||
                                       page.Regions.Count > 0 && page.BrowserRegions.Count == 0))
            throw new InvalidDataException("Browser references need explicit browser text-position and region contracts: " + contract.Id);
        if (suite.FontFamily.Any(character => !(character >= 'a' && character <= 'z')
                && !(character >= 'A' && character <= 'Z')
                && !(character >= '0' && character <= '9')
                && character is not (' ' or '-')))
            throw new InvalidDataException("Browser-reference font families may contain only ASCII letters, digits, spaces, and hyphens.");

        PageExpectation page = contract.Pages[0];
        string width = (page.Width * 96D / suite.Dpi).ToString(CultureInfo.InvariantCulture) + "px";
        string height = (page.Height * 96D / suite.Dpi).ToString(CultureInfo.InvariantCulture) + "px";
        string family = suite.FontFamily;
        string fontCss = "<style data-officeimo-qualification-font>" +
            "@font-face{font-family:'" + family + "';src:url(data:font/woff2;base64," + Convert.ToBase64String(fontBytes) + ") format('woff2');font-style:normal;font-weight:400;font-display:block}" +
            "</style>";
        string html = File.ReadAllText(source);
        int headEnd = html.IndexOf("</head>", StringComparison.OrdinalIgnoreCase);
        html = headEnd >= 0 ? html.Insert(headEnd, fontCss) : fontCss + html;
        var capture = await renderer.CaptureAsync(new HtmlBrowserPdfRequest(
            HtmlBrowserPdfSource.FromHtml(html),
            new HtmlBrowserPdfOptions(width: width, height: height, marginTop: "0", marginRight: "0", marginBottom: "0", marginLeft: "0", printBackground: true),
            readiness: new HtmlBrowserPdfReadiness(loadState: HtmlBrowserLoadState.Load, stable: true,
                stableMilliseconds: 250, timeout: 15000)), cancellationToken);
        if (capture.Diagnostics.BlockedRequestCount != 0 || capture.Diagnostics.Warnings.Count != 0)
            throw new InvalidDataException("Source-browser reference reported blocked resources or warnings: " + contract.Id);

        using var parsed = PdfPigDocument.Open(capture.PdfBytes);
        if (parsed.NumberOfPages != contract.Pages.Count)
            throw new InvalidDataException($"Source-browser reference page count for {contract.Id}: expected {contract.Pages.Count}, got {parsed.NumberOfPages}.");

        string relative = contract.Id + "/browser-reference.pdf";
        File.WriteAllBytes(ArtifactPaths.Resolve(output, relative), capture.PdfBytes);
        return new BrowserReferenceArtifact(relative, ArtifactPaths.Hash(capture.PdfBytes),
            capture.Diagnostics.BrowserVersion, parsed.NumberOfPages);
    }
}
