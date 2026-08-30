using HtmlTinkerX;
using Microsoft.Playwright;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal enum HtmlPdfComparisonEngine {
    OfficeIMO,
    PeachPDF,
    ITextPdfHtml,
    Chromium
}

/// <summary>
/// Owns the equivalent-work rendering calls shared by the benchmark and the
/// artifact evidence runner. Browser lifecycle and PDF capture remain owned by
/// HtmlTinkerX; this type only supplies the common comparison document.
/// </summary>
internal static class HtmlPdfComparisonRenderers {
    internal static IReadOnlyList<HtmlPdfComparisonEngine> AllEngines { get; } =
        Array.AsReadOnly(Enum.GetValues<HtmlPdfComparisonEngine>());

    internal static byte[] RenderManaged(HtmlPdfComparisonEngine engine, string html) => engine switch {
        HtmlPdfComparisonEngine.OfficeIMO => OfficeImoPdfGenerator.GenerateHtml(html),
        HtmlPdfComparisonEngine.PeachPDF => PeachPdfGenerator.Generate(html),
        HtmlPdfComparisonEngine.ITextPdfHtml => ITextPdfHtmlGenerator.Generate(html),
        HtmlPdfComparisonEngine.Chromium => throw new ArgumentException("Chromium requires an HtmlTinkerX browser session.", nameof(engine)),
        _ => throw new ArgumentOutOfRangeException(nameof(engine), engine, "Unknown HTML-to-PDF comparison engine.")
    };

    internal static Task<HtmlBrowserSession> OpenChromiumSessionAsync(CancellationToken cancellationToken = default) =>
        HtmlBrowser.OpenSessionAsync("about:blank", cancellationToken: cancellationToken);

    internal static async Task<byte[]> RenderChromiumAsync(
        HtmlBrowserSession session,
        string html,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(session);
        ArgumentNullException.ThrowIfNull(html);

        await PrepareChromiumPageAsync(session, html, cancellationToken).ConfigureAwait(false);
        return await CaptureChromiumPageAsync(session, cancellationToken).ConfigureAwait(false);
    }

    internal static async Task PrepareChromiumPageAsync(
        HtmlBrowserSession session,
        string html,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(session);
        ArgumentNullException.ThrowIfNull(html);

        await session.Page.SetContentAsync(
            html,
            new PageSetContentOptions { WaitUntil = WaitUntilState.Load }).WaitAsync(cancellationToken).ConfigureAwait(false);
    }

    internal static Task<byte[]> CaptureChromiumPageAsync(
        HtmlBrowserSession session,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(session);
        return HtmlBrowser.GetPagePdfAsync(session.Page,
#if HTMLTINKERX_SOURCE
            new HtmlBrowserPdfOptions(
                printBackground: true,
                format: PdfPageFormat.A4,
                preferCssPageSize: true,
                tagged: true),
            cancellationToken: cancellationToken
#else
            printBackground: true,
            format: PdfPageFormat.A4,
            preferCssPageSize: true,
            tagged: true,
            cancellationToken: cancellationToken
#endif
        );
    }
}
