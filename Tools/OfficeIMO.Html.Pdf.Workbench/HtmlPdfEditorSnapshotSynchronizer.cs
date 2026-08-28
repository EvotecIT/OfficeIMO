namespace OfficeIMO.Html.Pdf.Workbench;

internal static class HtmlPdfEditorSnapshotSynchronizer {
    internal static async Task<(string Html, string Css)> ReadBothAsync(
        Func<CancellationToken, Task<string>> readHtml,
        Func<CancellationToken, Task<string>> readCss,
        CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(readHtml);
        ArgumentNullException.ThrowIfNull(readCss);
        string html = await readHtml(cancellationToken).ConfigureAwait(false);
        string css = await readCss(cancellationToken).ConfigureAwait(false);
        return (html, css);
    }
}
