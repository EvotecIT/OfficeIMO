using OfficeIMO.Pdf;

namespace PSWriteOffice.Services.Pdf;

/// <summary>Connects PowerShell child text and live page tokens to native running text builders.</summary>
internal sealed class PdfRunningTextContent {
    private readonly List<(PdfTextRun Run, int Token)> _segments = new();
    internal void Run(PdfTextRun run) => _segments.Add((run, 0));
    internal void CurrentPage(PdfTextRun run) => _segments.Add((run, 1));
    internal void TotalPages(PdfTextRun run) => _segments.Add((run, 2));
    internal void Apply(PdfHeaderTextBuilder text) {
        foreach (var segment in _segments) {
            if (segment.Token == 1) text.CurrentPage(segment.Run);
            else if (segment.Token == 2) text.TotalPages(segment.Run);
            else text.Run(segment.Run);
        }
    }
    internal void Apply(PdfFooterTextBuilder text) {
        foreach (var segment in _segments) {
            if (segment.Token == 1) text.CurrentPage(segment.Run);
            else if (segment.Token == 2) text.TotalPages(segment.Run);
            else text.Run(segment.Run);
        }
    }
}
