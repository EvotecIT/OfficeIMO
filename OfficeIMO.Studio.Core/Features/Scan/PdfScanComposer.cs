using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;

namespace OfficeIMO.Studio.Features.Scan;

/// <summary>One scanned page: its image and the words a platform recognizer found on it.</summary>
/// <param name="Image">A JPEG or PNG file of the page.</param>
/// <param name="Words">Recognized words, each with its box from the image's top-left corner in fractions of the image's
/// width and height. Empty when the page was not recognized.</param>
internal sealed record PdfScanPage(string Image, IReadOnlyList<PdfScanWord> Words);

internal sealed record PdfScanWord(string Text, double X, double Y, double Width, double Height, double? Confidence = null);

internal sealed record PdfScanResult(int Pages, int SearchablePages, int Words);

/// <summary>
/// Makes a PDF from scanned page images, one page each at the image's size, with recognized words added as invisible
/// text over the image so the scan can be searched, selected and copied. Recognition belongs to the host, such as
/// Apple's Vision framework on device, through <see cref="HostTextRecognizer"/>.
/// </summary>
internal static class PdfScanComposer {
    internal const int MaximumPages = 500;

    internal static async Task<PdfScanResult> ComposeAsync(IReadOnlyList<PdfScanPage> pages, string outputPath, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(pages);
        if (pages.Count == 0) throw new ArgumentException("Scan at least one page.", nameof(pages));
        if (pages.Count > MaximumPages) throw new ArgumentOutOfRangeException(nameof(pages), $"A scan has at most {MaximumPages} pages.");
        if (string.IsNullOrWhiteSpace(outputPath) || !Path.IsPathFullyQualified(outputPath))
            throw new ArgumentException("Give the scan an absolute destination.", nameof(outputPath));
        var documents = new List<PdfDocument>(pages.Count);
        int searchable = 0, words = 0;
        foreach (PdfScanPage page in pages) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfImageDocumentSource image = PdfImageDocumentSource.FromFile(page.Image);
            IReadOnlyList<PdfScanWord> pageWords = HostTextRecognizer.Validate(page.Words);
            if (pageWords.Count == 0) {
                documents.Add(PdfDocument.CreateFromImages([image], null, cancellationToken));
                continue;
            }
            PdfSearchableOcrResult result = await image.MakeSearchableAsync(HostTextRecognizer.FromWords(pageWords), new PdfOcrMergeOptions { ReconstructLayout = false },
                cancellationToken).ConfigureAwait(false);
            documents.Add(result.Document);
            if (result.WasModified) searchable++;
            words += result.AddedWordCount;
        }
        PdfDocument scan = documents.Count == 1 ? documents[0] : PdfDocument.Merge(documents);
        cancellationToken.ThrowIfCancellationRequested();
        // Written beside the destination and moved over it only when complete.
        string temporary = outputPath + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try {
            await File.WriteAllBytesAsync(temporary, scan.ToBytes(), cancellationToken).ConfigureAwait(false);
            File.Move(temporary, outputPath, overwrite: true);
        } finally {
            if (File.Exists(temporary)) File.Delete(temporary);
        }
        return new PdfScanResult(pages.Count, searchable, words);
    }
}
