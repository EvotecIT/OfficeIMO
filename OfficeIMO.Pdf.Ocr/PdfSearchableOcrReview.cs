using System.Collections.ObjectModel;
using System.Threading;
using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Ocr;

/// <summary>Recognition evidence and a private source snapshot for reviewing a searchable layer before mutation.</summary>
/// <remarks>Selections may exclude eligible words. Rejected words cannot bypass confidence or native-text
/// overlap policy; change recognition options and prepare another review when those policies need to change.</remarks>
public sealed class PdfSearchableOcrReview {
    private readonly PdfDocument _source;
    private readonly PdfOcrMergeOptions _options;
    private readonly IReadOnlyDictionary<PdfRecognizedWord, int> _eligible;

    internal PdfSearchableOcrReview(PdfDocument source, PdfOcrMergeOptions options, PdfOcrMergeResult ocr) {
        _source = source;
        _options = options;
        Ocr = ocr;
        _eligible = ocr.Pages.SelectMany(page => page.Words.Select(word => new { Word = word, page.PageNumber }))
            .ToDictionary(item => item.Word, item => item.PageNumber);
    }

    /// <summary>Provider evidence and canonical merge decisions for the captured source.</summary>
    public PdfOcrMergeResult Ocr { get; }

    /// <summary>Returns the reviewed page size in the same visual points used by word geometry.</summary>
    public (double Width, double Height) GetPageSize(int pageNumber) {
        var page = Ocr.NativeDocument.Pages.SingleOrDefault(page => page.PageNumber == pageNumber)
            ?? throw new ArgumentOutOfRangeException(nameof(pageNumber), "Choose a page included in this review.");
        return page.GetVisualPageSize();
    }

    /// <summary>Renders one reviewed page from the captured source, using the recognition pixel budget.</summary>
    public PdfPageRenderResult RenderPage(int pageNumber, CancellationToken cancellationToken = default) {
        if (!Ocr.Pages.Any(page => page.PageNumber == pageNumber))
            throw new ArgumentOutOfRangeException(nameof(pageNumber), "Choose a page included in this review.");
        return PdfPageImageRenderer.RenderPages(_source.GetBytesForOperation(cancellationToken),
            PdfPageSelection.From(pageNumber), new PdfPageRenderOptions {
                Format = PdfPageRenderFormat.Png, Dpi = _options.Dpi, MaxPages = 1,
                MaxPixelsPerPage = _options.MaxPixelsPerPage, ContinueOnError = false
            }, _source.ReadOptions, cancellationToken).Single();
    }

    /// <summary>Creates a searchable artifact using all words accepted by the canonical merge policy.</summary>
    public PdfSearchableOcrResult ApplyAll(CancellationToken cancellationToken = default) =>
        Apply(Ocr.Pages.SelectMany(page => page.Words), cancellationToken);

    /// <summary>Creates a searchable artifact using only the selected eligible word instances from this review.</summary>
    /// <remarks>Foreign, rejected, null, and duplicate words are rejected before mutation. An empty selection
    /// produces an unchanged source copy. This method creates an in-memory artifact; it does not publish a file.</remarks>
    public PdfSearchableOcrResult Apply(IEnumerable<PdfRecognizedWord> selectedWords, CancellationToken cancellationToken = default) {
        Guard.NotNull(selectedWords, nameof(selectedWords));
        var selected = new HashSet<PdfRecognizedWord>();
        foreach (var word in selectedWords) {
            cancellationToken.ThrowIfCancellationRequested();
            if (word is null || !_eligible.ContainsKey(word))
                throw new ArgumentException("Select only eligible words belonging to this OCR review.", nameof(selectedWords));
            if (!selected.Add(word)) throw new ArgumentException("A word cannot be selected more than once.", nameof(selectedWords));
        }
        cancellationToken.ThrowIfCancellationRequested();
        var wordsByPage = Ocr.Pages.Select(page => new {
            page.PageNumber, Words = (IReadOnlyList<PdfRecognizedWord>)Array.AsReadOnly(page.Words.Where(selected.Contains).ToArray())
        }).Where(page => page.Words.Count > 0).ToDictionary(page => page.PageNumber, page => page.Words);
        var writtenWords = new ReadOnlyDictionary<int, IReadOnlyList<PdfRecognizedWord>>(wordsByPage);
        int[] modifiedPages = wordsByPage.Keys.ToArray();
        if (modifiedPages.Length == 0) {
            return new PdfSearchableOcrResult(PdfDocument.Load(_source.GetBytesForOperation(cancellationToken), _source.ReadOptions),
                Ocr, Array.Empty<int>(), writtenWords);
        }
        string pageSelector = string.Join(",", modifiedPages.Select(page => page.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        PdfDocument searchable = _source.Stamp.Content((canvas, context) => {
            cancellationToken.ThrowIfCancellationRequested();
            var canonicalPage = Ocr.Document.Pages.First(page => page.PageNumber == context.PageNumber);
            var logicalWords = PdfOcrLogicalDocumentBuilder.OrderWordsForLogicalReading(wordsByPage[context.PageNumber],
                canonicalPage, _options.ReadOptions.LayoutOptions.ReadingDirection, cancellationToken);
            foreach (var word in logicalWords) {
                cancellationToken.ThrowIfCancellationRequested();
                canvas.SearchableText(word.Text, word.X, word.Y, word.Width, word.Height);
            }
        }, new PdfCanvasStampOptions().UseTargetPages(pageSelector), _source.ReadOptions);
        return new PdfSearchableOcrResult(searchable, Ocr, Array.AsReadOnly(modifiedPages), writtenWords);
    }
}
