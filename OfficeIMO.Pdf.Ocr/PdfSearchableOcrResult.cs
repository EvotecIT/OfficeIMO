using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Ocr;

/// <summary>Searchable PDF artifact together with the OCR evidence used to create its invisible text layer.</summary>
public sealed class PdfSearchableOcrResult {
    internal PdfSearchableOcrResult(PdfDocument document, PdfOcrMergeResult ocr, IReadOnlyList<int> modifiedPages,
        IReadOnlyDictionary<int, IReadOnlyList<PdfRecognizedWord>> writtenWords) {
        Document = document;
        Ocr = ocr;
        ModifiedPages = modifiedPages;
        WrittenWords = writtenWords;
    }

    /// <summary>The original or rewritten PDF document containing accepted OCR text.</summary>
    public PdfDocument Document { get; }

    /// <summary>Recognition, filtering, geometry, and provider evidence before any user exclusions.</summary>
    public PdfOcrMergeResult Ocr { get; }

    /// <summary>One-based page numbers that received invisible searchable text.</summary>
    public IReadOnlyList<int> ModifiedPages { get; }

    /// <summary>True when at least one page received invisible searchable text.</summary>
    public bool WasModified => ModifiedPages.Count > 0;

    /// <summary>Number of accepted OCR words written to the searchable text layer.</summary>
    public int AddedWordCount => WrittenWords.Values.Sum(words => words.Count);

    /// <summary>Words actually written, grouped by one-based page number after review exclusions.</summary>
    public IReadOnlyDictionary<int, IReadOnlyList<PdfRecognizedWord>> WrittenWords { get; }
}
