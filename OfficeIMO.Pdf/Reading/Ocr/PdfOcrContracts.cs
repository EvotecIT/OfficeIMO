using System.Collections.ObjectModel;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Pdf;

/// <summary>External OCR provider seam. OfficeIMO.Pdf does not ship an OCR engine.</summary>
public interface IPdfOcrProvider {
    /// <summary>Recognizes one rendered page and returns pixel-space words.</summary>
    Task<PdfOcrResponse> RecognizeAsync(PdfOcrRequest request, CancellationToken cancellationToken = default);
}

/// <summary>Rendered page supplied to an external OCR provider.</summary>
public sealed class PdfOcrRequest {
    private readonly byte[] _png;

    internal PdfOcrRequest(int pageNumber, byte[] png, int pixelWidth, int pixelHeight, double pageWidth, double pageHeight, double scale) {
        PageNumber = pageNumber; _png = (byte[])png.Clone(); PixelWidth = pixelWidth; PixelHeight = pixelHeight; PageWidth = pageWidth; PageHeight = pageHeight; Scale = scale;
    }
    /// <summary>One-based page number.</summary>
    public int PageNumber { get; }
    /// <summary>Rendered PNG bytes.</summary>
    public byte[] Png => (byte[])_png.Clone();
    /// <summary>Rendered pixel width.</summary>
    public int PixelWidth { get; }
    /// <summary>Rendered pixel height.</summary>
    public int PixelHeight { get; }
    /// <summary>Page width in PDF points.</summary>
    public double PageWidth { get; }
    /// <summary>Page height in PDF points.</summary>
    public double PageHeight { get; }
    /// <summary>Pixels per PDF point.</summary>
    public double Scale { get; }
}

/// <summary>OCR provider response for one page.</summary>
public sealed class PdfOcrResponse {
    private const int AbsoluteMaximumWords = 1_000_000;
    private const int AbsoluteMaximumDiagnostics = 100_000;
    private const int AbsoluteMaximumTextCharacters = 16 * 1024 * 1024;
    /// <summary>Creates a provider response.</summary>
    public PdfOcrResponse(IEnumerable<PdfOcrWord> words, IEnumerable<string>? diagnostics = null) {
        Guard.NotNull(words as object, nameof(words));
        Words = MaterializeWords(words);
        Diagnostics = MaterializeDiagnostics(diagnostics);
    }
    /// <summary>Recognized words in pixel coordinates from the top-left.</summary>
    public IReadOnlyList<PdfOcrWord> Words { get; }
    /// <summary>Provider diagnostics retained in the merge report.</summary>
    public IReadOnlyList<string> Diagnostics { get; }

    private static ReadOnlyCollection<PdfOcrWord> MaterializeWords(IEnumerable<PdfOcrWord> words) {
        var result = new List<PdfOcrWord>();
        long characters = 0;
        foreach (PdfOcrWord word in words) {
            if (word == null) throw new ArgumentException("OCR words cannot contain null entries.", nameof(words));
            if (result.Count >= AbsoluteMaximumWords) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.OcrArtifacts, AbsoluteMaximumWords, result.Count + 1L);
            }
            characters = checked(characters + word.Text.Length);
            if (characters > AbsoluteMaximumTextCharacters) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.OcrArtifacts, AbsoluteMaximumTextCharacters, characters);
            }
            result.Add(word);
        }
        return result.AsReadOnly();
    }

    private static IReadOnlyList<string> MaterializeDiagnostics(IEnumerable<string>? diagnostics) {
        if (diagnostics == null) return Array.Empty<string>();
        var result = new List<string>();
        long characters = 0;
        foreach (string diagnostic in diagnostics) {
            if (result.Count >= AbsoluteMaximumDiagnostics) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.OcrArtifacts, AbsoluteMaximumDiagnostics, result.Count + 1L);
            }
            characters = checked(characters + (diagnostic?.Length ?? 0));
            if (characters > AbsoluteMaximumTextCharacters) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.OcrArtifacts, AbsoluteMaximumTextCharacters, characters);
            }
            result.Add(diagnostic ?? string.Empty);
        }
        return result.AsReadOnly();
    }
}

/// <summary>One OCR word in rendered pixel coordinates.</summary>
public sealed class PdfOcrWord {
    /// <summary>Creates a recognized pixel-space word.</summary>
    public PdfOcrWord(string text, double x, double y, double width, double height, double confidence) {
        Guard.NotNullOrWhiteSpace(text, nameof(text));
        Text = text; X = x; Y = y; Width = width; Height = height; Confidence = confidence;
    }
    /// <summary>Recognized text.</summary>
    public string Text { get; }
    /// <summary>Left pixel.</summary>
    public double X { get; }
    /// <summary>Top pixel.</summary>
    public double Y { get; }
    /// <summary>Pixel width.</summary>
    public double Width { get; }
    /// <summary>Pixel height.</summary>
    public double Height { get; }
    /// <summary>Provider confidence from 0 through 1.</summary>
    public double Confidence { get; }
}
