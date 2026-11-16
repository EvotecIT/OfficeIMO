using System.Text;

namespace OfficeIMO.Pdf;

/// <summary>
/// Text for the CIDs of a composite font whose ToUnicode map is absent or incomplete: the Adobe character collection
/// named by the descendant's CIDSystemInfo, then the embedded font program's own Unicode cmap or glyph names.
/// </summary>
internal sealed class PdfCidTextFallback {
    private readonly string? _collectionOrdering;
    private readonly Lazy<Func<int, string?>?>? _program;

    internal PdfCidTextFallback(string? collectionOrdering, Func<Func<int, string?>?>? programFactory) {
        _collectionOrdering = collectionOrdering;
        _program = programFactory == null ? null : new Lazy<Func<int, string?>?>(programFactory);
    }

    /// <summary>Appends the text of <paramref name="cid"/> and returns its source, or null when neither source maps it.</summary>
    internal PdfTextDecodingSource? TryAppend(int cid, StringBuilder text) {
        if (cid <= 0) return null;
        if (_collectionOrdering != null && PdfPredefinedCMap.TryAppendCollectionText(_collectionOrdering, cid, text)) {
            return PdfTextDecodingSource.PredefinedCMap;
        }
        string? programText = _program?.Value?.Invoke(cid);
        if (string.IsNullOrEmpty(programText)) return null;
        text.Append(programText);
        return PdfTextDecodingSource.FontProgram;
    }
}
