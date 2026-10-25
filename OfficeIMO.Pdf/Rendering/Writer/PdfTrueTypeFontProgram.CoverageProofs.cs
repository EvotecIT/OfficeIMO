namespace OfficeIMO.Pdf;

internal sealed partial class PdfTrueTypeFontProgram {
    /// <summary>Checks a prior successful scalar-only diagnostic pass against this immutable font, without recording document usage.</summary>
    internal bool HasVerifiedScalarCoverage(string text) => _scalarCoverageProofs.Contains(text);

    /// <summary>Retains only successful printable-ASCII coverage proofs; diagnostics, labels and shaped glyphs are never retained.</summary>
    internal void RememberVerifiedScalarCoverage(string text) => _scalarCoverageProofs.Add(text);

    private sealed class ScalarCoverageProofs {
        private const int MaximumEntries = 256;
        private const int MaximumTextLength = 256;
        private const int MaximumCharacters = 16_384;
        private readonly object _lock = new();
        private readonly HashSet<string> _texts = new(StringComparer.Ordinal);
        private int _characters;

        internal bool Contains(string text) {
            if (text.Length == 0 || text.Length > MaximumTextLength) return false;
            lock (_lock) return _texts.Contains(text);
        }

        internal void Add(string text) {
            if (text.Length == 0 || text.Length > MaximumTextLength) return;
            foreach (char character in text) if (character < 32 || character > 126) return;
            lock (_lock) {
                if (_texts.Count >= MaximumEntries || text.Length > MaximumCharacters - _characters) return;
                if (_texts.Add(text)) _characters += text.Length;
            }
        }
    }
}
