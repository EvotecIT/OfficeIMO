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
        // Checked for every measured string but changed at most MaximumEntries times, so readers use the
        // published set without locking and writers replace it with a copy.
        private volatile HashSet<string> _texts = new(StringComparer.Ordinal);
        private int _characters;

        internal bool Contains(string text) {
            if (text.Length == 0 || text.Length > MaximumTextLength) return false;
            return _texts.Contains(text);
        }

        internal void Add(string text) {
            if (text.Length == 0 || text.Length > MaximumTextLength) return;
            // Once the set is full or already holds the text, nothing can change; skip the lock.
            HashSet<string> published = _texts;
            if (published.Count >= MaximumEntries || published.Contains(text)) return;
            foreach (char character in text) if (character < 32 || character > 126) return;
            lock (_lock) {
                HashSet<string> current = _texts;
                if (current.Count >= MaximumEntries || text.Length > MaximumCharacters - _characters || current.Contains(text)) return;
                var next = new HashSet<string>(current, StringComparer.Ordinal) { text };
                _characters += text.Length;
                _texts = next;
            }
        }
    }
}
