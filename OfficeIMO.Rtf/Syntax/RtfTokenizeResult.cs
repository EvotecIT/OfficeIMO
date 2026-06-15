using OfficeIMO.Rtf.Diagnostics;

namespace OfficeIMO.Rtf.Syntax;

/// <summary>
/// Result of tokenizing RTF input.
/// </summary>
public sealed class RtfTokenizeResult {
    internal RtfTokenizeResult(IReadOnlyList<RtfToken> tokens, IReadOnlyList<RtfDiagnostic> diagnostics) {
        Tokens = tokens;
        Diagnostics = diagnostics;
    }

    /// <summary>Token stream including an end-of-file marker.</summary>
    public IReadOnlyList<RtfToken> Tokens { get; }

    /// <summary>Recoverable and fatal tokenizer diagnostics.</summary>
    public IReadOnlyList<RtfDiagnostic> Diagnostics { get; }
}
