namespace OfficeIMO.Rtf.Syntax;

/// <summary>
/// Token kinds produced by the dependency-free RTF tokenizer.
/// </summary>
public enum RtfTokenKind {
    /// <summary>Opening group brace.</summary>
    GroupStart,
    /// <summary>Closing group brace.</summary>
    GroupEnd,
    /// <summary>Alphabetic RTF control word, optionally with a numeric parameter.</summary>
    ControlWord,
    /// <summary>Single-character RTF control symbol.</summary>
    ControlSymbol,
    /// <summary>Literal text segment.</summary>
    Text,
    /// <summary>Raw binary payload following a <c>\binN</c> control word.</summary>
    Binary,
    /// <summary>End of input marker.</summary>
    EndOfFile
}
