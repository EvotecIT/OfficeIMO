namespace OfficeIMO.Rtf;

/// <summary>
/// Numeric fields in the RTF document information destination supported by the lossless editor.
/// </summary>
public enum RtfDocumentInfoNumberField {
    /// <summary>Total editing time in minutes.</summary>
    EditingMinutes,

    /// <summary>Document page count.</summary>
    NumberOfPages,

    /// <summary>Document word count.</summary>
    NumberOfWords,

    /// <summary>Document character count.</summary>
    NumberOfCharacters,

    /// <summary>Document character count including whitespace.</summary>
    NumberOfCharactersWithSpaces,

    /// <summary>Internal RTF version number.</summary>
    InternalVersion
}
