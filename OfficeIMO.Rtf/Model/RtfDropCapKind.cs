namespace OfficeIMO.Rtf;

/// <summary>
/// Drop-cap placement for an RTF paragraph.
/// </summary>
public enum RtfDropCapKind {
    /// <summary>Drop cap is placed in text, represented by <c>\dropcapt1</c>.</summary>
    InText,

    /// <summary>Drop cap is placed in the margin, represented by <c>\dropcapt2</c>.</summary>
    Margin
}
