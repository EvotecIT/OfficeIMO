namespace OfficeIMO.Rtf;

/// <summary>
/// Font entry used by an RTF document.
/// </summary>
public sealed class RtfFont {
    /// <summary>
    /// Initializes a font.
    /// </summary>
    public RtfFont(int id, string name) {
        Id = id;
        Name = string.IsNullOrWhiteSpace(name) ? "Calibri" : name.Trim();
    }

    /// <summary>RTF font identifier.</summary>
    public int Id { get; }

    /// <summary>Font display name.</summary>
    public string Name { get; }

    /// <summary>Optional RTF font family classification.</summary>
    public RtfFontFamily? Family { get; set; }

    /// <summary>Optional character set from <c>\fcharset</c>.</summary>
    public int? Charset { get; set; }

    /// <summary>Optional pitch request from <c>\fprq</c>.</summary>
    public int? Pitch { get; set; }

    /// <summary>Optional code page from <c>\cpg</c>.</summary>
    public int? CodePage { get; set; }

    /// <summary>Optional font bias from <c>\fbias</c>.</summary>
    public int? Bias { get; set; }

    /// <summary>Optional alternate font name from the <c>{\*\falt ...}</c> destination.</summary>
    public string? AlternateName { get; set; }

    /// <summary>Optional raw PANOSE classification from the <c>{\*\panose ...}</c> destination.</summary>
    public string? Panose { get; set; }

    /// <summary>Optional non-tagged font name from the <c>{\*\fname ...}</c> destination.</summary>
    public string? NonTaggedName { get; set; }

    /// <summary>Optional embedded font metadata from the <c>{\*\fontemb ...}</c> destination.</summary>
    public RtfFontEmbedding? Embedding { get; set; }
}
