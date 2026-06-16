namespace OfficeIMO.Rtf;

/// <summary>
/// Dependency-free representation of an RTF embedded or linked object.
/// </summary>
public sealed class RtfObject : IRtfInline, IRtfBlock {
    /// <summary>Creates an RTF object with optional raw object payload.</summary>
    public RtfObject(RtfObjectKind kind = RtfObjectKind.Unknown, byte[]? data = null) {
        Kind = kind;
        Data = data ?? Array.Empty<byte>();
    }

    /// <summary>Object embedding or linking kind.</summary>
    public RtfObjectKind Kind { get; set; }

    /// <summary>Object class text from <c>\objclass</c> when present.</summary>
    public string? ClassName { get; set; }

    /// <summary>Object name text from <c>\objname</c> when present.</summary>
    public string? Name { get; set; }

    /// <summary>Raw bytes from the <c>\objdata</c> destination.</summary>
    public byte[] Data { get; set; }

    /// <summary>Object source width in twips or object units when declared by <c>\objw</c>.</summary>
    public int? Width { get; set; }

    /// <summary>Object source height in twips or object units when declared by <c>\objh</c>.</summary>
    public int? Height { get; set; }

    /// <summary>Horizontal scale percentage when declared by <c>\objscalex</c>.</summary>
    public int? ScaleX { get; set; }

    /// <summary>Vertical scale percentage when declared by <c>\objscaley</c>.</summary>
    public int? ScaleY { get; set; }

    /// <summary>Textual/rich fallback result shown by readers that cannot activate the object.</summary>
    public RtfParagraph Result { get; } = new RtfParagraph();

    /// <summary>Picture fallback result when the object declares a result picture.</summary>
    public RtfImage? ResultImage { get; set; }

    /// <summary>Returns the fallback result as plain text.</summary>
    public string ToPlainText() => Result.ToPlainText();
}
