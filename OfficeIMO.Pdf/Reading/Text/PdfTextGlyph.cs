namespace OfficeIMO.Pdf;

/// <summary>Immutable evidence for one encoded PDF glyph in a positioned text span.</summary>
/// <remarks>Geometry uses the source page's default user-space units: one unit is /UserUnit physical points (1/72 inch each), with /UserUnit defaulting to 1. Coordinates are not automatically normalized to physical points. Explicitly named visual bounds use their documented visual coordinate system.</remarks>
public sealed class PdfTextGlyph {
    internal PdfTextGlyph(string text, int textStart, double x, double y, double advance,
        double paintedAdvance, byte[] encodedBytes) {
        Text = text;
        TextStart = textStart;
        X = x;
        Y = y;
        Advance = advance;
        PaintedAdvance = paintedAdvance;
        EncodedBytes = Array.AsReadOnly((byte[])encodedBytes.Clone());
    }

    /// <summary>Decoded text belonging to this glyph; a ligature can contain several Unicode characters.</summary>
    public string Text { get; }
    /// <summary>UTF-16 offset in the owning span's <see cref="PdfTextSpan.Text"/>.</summary>
    public int TextStart { get; }
    /// <summary>Number of UTF-16 code units belonging to this glyph.</summary>
    public int TextLength => Text.Length;
    /// <summary>Glyph origin in page user-space points.</summary>
    public double X { get; }
    /// <summary>Glyph origin in page user-space points.</summary>
    public double Y { get; }
    /// <summary>Baseline advance in page user-space units, including authored spacing.</summary>
    public double Advance { get; }
    /// <summary>Glyph width in page user-space units before additional character or word spacing.</summary>
    public double PaintedAdvance { get; }
    /// <summary>Original encoded PDF character code; callers cannot mutate the reader's bytes.</summary>
    public IReadOnlyList<byte> EncodedBytes { get; }
}
