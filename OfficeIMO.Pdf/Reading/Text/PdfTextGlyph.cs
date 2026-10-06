namespace OfficeIMO.Pdf;

/// <summary>Immutable evidence for one encoded PDF glyph in a positioned text span.</summary>
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
    /// <summary>Baseline advance in points, including authored spacing.</summary>
    public double Advance { get; }
    /// <summary>Glyph width in points before additional character or word spacing.</summary>
    public double PaintedAdvance { get; }
    /// <summary>Original encoded PDF character code; callers cannot mutate the reader's bytes.</summary>
    public IReadOnlyList<byte> EncodedBytes { get; }
}
