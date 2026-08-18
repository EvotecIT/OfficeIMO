namespace OfficeIMO.Pdf;

/// <summary>Origin of one logical text block.</summary>
public enum PdfLogicalContentSourceKind {
    /// <summary>Text decoded from native PDF text-showing operations.</summary>
    Native = 0,

    /// <summary>Text accepted from an external OCR provider through the bounded OfficeIMO.Pdf merge contract.</summary>
    Ocr = 1
}

/// <summary>Immutable bounds in top-left visual page coordinates.</summary>
public sealed class PdfLogicalVisualBounds {
    internal PdfLogicalVisualBounds(double left, double top, double right, double bottom) {
        Left = left;
        Top = top;
        Right = right;
        Bottom = bottom;
    }

    /// <summary>Left edge in visual PDF points.</summary>
    public double Left { get; }

    /// <summary>Top edge in visual PDF points.</summary>
    public double Top { get; }

    /// <summary>Right edge in visual PDF points.</summary>
    public double Right { get; }

    /// <summary>Bottom edge in visual PDF points.</summary>
    public double Bottom { get; }

    /// <summary>Visual width in PDF points.</summary>
    public double Width => Right - Left;

    /// <summary>Visual height in PDF points.</summary>
    public double Height => Bottom - Top;
}
