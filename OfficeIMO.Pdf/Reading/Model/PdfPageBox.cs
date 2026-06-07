namespace OfficeIMO.Pdf;

/// <summary>
/// Rectangular page boundary box read from a PDF page dictionary.
/// </summary>
public sealed class PdfPageBox {
    internal PdfPageBox(string name, double left, double bottom, double right, double top) {
        Name = name;
        Left = left;
        Bottom = bottom;
        Right = right;
        Top = top;
    }

    /// <summary>PDF boundary box name, for example MediaBox or TrimBox.</summary>
    public string Name { get; }

    /// <summary>Left coordinate in PDF default user-space units.</summary>
    public double Left { get; }

    /// <summary>Bottom coordinate in PDF default user-space units.</summary>
    public double Bottom { get; }

    /// <summary>Right coordinate in PDF default user-space units.</summary>
    public double Right { get; }

    /// <summary>Top coordinate in PDF default user-space units.</summary>
    public double Top { get; }

    /// <summary>Box width in PDF default user-space units.</summary>
    public double Width => Right - Left;

    /// <summary>Box height in PDF default user-space units.</summary>
    public double Height => Top - Bottom;
}
