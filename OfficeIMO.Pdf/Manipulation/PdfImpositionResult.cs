namespace OfficeIMO.Pdf;

/// <summary>Source page to output sheet placement for one imposed cell.</summary>
public sealed class PdfImpositionPlacement {
    internal PdfImpositionPlacement(int sourcePageNumber, int sheetPageNumber, int row, int column, PdfPageRectangle cell) {
        SourcePageNumber = sourcePageNumber;
        SheetPageNumber = sheetPageNumber;
        Row = row;
        Column = column;
        Cell = cell;
    }

    /// <summary>One-based source page.</summary>
    public int SourcePageNumber { get; }
    /// <summary>One-based output sheet.</summary>
    public int SheetPageNumber { get; }
    /// <summary>Zero-based row from the top.</summary>
    public int Row { get; }
    /// <summary>Zero-based column from the left.</summary>
    public int Column { get; }
    /// <summary>Target cell rectangle in output PDF points.</summary>
    public PdfPageRectangle Cell { get; }
}

/// <summary>Vector imposed PDF and its explicit source-to-sheet mapping.</summary>
public sealed class PdfImpositionResult {
    private readonly byte[] _bytes;
    internal PdfImpositionResult(byte[] bytes, IReadOnlyList<PdfImpositionPlacement> placements, int removedSignatureCount, PdfImpositionSourceFeatureLoss sourceFeatureLoss) {
        _bytes = (byte[])bytes.Clone();
        Placements = Array.AsReadOnly(placements.ToArray());
        RemovedSignatureCount = removedSignatureCount;
        SourceFeatureLoss = sourceFeatureLoss;
    }

    /// <summary>Serialized imposed PDF.</summary>
    public byte[] Bytes => (byte[])_bytes.Clone();
    /// <summary>One mapping per selected source page in output sheet order.</summary>
    public IReadOnlyList<PdfImpositionPlacement> Placements { get; }
    /// <summary>Signature definitions removed from an explicit unsigned source derivative.</summary>
    public int RemovedSignatureCount { get; }
    /// <summary>Source features that the output does not retain.</summary>
    public PdfImpositionSourceFeatureLoss SourceFeatureLoss { get; }
    /// <summary>Opens the imposed PDF through the normal engine API.</summary>
    public PdfDocument ToDocument() => PdfDocument.Load(_bytes);
}
