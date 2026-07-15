namespace OfficeIMO.Pdf;

/// <summary>Immutable projection of one active indirect PDF object.</summary>
public sealed class PdfRawObjectView {
    internal PdfRawObjectView(int objectNumber, int generation, PdfRawValue value) {
        ObjectNumber = objectNumber;
        Generation = generation;
        Value = value;
    }

    /// <summary>Indirect object number.</summary>
    public int ObjectNumber { get; }
    /// <summary>Indirect object generation.</summary>
    public int Generation { get; }
    /// <summary>Bounded immutable object value.</summary>
    public PdfRawValue Value { get; }
}
