namespace OfficeIMO.Pdf;

/// <summary>
/// PDF/UA identification metadata emitted in the generated XMP packet.
/// </summary>
/// <remarks>
/// This model writes the ISO 14289 PDF/UA identification field only. It does not
/// certify that the generated PDF satisfies PDF/UA validation requirements.
/// </remarks>
public sealed class PdfUaIdentification {
    /// <summary>PDF/UA identification namespace URI.</summary>
    public const string NamespaceUri = "http://www.aiim.org/pdfua/ns/id/";

    /// <summary>Creates PDF/UA identification metadata for a supported PDF/UA part.</summary>
    public PdfUaIdentification(int part) {
        ValidatePart(part, nameof(part));
        Part = part;
    }

    /// <summary>PDF/UA standard part, currently 1 or 2.</summary>
    public int Part { get; }

    /// <summary>Creates PDF/UA-1 identification metadata.</summary>
    public static PdfUaIdentification PdfUa1() {
        return new PdfUaIdentification(1);
    }

    /// <summary>Creates PDF/UA-2 identification metadata.</summary>
    public static PdfUaIdentification PdfUa2() {
        return new PdfUaIdentification(2);
    }

    internal PdfUaIdentification Clone() => new PdfUaIdentification(Part);

    private static void ValidatePart(int part, string paramName) {
        if (part != 1 && part != 2) {
            throw new ArgumentOutOfRangeException(paramName, "PDF/UA identification supports part 1 or 2.");
        }
    }
}
