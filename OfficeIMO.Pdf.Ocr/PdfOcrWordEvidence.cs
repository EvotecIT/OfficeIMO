namespace OfficeIMO.Pdf.Ocr;

/// <summary>The canonical merge decision for a recognized word with valid page geometry.</summary>
public enum PdfOcrWordDisposition {
    /// <summary>The word passed confidence and native-text overlap filtering.</summary>
    Accepted,
    /// <summary>The word fell below the confidence threshold; native overlap was not evaluated.</summary>
    LowConfidence,
    /// <summary>Existing native PDF text already covers the word's region.</summary>
    NativeTextOverlap
}

/// <summary>Recognized word geometry and the decision that admitted or rejected it.</summary>
/// <remarks>Invalid geometry and invalid confidence are reported through page diagnostics instead.</remarks>
public sealed class PdfOcrWordEvidence {
    internal PdfOcrWordEvidence(PdfRecognizedWord word, PdfOcrWordDisposition disposition) {
        Word = word;
        Disposition = disposition;
    }

    /// <summary>Normalized text, confidence, hierarchy, and top-left visual page geometry.</summary>
    public PdfRecognizedWord Word { get; }

    /// <summary>Engine merge decision; rejected words do not enter the searchable layer.</summary>
    public PdfOcrWordDisposition Disposition { get; }
}
