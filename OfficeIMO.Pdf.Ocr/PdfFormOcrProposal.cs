using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Ocr;

/// <summary>A possible value for one existing named field, awaiting human review.</summary>
/// <remarks>Recognition never creates a field or accepts a proposal. Word geometry is in original-page visual points.</remarks>
public sealed class PdfFormOcrProposal {
    internal PdfFormOcrProposal(PdfFormField field, string value, IReadOnlyList<PdfFormOcrEvidence> evidence,
        bool ambiguous, string? rejection) {
        Field = field; SuggestedValue = value; Evidence = evidence; IsAmbiguous = ambiguous; RejectionReason = rejection;
    }
    /// <summary>Existing field and declared constraints from the captured document.</summary>
    public PdfFormField Field { get; }
    /// <summary>OCR text inside visible widget bounds. A correction may be required.</summary>
    public string SuggestedValue { get; }
    /// <summary>Original recognized words and widget locations, including low confidence evidence.</summary>
    public IReadOnlyList<PdfFormOcrEvidence> Evidence { get; }
    /// <summary>Different widgets or overlapping fields produced competing assignments.</summary>
    public bool IsAmbiguous { get; }
    /// <summary>True when any contributing word fell below the recognition threshold.</summary>
    public bool HasLowConfidence => Evidence.Any(item => item.Word.Disposition == PdfOcrWordDisposition.LowConfidence);
    /// <summary>Lowest contributing provider confidence; zero for a field without recognized content.</summary>
    public double Confidence => Evidence.Count == 0 ? 0 : Evidence.Min(item => item.Word.Word.Confidence);
    /// <summary>Why this field cannot be filled through the OCR review; null for supported fields.</summary>
    public string? RejectionReason { get; }
    /// <summary>Whether the field supports reviewed text/choice values without unqualified script constraints.</summary>
    public bool CanAccept => RejectionReason is null;
}

/// <summary>One original OCR word matched to a visible existing widget.</summary>
public sealed class PdfFormOcrEvidence {
    internal PdfFormOcrEvidence(int pageNumber, PdfSelectionQuad widgetBounds, PdfOcrWordEvidence word) {
        PageNumber = pageNumber; WidgetBounds = widgetBounds; Word = word;
    }
    /// <summary>One-based original source page.</summary>
    public int PageNumber { get; }
    /// <summary>Existing widget rectangle mapped through crop, rotation and UserUnit.</summary>
    public PdfSelectionQuad WidgetBounds { get; }
    /// <summary>Unmodified provider word and recognition disposition.</summary>
    public PdfOcrWordEvidence Word { get; }
}
