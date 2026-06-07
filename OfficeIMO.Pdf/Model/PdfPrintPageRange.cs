namespace OfficeIMO.Pdf;

/// <summary>
/// One-based inclusive page range used to initialize a PDF viewer print dialog.
/// </summary>
public sealed class PdfPrintPageRange {
    /// <summary>
    /// Creates a one-based inclusive print page range.
    /// </summary>
    public PdfPrintPageRange(int startPageNumber, int endPageNumber) {
        if (startPageNumber < 1) {
            throw new ArgumentOutOfRangeException(nameof(startPageNumber), "PDF print page range start page must be positive.");
        }

        if (endPageNumber < startPageNumber) {
            throw new ArgumentOutOfRangeException(nameof(endPageNumber), "PDF print page range end page must be greater than or equal to the start page.");
        }

        StartPageNumber = startPageNumber;
        EndPageNumber = endPageNumber;
    }

    /// <summary>One-based first page included in the print range.</summary>
    public int StartPageNumber { get; }

    /// <summary>One-based last page included in the print range.</summary>
    public int EndPageNumber { get; }

    internal PdfPrintPageRange Clone() => new PdfPrintPageRange(StartPageNumber, EndPageNumber);
}
