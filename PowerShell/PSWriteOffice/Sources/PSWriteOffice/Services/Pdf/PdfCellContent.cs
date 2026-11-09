using OfficeIMO.Pdf;

namespace PSWriteOffice.Services.Pdf;

/// <summary>Collects native table-cell text, image and form values during pipeline-thread composition.</summary>
internal sealed class PdfCellContent {
    internal List<PdfTextRun> Runs { get; } = new();
    internal List<PdfTableCellImage> Images { get; } = new();
    internal List<PdfTableCellCheckBox> CheckBoxes { get; } = new();
    internal List<PdfTableCellFormField> Fields { get; } = new();
}
