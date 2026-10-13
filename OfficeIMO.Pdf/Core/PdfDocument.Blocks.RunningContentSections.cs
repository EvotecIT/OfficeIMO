namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>Preserves imported section page counts across continuous page flow.</summary>
    internal void BeginRunningContentSection(int sectionId) => AddBlock(new RunningContentSectionBlock(sectionId));
}
