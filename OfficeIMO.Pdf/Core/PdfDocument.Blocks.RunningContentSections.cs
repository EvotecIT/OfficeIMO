namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>Preserves imported section page counts across continuous page flow.</summary>
    internal void BeginRunningContentSection(int sectionId) => AddBlock(new RunningContentSectionBlock(sectionId));

    /// <summary>Preserves continuous body flow while selecting the resolved running stories for the next physical page.</summary>
    /// <remarks>Importers must retain physical geometry and numbering, and qualify their supported source profile.</remarks>
    internal void BeginRunningContentSection(int sectionId, PdfOptions continuationOptions) {
        Guard.NotNull(continuationOptions, nameof(continuationOptions));
        AddBlock(new RunningContentSectionBlock(sectionId, continuationOptions.Clone()));
    }
}
