namespace OfficeIMO.Pdf;

/// <summary>Marks a logical source section without interrupting physical page flow.</summary>
internal sealed class RunningContentSectionBlock : IPdfBlock {
    internal RunningContentSectionBlock(int sectionId) => SectionId = sectionId;
    internal int SectionId { get; }
}
