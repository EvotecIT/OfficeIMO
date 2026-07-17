namespace OfficeIMO.OneNote;

internal static class OneNoteNotebookHierarchy {
    internal static IReadOnlyList<OneNoteNotebookHierarchyItem> Order(
        IList<OneNoteSection> sections,
        IList<OneNoteSectionGroup> groups) {
        var items = new List<OneNoteNotebookHierarchyItem>(sections.Count + groups.Count);
        int sequence = 0;
        foreach (OneNoteSection section in sections) items.Add(new OneNoteNotebookHierarchyItem(section, sequence++));
        foreach (OneNoteSectionGroup group in groups) items.Add(new OneNoteNotebookHierarchyItem(group, sequence++));
        return items
            .OrderBy(item => item.SourceOrder.HasValue ? 0 : 1)
            .ThenBy(item => item.SourceOrder ?? uint.MaxValue)
            .ThenBy(item => item.Sequence)
            .ToArray();
    }
}

internal sealed class OneNoteNotebookHierarchyItem {
    internal OneNoteNotebookHierarchyItem(OneNoteSection section, int sequence) {
        Section = section;
        SourceOrder = section.TableOfContentsOrder;
        Sequence = sequence;
    }

    internal OneNoteNotebookHierarchyItem(OneNoteSectionGroup group, int sequence) {
        Group = group;
        SourceOrder = group.TableOfContentsOrder;
        Sequence = sequence;
    }

    internal OneNoteSection? Section { get; }
    internal OneNoteSectionGroup? Group { get; }
    internal uint? SourceOrder { get; }
    internal int Sequence { get; }
}
