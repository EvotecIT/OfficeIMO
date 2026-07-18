using System.Runtime.CompilerServices;

namespace OfficeIMO.OneNote;

/// <summary>Identifies one page in a flattened section or notebook traversal.</summary>
public sealed class OneNotePageReference {
    internal OneNotePageReference(OneNotePage page, OneNoteSection section, string sectionPath, int index) {
        Page = page;
        Section = section;
        SectionPath = sectionPath;
        Index = index;
    }

    /// <summary>The page at this position.</summary>
    public OneNotePage Page { get; }

    /// <summary>The section that directly owns the page.</summary>
    public OneNoteSection Section { get; }

    /// <summary>Slash-separated section-group and section display path.</summary>
    public string SectionPath { get; }

    /// <summary>Zero-based position in the flattened traversal.</summary>
    public int Index { get; }
}

/// <summary>Provides the canonical ordered traversal of current OneNote pages.</summary>
public static class OneNotePageTraversal {
    /// <summary>Enumerates a section's current pages in section order.</summary>
    public static IReadOnlyList<OneNotePageReference> Flatten(OneNoteSection section) {
        if (section == null) throw new ArgumentNullException(nameof(section));
        var results = new List<OneNotePageReference>(section.Pages.Count);
        AddSection(section, section.Name, results, new HashSet<OneNotePage>(ReferenceComparer<OneNotePage>.Instance));
        return results.AsReadOnly();
    }

    /// <summary>Enumerates a notebook's current pages in native table-of-contents order.</summary>
    public static IReadOnlyList<OneNotePageReference> Flatten(
        OneNoteNotebook notebook,
        int maxSectionGroupDepth = OneNoteNotebookReaderOptions.DefaultMaxSectionGroupDepth) {
        if (notebook == null) throw new ArgumentNullException(nameof(notebook));
        if (maxSectionGroupDepth < 0 || maxSectionGroupDepth > OneNoteWriterOptions.MaximumTraversalDepth) {
            throw new ArgumentOutOfRangeException(nameof(maxSectionGroupDepth));
        }

        var results = new List<OneNotePageReference>();
        var seenGroups = new HashSet<OneNoteSectionGroup>(ReferenceComparer<OneNoteSectionGroup>.Instance);
        var activeGroups = new HashSet<OneNoteSectionGroup>(ReferenceComparer<OneNoteSectionGroup>.Instance);
        var seenSections = new HashSet<OneNoteSection>(ReferenceComparer<OneNoteSection>.Instance);
        var seenPages = new HashSet<OneNotePage>(ReferenceComparer<OneNotePage>.Instance);
        AddHierarchy(notebook.Sections, notebook.SectionGroups, string.Empty, 0, maxSectionGroupDepth, results, seenGroups, activeGroups, seenSections, seenPages);
        return results.AsReadOnly();
    }

    private static void AddHierarchy(
        IList<OneNoteSection> sections,
        IList<OneNoteSectionGroup> groups,
        string groupPath,
        int depth,
        int maxDepth,
        List<OneNotePageReference> results,
        HashSet<OneNoteSectionGroup> seenGroups,
        HashSet<OneNoteSectionGroup> activeGroups,
        HashSet<OneNoteSection> seenSections,
        HashSet<OneNotePage> seenPages) {
        if (depth > maxDepth) {
            throw new OneNoteFormatException("ONENOTE_TRAVERSAL_GROUP_DEPTH", "The notebook section-group traversal depth limit was exceeded.");
        }

        foreach (OneNoteNotebookHierarchyItem item in OneNoteNotebookHierarchy.Order(sections, groups)) {
            if (item.Section != null) {
                if (!seenSections.Add(item.Section)) {
                    throw new OneNoteFormatException("ONENOTE_TRAVERSAL_SHARED_SECTION", "A section instance occurs more than once in the notebook hierarchy.");
                }
                string sectionPath = CombinePath(groupPath, item.Section.Name);
                AddSection(item.Section, sectionPath, results, seenPages);
                continue;
            }

            OneNoteSectionGroup group = item.Group!;
            if (activeGroups.Contains(group)) {
                throw new OneNoteFormatException("ONENOTE_TRAVERSAL_GROUP_CYCLE", "The notebook section-group hierarchy contains a cycle.");
            }
            if (!seenGroups.Add(group)) {
                throw new OneNoteFormatException("ONENOTE_TRAVERSAL_SHARED_GROUP", "A section-group instance occurs more than once in the notebook hierarchy.");
            }

            activeGroups.Add(group);
            string childPath = CombinePath(groupPath, group.Name);
            AddHierarchy(group.Sections, group.SectionGroups, childPath, depth + 1, maxDepth, results, seenGroups, activeGroups, seenSections, seenPages);
            activeGroups.Remove(group);
        }
    }

    private static void AddSection(
        OneNoteSection section,
        string sectionPath,
        List<OneNotePageReference> results,
        HashSet<OneNotePage> seenPages) {
        foreach (OneNotePage page in section.Pages) {
            if (page == null) {
                throw new OneNoteFormatException("ONENOTE_TRAVERSAL_NULL_PAGE", "A section contains a null page entry.");
            }
            if (!seenPages.Add(page)) {
                throw new OneNoteFormatException("ONENOTE_TRAVERSAL_SHARED_PAGE", "A page instance occurs more than once in the section or notebook hierarchy.");
            }
            results.Add(new OneNotePageReference(page, section, sectionPath, results.Count));
        }
    }

    private static string CombinePath(string parent, string name) {
        string component = string.IsNullOrWhiteSpace(name) ? "Untitled" : name.Trim();
        return parent.Length == 0 ? component : parent + "/" + component;
    }

    private sealed class ReferenceComparer<T> : IEqualityComparer<T> where T : class {
        internal static readonly ReferenceComparer<T> Instance = new ReferenceComparer<T>();
        public bool Equals(T? x, T? y) => ReferenceEquals(x, y);
        public int GetHashCode(T obj) => RuntimeHelpers.GetHashCode(obj);
    }
}
