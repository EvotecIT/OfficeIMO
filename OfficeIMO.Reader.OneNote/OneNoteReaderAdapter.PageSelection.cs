using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Markdown;

namespace OfficeIMO.Reader.OneNote;

internal static partial class OneNoteReaderAdapter {
    private static PageSelection SelectPages(OneNoteSection source, ReaderOneNoteOptions options) {
        var projection = new OneNoteSection {
            Name = source.Name,
            SourcePath = source.SourcePath,
            Id = source.Id,
            ColorArgb = source.ColorArgb,
            TableOfContentsOrder = source.TableOfContentsOrder
        };
        foreach (OneNoteDiagnostic diagnostic in source.Diagnostics) projection.Diagnostics.Add(diagnostic);
        foreach (OneNoteRevision revision in source.Revisions) projection.Revisions.Add(revision);
        foreach (OneNoteOpaqueObject item in source.UnknownObjects) projection.UnknownObjects.Add(item);

        string[] currentHierarchy = BuildPageHierarchy(source);
        var hierarchy = new List<string>();
        var visited = new HashSet<OneNotePage>();
        for (int index = 0; index < source.Pages.Count; index++) {
            OneNotePage page = source.Pages[index];
            AddProjectedPage(projection, hierarchy, visited, page, currentHierarchy[index], options);
        }
        return new PageSelection(projection, hierarchy.ToArray());
    }

    private static void AddProjectedPage(
        OneNoteSection projection,
        ICollection<string> hierarchy,
        ISet<OneNotePage> visited,
        OneNotePage page,
        string pageHierarchy,
        ReaderOneNoteOptions options) {
        if (!visited.Add(page)) return;
        projection.Pages.Add(page);
        hierarchy.Add(pageHierarchy);

        if (options.IncludeConflictPages) {
            foreach (OneNotePage conflict in page.ConflictPages) {
                AddRelatedPage(projection, hierarchy, visited, conflict, pageHierarchy, "Conflict", options);
            }
        }
        if (options.IncludeVersionHistory) {
            foreach (OneNotePage version in page.VersionHistory) {
                AddRelatedPage(projection, hierarchy, visited, version, pageHierarchy, "Version", options);
            }
        }
    }

    private static void AddRelatedPage(
        OneNoteSection projection,
        ICollection<string> hierarchy,
        ISet<OneNotePage> visited,
        OneNotePage page,
        string parentHierarchy,
        string relation,
        ReaderOneNoteOptions options) {
        string title = string.IsNullOrWhiteSpace(page.Title) ? "Untitled page" : OneNoteTextProjection.Normalize(page.Title);
        AddProjectedPage(projection, hierarchy, visited, page, parentHierarchy + " > " + relation + ": " + title, options);
    }

    private sealed class PageSelection {
        internal PageSelection(OneNoteSection section, string[] hierarchy) {
            Section = section;
            Hierarchy = hierarchy;
        }

        internal OneNoteSection Section { get; }
        internal string[] Hierarchy { get; }
    }
}
