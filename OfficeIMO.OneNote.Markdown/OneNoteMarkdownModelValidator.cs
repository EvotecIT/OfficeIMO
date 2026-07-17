using System.Runtime.CompilerServices;

namespace OfficeIMO.OneNote.Markdown;

/// <summary>Validates caller-owned model graphs before recursive text or Markdown projection.</summary>
internal static class OneNoteMarkdownModelValidator {
    internal static void ValidateElement(OneNoteElement element, OneNoteMarkdownOptions options) {
        if (element == null) throw new ArgumentNullException(nameof(element));
        if (options == null) throw new ArgumentNullException(nameof(options));
        new ValidationState(options).ValidateElement(element, 1);
    }

    internal static void ValidateCell(OneNoteTableCell cell, OneNoteMarkdownOptions options) {
        if (cell == null) throw new ArgumentNullException(nameof(cell));
        if (options == null) throw new ArgumentNullException(nameof(options));
        var state = new ValidationState(options);
        foreach (OneNoteElement element in cell.Content) state.ValidateElement(element, 1);
    }

    internal static void ValidatePageContent(OneNotePage page, OneNoteMarkdownOptions options) {
        if (page == null) throw new ArgumentNullException(nameof(page));
        if (options == null) throw new ArgumentNullException(nameof(options));
        new ValidationState(options).ValidatePageContent(page);
    }

    internal static void ValidateSection(OneNoteSection section, OneNoteMarkdownOptions options) {
        if (section == null) throw new ArgumentNullException(nameof(section));
        if (options == null) throw new ArgumentNullException(nameof(options));
        new ValidationState(options).ValidateSection(section);
    }

    internal static void ValidateNotebook(OneNoteNotebook notebook, OneNoteMarkdownOptions options) {
        if (notebook == null) throw new ArgumentNullException(nameof(notebook));
        if (options == null) throw new ArgumentNullException(nameof(options));
        var state = new ValidationState(options);
        foreach (OneNoteSection section in notebook.Sections) state.ValidateSection(section);
        foreach (OneNoteSectionGroup group in notebook.SectionGroups) state.ValidateGroup(group, 1);
    }

    private sealed class ValidationState {
        private readonly OneNoteMarkdownOptions _options;
        private readonly HashSet<OneNoteSectionGroup> _activeGroups = new HashSet<OneNoteSectionGroup>(ReferenceComparer<OneNoteSectionGroup>.Instance);
        private readonly HashSet<OneNoteSectionGroup> _visitedGroups = new HashSet<OneNoteSectionGroup>(ReferenceComparer<OneNoteSectionGroup>.Instance);
        private readonly HashSet<OneNoteSection> _visitedSections = new HashSet<OneNoteSection>(ReferenceComparer<OneNoteSection>.Instance);
        private readonly HashSet<OneNotePage> _activePages = new HashSet<OneNotePage>(ReferenceComparer<OneNotePage>.Instance);
        private readonly HashSet<OneNotePage> _visitedPages = new HashSet<OneNotePage>(ReferenceComparer<OneNotePage>.Instance);
        private readonly HashSet<OneNoteElement> _activeElements = new HashSet<OneNoteElement>(ReferenceComparer<OneNoteElement>.Instance);
        private readonly HashSet<OneNoteElement> _visitedElements = new HashSet<OneNoteElement>(ReferenceComparer<OneNoteElement>.Instance);

        internal ValidationState(OneNoteMarkdownOptions options) {
            _options = options;
        }

        internal void ValidateSection(OneNoteSection section) {
            if (section == null) throw Error("ONENOTE_PROJECTION_NULL_SECTION", "A projected notebook hierarchy cannot contain a null section.");
            if (!_visitedSections.Add(section)) throw Error("ONENOTE_PROJECTION_SHARED_SECTION", "A OneNote section instance can appear in only one projected notebook location.");
            foreach (OneNotePage page in section.Pages) ValidatePage(page, 1);
        }

        internal void ValidateGroup(OneNoteSectionGroup group, int depth) {
            if (group == null) throw Error("ONENOTE_PROJECTION_NULL_GROUP", "A projected notebook hierarchy cannot contain a null section group.");
            if (depth > _options.MaxSectionGroupDepth) throw Error("ONENOTE_PROJECTION_GROUP_DEPTH", "The projected section-group nesting depth limit was exceeded.");
            if (_activeGroups.Contains(group)) throw Error("ONENOTE_PROJECTION_GROUP_CYCLE", "Projected section-group relationships must not contain cycles.");
            if (!_visitedGroups.Add(group)) throw Error("ONENOTE_PROJECTION_SHARED_GROUP", "A OneNote section-group instance can appear in only one projected notebook location.");

            _activeGroups.Add(group);
            try {
                foreach (OneNoteSection section in group.Sections) ValidateSection(section);
                foreach (OneNoteSectionGroup child in group.SectionGroups) ValidateGroup(child, depth + 1);
            } finally {
                _activeGroups.Remove(group);
            }
        }

        internal void ValidatePageContent(OneNotePage page) {
            if (page == null) throw new ArgumentNullException(nameof(page));
            ValidatePageElements(page);
        }

        private void ValidatePage(OneNotePage page, int depth) {
            if (page == null) throw Error("ONENOTE_PROJECTION_NULL_PAGE", "A projected OneNote page relationship cannot contain null.");
            if (depth > _options.MaxPageRelationshipDepth) throw Error("ONENOTE_PROJECTION_PAGE_DEPTH", "The projected conflict and version-history page depth limit was exceeded.");
            if (_activePages.Contains(page)) throw Error("ONENOTE_PROJECTION_PAGE_CYCLE", "Projected conflict and version-history page relationships must not contain cycles.");
            if (!_visitedPages.Add(page)) throw Error("ONENOTE_PROJECTION_SHARED_PAGE", "A OneNote page instance can appear in only one projected location.");

            _activePages.Add(page);
            try {
                ValidatePageElements(page);
                if (_options.IncludeConflictPages) {
                    foreach (OneNotePage conflict in page.ConflictPages) ValidatePage(conflict, depth + 1);
                }
                if (_options.IncludeVersionHistory) {
                    foreach (OneNotePage version in page.VersionHistory) ValidatePage(version, depth + 1);
                }
            } finally {
                _activePages.Remove(page);
            }
        }

        private void ValidatePageElements(OneNotePage page) {
            foreach (OneNoteOutline outline in page.Outlines) ValidateElement(outline, 1);
            foreach (OneNoteElement element in page.DirectContent) ValidateElement(element, 1);
        }

        internal void ValidateElement(OneNoteElement element, int depth) {
            if (element == null) throw Error("ONENOTE_PROJECTION_NULL_CONTENT", "Projected OneNote content collections cannot contain null.");
            if (depth > _options.MaxContentDepth) throw Error("ONENOTE_PROJECTION_CONTENT_DEPTH", "The projected OneNote content depth limit was exceeded.");
            if (_activeElements.Contains(element)) throw Error("ONENOTE_PROJECTION_CONTENT_CYCLE", "Projected outlines, paragraphs, and table cells must not contain cyclic content relationships.");
            if (!_visitedElements.Add(element)) throw Error("ONENOTE_PROJECTION_SHARED_CONTENT", "A OneNote content element instance can appear in only one projected location.");

            _activeElements.Add(element);
            try {
                if (element is OneNoteParagraph paragraph) {
                    foreach (OneNoteTextRun run in paragraph.Runs) {
                        if (run == null) throw Error("ONENOTE_PROJECTION_NULL_TEXT_RUN", "A projected OneNote paragraph cannot contain a null text run.");
                    }
                    foreach (OneNoteElement child in paragraph.Children) ValidateElement(child, depth + 1);
                } else if (element is OneNoteOutline outline) {
                    foreach (OneNoteElement child in outline.Children) ValidateElement(child, depth + 1);
                } else if (element is OneNoteTable table) {
                    foreach (OneNoteTableRow row in table.Rows) {
                        if (row == null) throw Error("ONENOTE_PROJECTION_NULL_TABLE_ROW", "A projected OneNote table cannot contain a null row.");
                        foreach (OneNoteTableCell cell in row.Cells) {
                            if (cell == null) throw Error("ONENOTE_PROJECTION_NULL_TABLE_CELL", "A projected OneNote table row cannot contain a null cell.");
                            foreach (OneNoteElement child in cell.Content) ValidateElement(child, depth + 1);
                        }
                    }
                }
            } finally {
                _activeElements.Remove(element);
            }
        }

        private static OneNoteFormatException Error(string code, string message) => new OneNoteFormatException(code, message);
    }

    private sealed class ReferenceComparer<T> : IEqualityComparer<T> where T : class {
        internal static readonly ReferenceComparer<T> Instance = new ReferenceComparer<T>();

        public bool Equals(T? left, T? right) => ReferenceEquals(left, right);

        public int GetHashCode(T value) => RuntimeHelpers.GetHashCode(value);
    }
}
