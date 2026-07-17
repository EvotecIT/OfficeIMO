using System.Runtime.CompilerServices;

namespace OfficeIMO.OneNote;

/// <summary>Validates recursive public-model relationships before native serialization mutates or descends into them.</summary>
internal static class OneNoteWriteModelValidator {
    internal static void ValidateSection(
        OneNoteSection section,
        int maxPageRelationshipDepth,
        int maxContentDepth) {
        if (section == null) throw new ArgumentNullException(nameof(section));
        var state = new ValidationState(
            OneNoteWriterOptions.DefaultMaxSectionGroupDepth,
            maxPageRelationshipDepth,
            maxContentDepth,
            validateSectionContent: true);
        state.ValidateSection(section);
    }

    internal static void ValidateNotebook(
        OneNoteNotebook notebook,
        OneNoteWriterOptions options,
        bool validateSectionContent) {
        if (notebook == null) throw new ArgumentNullException(nameof(notebook));
        if (options == null) throw new ArgumentNullException(nameof(options));
        var state = new ValidationState(
            options.MaxSectionGroupDepth,
            options.MaxPageRelationshipDepth,
            options.MaxContentDepth,
            validateSectionContent);
        foreach (OneNoteSection section in notebook.Sections) state.ValidateSection(section);
        foreach (OneNoteSectionGroup group in notebook.SectionGroups) state.ValidateGroup(group, 1);
    }

    private sealed class ValidationState {
        private readonly int _maxSectionGroupDepth;
        private readonly int _maxPageRelationshipDepth;
        private readonly int _maxContentDepth;
        private readonly bool _validateSectionContent;
        private readonly HashSet<OneNoteSectionGroup> _activeGroups = new HashSet<OneNoteSectionGroup>(ReferenceComparer<OneNoteSectionGroup>.Instance);
        private readonly HashSet<OneNoteSectionGroup> _visitedGroups = new HashSet<OneNoteSectionGroup>(ReferenceComparer<OneNoteSectionGroup>.Instance);
        private readonly HashSet<OneNoteSection> _visitedSections = new HashSet<OneNoteSection>(ReferenceComparer<OneNoteSection>.Instance);
        private readonly HashSet<OneNotePage> _activePages = new HashSet<OneNotePage>(ReferenceComparer<OneNotePage>.Instance);
        private readonly HashSet<OneNotePage> _visitedPages = new HashSet<OneNotePage>(ReferenceComparer<OneNotePage>.Instance);
        private readonly HashSet<OneNoteElement> _activeElements = new HashSet<OneNoteElement>(ReferenceComparer<OneNoteElement>.Instance);
        private readonly HashSet<OneNoteElement> _visitedElements = new HashSet<OneNoteElement>(ReferenceComparer<OneNoteElement>.Instance);
        private readonly HashSet<OneNoteTableRow> _visitedTableRows = new HashSet<OneNoteTableRow>(ReferenceComparer<OneNoteTableRow>.Instance);
        private readonly HashSet<OneNoteTableCell> _visitedTableCells = new HashSet<OneNoteTableCell>(ReferenceComparer<OneNoteTableCell>.Instance);

        internal ValidationState(
            int maxSectionGroupDepth,
            int maxPageRelationshipDepth,
            int maxContentDepth,
            bool validateSectionContent) {
            _maxSectionGroupDepth = maxSectionGroupDepth;
            _maxPageRelationshipDepth = maxPageRelationshipDepth;
            _maxContentDepth = maxContentDepth;
            _validateSectionContent = validateSectionContent;
        }

        internal void ValidateSection(OneNoteSection section) {
            if (section == null) {
                throw new OneNoteFormatException("ONENOTE_WRITE_NULL_SECTION", "A OneNote notebook hierarchy cannot contain a null section.");
            }
            if (!_visitedSections.Add(section)) {
                throw new OneNoteFormatException("ONENOTE_WRITE_SHARED_SECTION", "A OneNote section instance can appear in only one notebook location.");
            }
            if (_validateSectionContent) {
                foreach (OneNotePage page in section.Pages) ValidatePage(page, 1);
            }
        }

        internal void ValidateGroup(OneNoteSectionGroup group, int depth) {
            if (group == null) {
                throw new OneNoteFormatException("ONENOTE_WRITE_NULL_GROUP", "A OneNote notebook hierarchy cannot contain a null section group.");
            }
            if (depth > _maxSectionGroupDepth) {
                throw new OneNoteFormatException("ONENOTE_WRITE_GROUP_DEPTH", "The section-group nesting depth limit was exceeded.");
            }
            if (_activeGroups.Contains(group)) {
                throw new OneNoteFormatException("ONENOTE_WRITE_GROUP_CYCLE", "Section-group relationships must not contain cycles.");
            }
            if (!_visitedGroups.Add(group)) {
                throw new OneNoteFormatException("ONENOTE_WRITE_SHARED_GROUP", "A section-group instance can appear in only one notebook location.");
            }

            _activeGroups.Add(group);
            try {
                foreach (OneNoteSection section in group.Sections) ValidateSection(section);
                foreach (OneNoteSectionGroup child in group.SectionGroups) ValidateGroup(child, depth + 1);
            } finally {
                _activeGroups.Remove(group);
            }
        }

        internal void ValidatePage(OneNotePage page, int depth) {
            if (page == null) {
                throw new OneNoteFormatException("ONENOTE_WRITE_NULL_PAGE", "A OneNote page relationship cannot contain null.");
            }
            if (depth > _maxPageRelationshipDepth) {
                throw new OneNoteFormatException(
                    "ONENOTE_WRITE_PAGE_DEPTH",
                    "The conflict and version-history page relationship depth limit was exceeded.");
            }
            if (_activePages.Contains(page)) {
                throw new OneNoteFormatException(
                    "ONENOTE_WRITE_PAGE_CYCLE",
                    "Conflict and version-history page relationships must not contain cycles.");
            }
            if (!_visitedPages.Add(page)) {
                throw new OneNoteFormatException(
                    "ONENOTE_WRITE_SHARED_PAGE",
                    "A OneNote page instance can appear in only one section or related-page location.");
            }

            _activePages.Add(page);
            try {
                foreach (OneNoteElement element in page.DirectContent) ValidateElement(element, 2);
                foreach (OneNoteOutline outline in page.Outlines) ValidateElement(outline, 1);
                foreach (OneNotePage conflict in page.ConflictPages) ValidatePage(conflict, depth + 1);
                foreach (OneNotePage version in page.VersionHistory) ValidatePage(version, depth + 1);
            } finally {
                _activePages.Remove(page);
            }
        }

        private void ValidateElement(OneNoteElement element, int depth) {
            if (element == null) {
                throw new OneNoteFormatException("ONENOTE_WRITE_NULL_CONTENT", "OneNote content collections cannot contain null.");
            }
            if (depth > _maxContentDepth) {
                throw new OneNoteFormatException(
                    "ONENOTE_WRITE_CONTENT_DEPTH",
                    "The recursive OneNote content depth limit was exceeded.");
            }
            if (_activeElements.Contains(element)) {
                throw new OneNoteFormatException(
                    "ONENOTE_WRITE_CONTENT_CYCLE",
                    "Outlines, paragraphs, and table cells must not contain cyclic content relationships.");
            }
            if (!_visitedElements.Add(element)) {
                throw new OneNoteFormatException(
                    "ONENOTE_WRITE_SHARED_CONTENT",
                    "A OneNote content element instance can appear in only one location.");
            }

            _activeElements.Add(element);
            try {
                if (element is OneNoteParagraph paragraph) {
                    ValidateList(paragraph.List);
                    foreach (OneNoteTextRun run in paragraph.Runs) {
                        if (run == null) {
                            throw new OneNoteFormatException("ONENOTE_WRITE_NULL_TEXT_RUN", "A OneNote paragraph cannot contain a null text run.");
                        }
                    }
                    foreach (OneNoteElement child in paragraph.Children) ValidateElement(child, depth + 1);
                } else if (element is OneNoteOutline outline) {
                    ValidateList(outline.WrapperList);
                    foreach (OneNoteElement child in outline.Children) ValidateElement(child, depth + 1);
                } else if (element is OneNoteTable table) {
                    foreach (OneNoteTableRow row in table.Rows) {
                        if (row == null) {
                            throw new OneNoteFormatException("ONENOTE_WRITE_NULL_TABLE_ROW", "A OneNote table cannot contain a null row.");
                        }
                        if (!_visitedTableRows.Add(row)) {
                            throw new OneNoteFormatException(
                                "ONENOTE_WRITE_SHARED_TABLE_ROW",
                                "A OneNote table row instance can appear in only one location.");
                        }
                        foreach (OneNoteTableCell cell in row.Cells) {
                            if (cell == null) {
                                throw new OneNoteFormatException("ONENOTE_WRITE_NULL_TABLE_CELL", "A OneNote table row cannot contain a null cell.");
                            }
                            if (!_visitedTableCells.Add(cell)) {
                                throw new OneNoteFormatException(
                                    "ONENOTE_WRITE_SHARED_TABLE_CELL",
                                    "A OneNote table cell instance can appear in only one location.");
                            }
                            foreach (OneNoteElement child in cell.Content) ValidateElement(child, depth + 1);
                        }
                    }
                }
            } finally {
                _activeElements.Remove(element);
            }
        }

        private static void ValidateList(OneNoteListInfo? list) {
            if (list != null && (list.Level < 0 || list.Level > OneNoteListInfo.MaxLevel)) {
                throw new OneNoteFormatException(
                    "ONENOTE_WRITE_LIST_LEVEL",
                    "A native OneNote list level must be from 0 through " + OneNoteListInfo.MaxLevel + ".");
            }
        }
    }

    private sealed class ReferenceComparer<T> : IEqualityComparer<T> where T : class {
        internal static readonly ReferenceComparer<T> Instance = new ReferenceComparer<T>();

        public bool Equals(T? left, T? right) => ReferenceEquals(left, right);

        public int GetHashCode(T value) => RuntimeHelpers.GetHashCode(value);
    }
}
