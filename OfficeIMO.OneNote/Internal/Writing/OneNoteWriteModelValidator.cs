using System.Runtime.CompilerServices;

namespace OfficeIMO.OneNote;

/// <summary>Validates recursive public-model relationships before native serialization mutates or descends into them.</summary>
internal static class OneNoteWriteModelValidator {
    internal static void ValidateSection(
        OneNoteSection section,
        int maxPageRelationshipDepth,
        int maxContentDepth) {
        if (section == null) throw new ArgumentNullException(nameof(section));
        var state = new ValidationState(maxPageRelationshipDepth, maxContentDepth);
        foreach (OneNotePage page in section.Pages) state.ValidatePage(page, 1);
    }

    private sealed class ValidationState {
        private readonly int _maxPageRelationshipDepth;
        private readonly int _maxContentDepth;
        private readonly HashSet<OneNotePage> _activePages = new HashSet<OneNotePage>(ReferenceComparer<OneNotePage>.Instance);
        private readonly HashSet<OneNoteElement> _activeElements = new HashSet<OneNoteElement>(ReferenceComparer<OneNoteElement>.Instance);

        internal ValidationState(int maxPageRelationshipDepth, int maxContentDepth) {
            _maxPageRelationshipDepth = maxPageRelationshipDepth;
            _maxContentDepth = maxContentDepth;
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
            if (!_activePages.Add(page)) {
                throw new OneNoteFormatException(
                    "ONENOTE_WRITE_PAGE_CYCLE",
                    "Conflict and version-history page relationships must not contain cycles.");
            }

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
            if (!_activeElements.Add(element)) {
                throw new OneNoteFormatException(
                    "ONENOTE_WRITE_CONTENT_CYCLE",
                    "Outlines, paragraphs, and table cells must not contain cyclic content relationships.");
            }

            try {
                if (element is OneNoteParagraph paragraph) {
                    ValidateList(paragraph.List);
                    foreach (OneNoteElement child in paragraph.Children) ValidateElement(child, depth + 1);
                } else if (element is OneNoteOutline outline) {
                    ValidateList(outline.WrapperList);
                    foreach (OneNoteElement child in outline.Children) ValidateElement(child, depth + 1);
                } else if (element is OneNoteTable table) {
                    foreach (OneNoteTableRow row in table.Rows) {
                        foreach (OneNoteTableCell cell in row.Cells) {
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
