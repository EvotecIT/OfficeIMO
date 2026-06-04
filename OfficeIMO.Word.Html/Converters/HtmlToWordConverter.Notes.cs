using AngleSharp.Dom;
using System.Threading;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private bool TryProcessNoteAnchor(
            string anchor,
            WordSection section,
            HtmlToWordOptions options,
            ref WordParagraph? currentParagraph,
            WordTableCell? cell,
            WordHeaderFooter? headerFooter) {
            if (_footnoteMap.TryGetValue(anchor, out var fnText)) {
                currentParagraph ??= cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
                var noteRef = AddNoteReference(currentParagraph!, fnText ?? string.Empty, options, NoteReferenceType.Footnote);
                TryLinkNoteReference(noteRef, fnText ?? string.Empty, options, NoteReferenceType.Footnote);
                return true;
            }

            if (_endnoteMap.TryGetValue(anchor, out var enText)) {
                currentParagraph ??= cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
                var noteRef = AddNoteReference(currentParagraph!, enText ?? string.Empty, options, NoteReferenceType.Endnote);
                TryLinkNoteReference(noteRef, enText ?? string.Empty, options, NoteReferenceType.Endnote);
                return true;
            }

            return false;
        }

        private void CaptureNoteSections(IDocument document, CancellationToken cancellationToken = default) {
            CaptureNoteSection(document.QuerySelector("section.footnotes"), _footnoteMap, cancellationToken);
            CaptureNoteSection(document.QuerySelector("section.endnotes"), _endnoteMap, cancellationToken);
        }

        private static void CaptureNoteSection(IElement? noteSection, Dictionary<string, string> noteMap, CancellationToken cancellationToken) {
            if (noteSection == null) {
                return;
            }

            foreach (var li in noteSection.QuerySelectorAll("li")) {
                cancellationToken.ThrowIfCancellationRequested();
                var id = li.GetAttribute("id");
                if (!string.IsNullOrEmpty(id)) {
                    noteMap[id!] = li.TextContent?.Trim() ?? string.Empty;
                }
            }

            noteSection.Remove();
        }
    }
}
