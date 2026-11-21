using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;

namespace OfficeIMO.Word {
    internal static partial class WordFieldUpdater {
        /// <summary>Names of the temporary layout markers around body PAGE fields.</summary>
        internal const string PageFieldBookmarkPrefix = "_OfficeIMO_PageField_";

        /// <summary>
        /// Runs the layout with a temporary bookmark in front of every body PAGE field, so the layout reports the page
        /// each field is on the same way it reports bookmark pages for PAGEREF. When the field's runs sit directly in
        /// its paragraph, they are also wrapped in a link to that bookmark: the link is drawn on the field's own line,
        /// so a field on a later line of a paragraph that continues onto another page takes that later page. The
        /// markers are removed afterwards.
        /// </summary>
        private static WordLayoutPagination LayOutWithPageFieldBookmarks(WordDocument document, MainDocumentPart mainPart,
            Func<WordDocument, WordLayoutPagination> provider, FieldEvaluationState state) {
            var inserted = new List<OpenXmlElement>();
            var links = new List<Hyperlink>();
            try {
                Body? body = mainPart.Document?.Body;
                if (body != null) {
                    int nextId = body.Descendants<BookmarkStart>()
                        .Select(start => int.TryParse(start.Id?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int id) ? id : 0)
                        .DefaultIfEmpty(0).Max() + 1;
                    foreach (MutableFieldCandidate field in EnumerateFields(mainPart)) {
                        if (field.LocationKind != WordFieldLocationKind.Body ||
                            WordFieldInventory.ParseInstruction(field.InstructionText).FieldType != WordFieldType.Page) continue;
                        Paragraph? paragraph = field.AnchorElement.Ancestors<Paragraph>().FirstOrDefault();
                        if (paragraph == null || !paragraph.Ancestors<Body>().Any()) continue;
                        OpenXmlElement child = field.AnchorElement;
                        while (child.Parent != null && !ReferenceEquals(child.Parent, paragraph)) child = child.Parent;
                        if (child.Parent == null) continue;

                        string id = (nextId++).ToString(CultureInfo.InvariantCulture);
                        string name = PageFieldBookmarkPrefix + id;
                        var start = new BookmarkStart { Id = id, Name = name };
                        var end = new BookmarkEnd { Id = id };
                        child.InsertBeforeSelf(start);
                        start.InsertAfterSelf(end);
                        inserted.Add(start);
                        inserted.Add(end);
                        if (WrapFieldInLink(field, paragraph, child, name) is { } link) links.Add(link);
                        state.PageFieldBookmarks[field.AnchorElement] = name;
                    }
                }

                return provider(document);
            } finally {
                foreach (Hyperlink link in links) {
                    foreach (OpenXmlElement child in link.ChildElements.ToList()) {
                        child.Remove();
                        link.InsertBeforeSelf(child);
                    }
                    link.Remove();
                }
                foreach (OpenXmlElement element in inserted) element.Remove();
            }
        }

        private static Hyperlink? WrapFieldInLink(MutableFieldCandidate field, Paragraph paragraph, OpenXmlElement first, string name) {
            OpenXmlElement last = first;
            if (field.Representation == WordFieldRepresentation.Complex) {
                if (field.EndRun == null || !ReferenceEquals(field.EndRun.Parent, paragraph) || first is not Run) return null;
                last = field.EndRun;
                for (OpenXmlElement? current = first; !ReferenceEquals(current, last); current = current.NextSibling()) {
                    if (current is not Run) return null;
                }
            } else if (first is not SimpleField) {
                return null;
            }

            var link = new Hyperlink { Anchor = name };
            first.InsertBeforeSelf(link);
            OpenXmlElement? next = first;
            while (next != null) {
                OpenXmlElement? following = ReferenceEquals(next, last) ? null : next.NextSibling();
                next.Remove();
                link.Append(next);
                next = following;
            }
            return link;
        }

        private static bool TryEvaluateLayoutPage(
            WordFieldInventory.ParsedFieldInstruction parsed,
            WordLayoutPagination layout,
            string bookmarkName,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            value = null;
            status = WordFieldUpdateStatus.Skipped;
            if (!layout.TryGetBookmarkPage(bookmarkName, out int page, out string? label)) {
                message = "PAGE field could not be placed on a page by OfficeIMO layout; its result was left unchanged.";
                return false;
            }

            if (!TryFormatPageReferenceValue(page, parsed.FormatSwitches, parsed.NumericPictureSwitch, out string pageText, out string? unsupportedFormat, out _)) {
                status = WordFieldUpdateStatus.Unsupported;
                message = $"PAGE field format switch {unsupportedFormat} is not supported for deterministic page refresh.";
                return false;
            }

            if (string.IsNullOrWhiteSpace(parsed.NumericPictureSwitch) && GetLastMeaningfulFormat(parsed.FormatSwitches) == null) {
                if (label == null) {
                    status = WordFieldUpdateStatus.Unsupported;
                    message = "The section page-number format is not supported and the PAGE result was left unchanged.";
                    return false;
                }
                pageText = label;
            }

            value = pageText;
            status = WordFieldUpdateStatus.Updated;
            message = "Updated from OfficeIMO page layout.";
            return true;
        }
    }
}
