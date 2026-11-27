using OfficeIMO.OpenXml.Internal;
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
        private static WordLayoutPagination LayOutWithPageFieldBookmarks(WordDocument document,
            Func<WordDocument, WordLayoutPagination> provider, FieldEvaluationState state) {
            const int W = FieldNamespace;
            var inserted = new List<OfficeXmlNode>();
            var links = new List<OfficeXmlNode>();
            try {
                OfficeXmlNode body = WordFieldInventory.BodyNode(document);
                if (body.Exists) {
                    int nextId = WordFieldInventory.EnumerateFieldContent(body, reserveInactiveCompatibilityIds: true)
                        .Where(node => node.Is(W, "bookmarkStart"))
                        .Select(start => int.TryParse(start.Attribute(W, "id"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int id) ? id : 0)
                        .DefaultIfEmpty(0).Max() + 1;
                    foreach (MutableFieldCandidate field in EnumerateFields(document)) {
                        if (field.LocationKind != WordFieldLocationKind.Body ||
                            WordFieldInventory.ParseInstruction(field.InstructionText).FieldType != WordFieldType.Page) continue;
                        OfficeXmlNode paragraph = field.AnchorElement.AncestorOrSelf(W, "p");
                        if (paragraph.IsNone || !paragraph.IsInside(body)) continue;
                        OfficeXmlNode child = field.AnchorElement;
                        while (child.Exists && child.Parent != paragraph) child = child.Parent;
                        if (child.IsNone) continue;

                        string id = (nextId++).ToString(CultureInfo.InvariantCulture);
                        string name = PageFieldBookmarkPrefix + id;
                        OfficeXmlNode start = paragraph.CreateElement(W, "bookmarkStart");
                        start.SetAttribute(W, "id", id);
                        start.SetAttribute(W, "name", name);
                        OfficeXmlNode end = paragraph.CreateElement(W, "bookmarkEnd");
                        end.SetAttribute(W, "id", id);
                        paragraph.InsertBefore(start, child);
                        paragraph.InsertBefore(end, child);
                        inserted.Add(start);
                        inserted.Add(end);
                        OfficeXmlNode link = WrapFieldInLink(field, paragraph, child, name);
                        if (link.Exists) links.Add(link);
                        state.PageFieldBookmarks[field.AnchorElement] = name;
                    }
                }

                return provider(document);
            } finally {
                foreach (OfficeXmlNode link in links) {
                    OfficeXmlNode paragraph = link.Parent;
                    while (link.FirstChild.Exists) {
                        OfficeXmlNode child = link.FirstChild;
                        child.Remove();
                        paragraph.InsertBefore(child, link);
                    }
                    link.Remove();
                }
                foreach (OfficeXmlNode element in inserted) element.Remove();
            }
        }

        private static OfficeXmlNode WrapFieldInLink(MutableFieldCandidate field, OfficeXmlNode paragraph,
            OfficeXmlNode first, string name) {
            const int W = FieldNamespace;
            OfficeXmlNode last = first;
            if (field.Representation == WordFieldRepresentation.Complex) {
                last = field.EndRun;
                if (!last.Is(W, "r") || last.Parent != paragraph) return default;
                for (OfficeXmlNode run = first; run != last; run = run.NextSibling) {
                    if (!run.Is(W, "r")) return default;
                }
            } else if (!first.Is(W, "fldSimple")) {
                return default;
            }

            OfficeXmlNode link = paragraph.CreateElement(W, "hyperlink");
            link.SetAttribute(W, "anchor", name);
            paragraph.InsertBefore(link, first);
            for (OfficeXmlNode child = first; child.Exists;) {
                OfficeXmlNode next = child == last ? default : child.NextSibling;
                child.Remove();
                link.Append(child);
                child = next;
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
