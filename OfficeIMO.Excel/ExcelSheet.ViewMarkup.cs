using System.Text;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static string? EditWorksheetViewsMarkup(string? markup, Func<OfficeXmlNode, bool> edit) {
            const int spreadsheet = OfficeXmlNamespaces.SpreadsheetML;
            OfficeXmlTree tree = OfficeXmlTree.Create(new OfficeXmlNamespaces.Table(), spreadsheet, "worksheet");
            var worksheet = new OfficeXmlNode(tree, tree.Root);
            if (!string.IsNullOrEmpty(markup)) worksheet.Append(ParseWorksheetViews(markup!));
            if (!edit(worksheet)) return markup;

            OfficeXmlNode views = worksheet.Child(spreadsheet, "sheetViews");
            return views.Exists ? Encoding.UTF8.GetString(views.OuterXml()) : null;
        }

        internal static ExcelWorksheetViewInfo ReadWorksheetViewMarkup(string markup) =>
            ReadPrimaryView(ParseWorksheetViews(markup));

        internal static ExcelWorksheetViewInfo ReadWorksheetViewStart(OfficeXmlPullReader reader) {
            const int spreadsheet = OfficeXmlNamespaces.SpreadsheetML;
            while (reader.Read()) {
                if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                if (reader.Depth == 1 && !reader.IsElement(spreadsheet, "worksheet")) break;
                if (reader.Depth != 2) continue;
                if (reader.IsElement(spreadsheet, "sheetViews")) {
                    return ReadPrimaryView(ReadWorksheetViews(reader));
                }

                // The compatibility containers are opaque in a NoProcess load, just like extLst.
                if (!reader.IsEmptyElement) reader.SkipElement();
            }
            return new ExcelWorksheetViewInfo();
        }

        private static OfficeXmlNode ReadWorksheetViews(OfficeXmlPullReader reader) {
            const int spreadsheet = OfficeXmlNamespaces.SpreadsheetML;
            OfficeXmlTree tree = OfficeXmlTree.Create(new OfficeXmlNamespaces.Table(), spreadsheet, "sheetViews");
            var views = new OfficeXmlNode(tree, tree.Root);
            if (reader.IsEmptyElement) return views;

            while (reader.Read()) {
                if (reader.Depth == 2 && reader.NodeKind == OfficeXmlNodeKind.EndElement) break;
                if (reader.Depth != 3 || reader.NodeKind != OfficeXmlNodeKind.StartElement
                    || !reader.IsElement(spreadsheet, "sheetView")) continue;

                OfficeXmlNode view = views.Append(views.CreateElement(spreadsheet, "sheetView"));
                CopyUnqualifiedViewAttributes(reader, view);
                if (reader.IsEmptyElement) break;

                while (reader.Read()) {
                    if (reader.Depth == 3 && reader.NodeKind == OfficeXmlNodeKind.EndElement) break;
                    if (reader.Depth == 4 && reader.NodeKind == OfficeXmlNodeKind.StartElement
                        && reader.IsElement(spreadsheet, "pane")) {
                        OfficeXmlNode pane = view.Append(view.CreateElement(spreadsheet, "pane"));
                        CopyUnqualifiedViewAttributes(reader, pane);
                        break;
                    }
                }
                break;
            }
            return views;
        }

        private static void CopyUnqualifiedViewAttributes(OfficeXmlPullReader reader, OfficeXmlNode element) {
            for (int index = 0; index < reader.AttributeCount; index++) {
                if (reader.GetAttributeNamespace(index, out _) == OfficeXmlNamespaces.None) {
                    element.SetAttribute(reader.GetAttributeLocalName(index), reader.GetAttributeValue(index));
                }
            }
        }

        internal static bool TryEditWorksheetViewMarkup(
            OfficePooledBytes part, Func<OfficeXmlNode, bool> edit, out byte[]? replacement) {
            const int spreadsheet = OfficeXmlNamespaces.SpreadsheetML;
            replacement = null;
            var reader = new OfficeXmlPullReader(part.Array, part.Length, processMarkupCompatibility: false);
            if (reader.Transcoded) return false;

            while (reader.Read() && reader.NodeKind != OfficeXmlNodeKind.StartElement) { }
            if (!reader.IsElement(spreadsheet, "worksheet") || reader.IsEmptyElement) return false;

            int rootStart = reader.MarkupStart;
            int rootEnd = reader.MarkupEnd;
            string prefix = reader.ElementPrefix;
            int viewStart = -1;
            int viewEnd = -1;
            int insertAt = -1;
            while (reader.Read()) {
                if (reader.Depth == 1 && reader.NodeKind == OfficeXmlNodeKind.EndElement) {
                    if (insertAt < 0) insertAt = reader.MarkupStart;
                    break;
                }
                if (reader.Depth != 2 || reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                if (reader.IsElement(spreadsheet, "sheetViews")) {
                    viewStart = reader.MarkupStart;
                    if (!reader.IsEmptyElement) reader.SkipElement();
                    viewEnd = reader.MarkupEnd;
                    break;
                }
                if (insertAt < 0 && !reader.IsElement(spreadsheet, "sheetPr") && !reader.IsElement(spreadsheet, "dimension")) {
                    insertAt = reader.MarkupStart;
                }
                if (!reader.IsEmptyElement) reader.SkipElement();
            }

            if (viewStart < 0 && insertAt < 0) return false;
            string closingTag = prefix.Length == 0 ? "</worksheet>" : "</" + prefix + ":worksheet>";
            byte[] close = Encoding.UTF8.GetBytes(closingTag);
            using var fragment = new MemoryStream(rootEnd - rootStart + Math.Max(0, viewEnd - viewStart) + close.Length);
            fragment.Write(part.Array, rootStart, rootEnd - rootStart);
            if (viewStart >= 0) fragment.Write(part.Array, viewStart, viewEnd - viewStart);
            fragment.Write(close, 0, close.Length);
            OfficeXmlTree tree = OfficeXmlTree.Parse(fragment.GetBuffer(), (int)fragment.Length, new OfficeXmlNamespaces.Table());
            var worksheet = new OfficeXmlNode(tree, tree.Root);
            if (!edit(worksheet)) return true;
            CleanupSheetViewArtifacts(worksheet);

            OfficeXmlNode views = worksheet.Child(spreadsheet, "sheetViews");
            string newMarkup = views.Exists ? Encoding.UTF8.GetString(views.OuterXml()) : string.Empty;
            var edits = new ExcelPartEdits();
            if (viewStart >= 0) edits.Replace(viewStart, viewEnd - viewStart, newMarkup);
            else edits.Insert(insertAt, newMarkup);
            replacement = edits.Apply(part.Array, part.Length, out _);
            return true;
        }

        private static OfficeXmlNode ParseWorksheetViews(string markup) {
            byte[] xml = Encoding.UTF8.GetBytes(markup);
            OfficeXmlTree tree = OfficeXmlTree.Parse(xml, xml.Length, new OfficeXmlNamespaces.Table());
            return new OfficeXmlNode(tree, tree.Root);
        }
    }
}
