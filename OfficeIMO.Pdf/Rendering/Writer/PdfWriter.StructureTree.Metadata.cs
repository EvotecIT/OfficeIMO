namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static string? StructureId(PageStructElement element) => element.StructureType switch {
        "Note" => NoteStructureId(element.ObjectId),
        "TH" => "header-" + element.ObjectId.ToString(System.Globalization.CultureInfo.InvariantCulture),
        _ => null
    };

    private static Dictionary<PageStructElement, PageStructElement?> BuildStructureParents(IReadOnlyList<LayoutResult.Page> pages) {
        var parents = new Dictionary<PageStructElement, PageStructElement?>();
        foreach (var page in pages) {
            foreach (var element in page.StructElements) {
                var parent = element.ParentElement;
                if (parent == null && element.ParentElementIndex is int index && index >= 0 && index < page.StructElements.Count)
                    parent = page.StructElements[index];
                parents[element] = parent;
            }
        }
        return parents;
    }

    private static string? InheritedStructureLanguage(PageStructElement element, IReadOnlyDictionary<PageStructElement, PageStructElement?> parents, string? documentLanguage) {
        while (parents.TryGetValue(element, out var parent) && parent != null) {
            if (parent.Language != null) return parent.Language;
            element = parent;
        }
        return documentLanguage;
    }

    private static Dictionary<PageStructElement, IReadOnlyList<string>> BuildTableHeaderAssociations(IReadOnlyList<LayoutResult.Page> pages, IReadOnlyDictionary<PageStructElement, PageStructElement?> parents) {
        var children = new Dictionary<PageStructElement, List<PageStructElement>>();
        var all = pages.SelectMany(page => page.StructElements).ToArray();
        foreach (var pair in parents) {
            if (pair.Value == null) continue;
            if (!children.TryGetValue(pair.Value, out var siblings)) children[pair.Value] = siblings = new();
            siblings.Add(pair.Key);
        }
        IEnumerable<PageStructElement> OrderedChildren(PageStructElement parent) {
            if (!children.TryGetValue(parent, out var list)) return Enumerable.Empty<PageStructElement>();
            var ordered = new List<PageStructElement>(list);
            OrderStructureChildren(ordered);
            return ordered;
        }
        IEnumerable<PageStructElement> Rows(PageStructElement parent) {
            foreach (var child in OrderedChildren(parent)) {
                if (child.StructureType == "TR") yield return child;
                else if (child.StructureType is "THead" or "TBody" or "TFoot") {
                    foreach (var row in Rows(child)) yield return row;
                }
            }
        }
        var result = new Dictionary<PageStructElement, IReadOnlyList<string>>();
        foreach (var table in all.Where(element => element.StructureType == "Table")) {
            var cells = new List<(PageStructElement Element, int Row, long Column)>();
            var activeRowSpans = new List<(PageStructElement Element, int Row, long Column)>();
            int rowIndex = 0;
            foreach (var row in Rows(table)) {
                activeRowSpans.RemoveAll(span => (long)span.Row + span.Element.TableRowSpan <= rowIndex);
                long column = 0;
                foreach (var cell in OrderedChildren(row).Where(cell => cell.StructureType is "TH" or "TD")) {
                    // LogicalOrder orders siblings; it is not a physical column index.
                    // Skip columns occupied by row-spanning cells from earlier rows.
                    {
                        bool occupied;
                        do {
                            occupied = false;
                            foreach (var prior in activeRowSpans) {
                                if (prior.Row < rowIndex && (long)prior.Row + prior.Element.TableRowSpan > rowIndex &&
                                    prior.Column < column + cell.TableColumnSpan && prior.Column + prior.Element.TableColumnSpan > column) {
                                    column = prior.Column + prior.Element.TableColumnSpan;
                                    occupied = true;
                                }
                            }
                        } while (occupied);
                    }
                    cells.Add((cell, rowIndex, column));
                    if (cell.TableRowSpan > 1) activeRowSpans.Add((cell, rowIndex, column));
                    column += cell.TableColumnSpan;
                }
                rowIndex++;
            }
            var headers = cells.Where(cell => cell.Element.StructureType == "TH").ToArray();
            foreach (var cell in cells) {
                var applicable = headers.Where(header => !ReferenceEquals(header.Element, cell.Element) && (
                    ((header.Element.TableHeaderScope is "" or "Column" or "Both") && header.Row < cell.Row &&
                        header.Column < cell.Column + cell.Element.TableColumnSpan && header.Column + header.Element.TableColumnSpan > cell.Column) ||
                    ((header.Element.TableHeaderScope is "Row" or "Both") && header.Column < cell.Column &&
                        header.Row < (long)cell.Row + cell.Element.TableRowSpan && (long)header.Row + header.Element.TableRowSpan > cell.Row)))
                    .Select(header => StructureId(header.Element)!).Distinct(StringComparer.Ordinal).ToArray();
                if (applicable.Length > 0) result[cell.Element] = applicable;
            }
        }
        return result;
    }
}
