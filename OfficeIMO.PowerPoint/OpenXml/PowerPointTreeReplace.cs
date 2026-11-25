using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;

namespace OfficeIMO.PowerPoint.OpenXml;

/// <summary>
/// Plans <see cref="PowerPointSlide.ReplaceAll"/> on a slide's editable XML tree, as the slide model does it: text boxes
/// and table cells among the shape tree's direct children, runs grouped while adjacent, the replacement in the run where
/// a match starts. A slide the model would read differently, or edit beyond the run text, makes the plan decline, and
/// nothing is written.
/// </summary>
internal static class PowerPointTreeReplace {
    /// <summary>The <c>a:t</c> element of a run and its new text.</summary>
    internal readonly struct Edit {
        internal Edit(OfficeXmlTree tree, int text, string value) {
            Tree = tree;
            Text = text;
            Value = value;
        }

        internal OfficeXmlTree Tree { get; }
        internal int Text { get; }
        internal string Value { get; }
    }

    internal sealed class Search {
        internal Search(string text, string? replacement, StringComparison comparison, bool wholeWords, bool includeTables) {
            Text = text;
            Replacement = replacement;
            Comparison = comparison;
            WholeWords = wholeWords;
            IncludeTables = includeTables;
        }

        internal string Text { get; }
        internal string? Replacement { get; }
        internal StringComparison Comparison { get; }
        internal bool WholeWords { get; }
        internal bool IncludeTables { get; }
    }

    /// <summary>Adds the slide's edits and returns its match count, or -1 when the model has to replace instead.</summary>
    internal static int PlanSlide(OfficeXmlTree slide, Search search, List<Edit> edits) {
        // The model keeps comments and instructions as nodes between elements, where they separate runs.
        if (slide.HasCommentsOrInstructions()) return -1;
        int root = slide.Root;
        if (!slide.IsElement(root, OfficeXmlNamespaces.PresentationML, "sld")) return -1;
        int data = slide.FindChild(root, OfficeXmlNamespaces.PresentationML, "cSld");
        int tree = data == OfficeXmlTree.NoNode ? OfficeXmlTree.NoNode : slide.FindChild(data, OfficeXmlNamespaces.PresentationML, "spTree");
        if (tree == OfficeXmlTree.NoNode) return 0;

        int count = 0;
        // Text boxes come first and tables after them, as the model visits them.
        for (int shape = slide.FirstChildOf(tree); shape != OfficeXmlTree.NoNode; shape = slide.NextSiblingOf(shape)) {
            if (!slide.IsElement(shape, OfficeXmlNamespaces.PresentationML, "sp")) continue;
            int body = slide.FindChild(shape, OfficeXmlNamespaces.PresentationML, "txBody");
            if (body == OfficeXmlTree.NoNode) continue;
            int found = PlanBody(slide, body, search, edits);
            if (found < 0) return -1;
            count += found;
        }
        if (!search.IncludeTables) return count;
        for (int frame = slide.FirstChildOf(tree); frame != OfficeXmlTree.NoNode; frame = slide.NextSiblingOf(frame)) {
            if (!slide.IsElement(frame, OfficeXmlNamespaces.PresentationML, "graphicFrame")) continue;
            int graphic = slide.FindChild(frame, OfficeXmlNamespaces.DrawingML, "graphic");
            int graphicData = graphic == OfficeXmlTree.NoNode ? OfficeXmlTree.NoNode : slide.FindChild(graphic, OfficeXmlNamespaces.DrawingML, "graphicData");
            int table = graphicData == OfficeXmlTree.NoNode ? OfficeXmlTree.NoNode : slide.FindChild(graphicData, OfficeXmlNamespaces.DrawingML, "tbl");
            if (table == OfficeXmlTree.NoNode) continue;
            int found = PlanTable(slide, table, search, edits);
            if (found < 0) return -1;
            count += found;
        }
        return count;
    }

    // Rows in order and, in each, the cells the grid has columns for; a row with fewer cells, or a table without a grid,
    // fails in the model, which then reports it.
    private static int PlanTable(OfficeXmlTree slide, int table, Search search, List<Edit> edits) {
        int grid = slide.FindChild(table, OfficeXmlNamespaces.DrawingML, "tblGrid");
        if (grid == OfficeXmlTree.NoNode) return -1;
        int columns = 0;
        for (int column = slide.FindChild(grid, OfficeXmlNamespaces.DrawingML, "gridCol"); column != OfficeXmlTree.NoNode;
             column = slide.FindNextSibling(column, OfficeXmlNamespaces.DrawingML, "gridCol")) {
            columns++;
        }
        int count = 0;
        for (int row = slide.FindChild(table, OfficeXmlNamespaces.DrawingML, "tr"); row != OfficeXmlTree.NoNode;
             row = slide.FindNextSibling(row, OfficeXmlNamespaces.DrawingML, "tr")) {
            int cell = slide.FindChild(row, OfficeXmlNamespaces.DrawingML, "tc");
            for (int column = 0; column < columns; column++) {
                if (cell == OfficeXmlTree.NoNode) return -1;
                int body = slide.FindChild(cell, OfficeXmlNamespaces.DrawingML, "txBody");
                if (body != OfficeXmlTree.NoNode) {
                    int found = PlanBody(slide, body, search, edits);
                    if (found < 0) return -1;
                    count += found;
                }
                cell = slide.FindNextSibling(cell, OfficeXmlNamespaces.DrawingML, "tc");
            }
        }
        return count;
    }

    private static int PlanBody(OfficeXmlTree slide, int body, Search search, List<Edit> edits) {
        int count = 0;
        for (int paragraph = slide.FindChild(body, OfficeXmlNamespaces.DrawingML, "p"); paragraph != OfficeXmlTree.NoNode;
             paragraph = slide.FindNextSibling(paragraph, OfficeXmlNamespaces.DrawingML, "p")) {
            int found = PlanParagraph(slide, paragraph, search, edits);
            if (found < 0) return -1;
            count += found;
        }
        return count;
    }

    private static int PlanParagraph(OfficeXmlTree slide, int paragraph, Search search, List<Edit> edits) {
        int count = 0;
        var runs = new List<int>();
        var texts = new List<string>();
        int previous = OfficeXmlTree.NoNode;
        int child = NextElement(slide, slide.FirstChildOf(paragraph), out bool mixed);
        if (mixed) return -1;
        while (child != OfficeXmlTree.NoNode) {
            if (!IsRun(slide, child)) {
                previous = child;
                child = NextElement(slide, slide.NextSiblingOf(child), out mixed);
                if (mixed) return -1;
                continue;
            }

            runs.Clear();
            texts.Clear();
            int next = child;
            while (next != OfficeXmlTree.NoNode && IsRun(slide, next)) {
                if (!TryGetRunText(slide, next, out string text)) return -1;
                runs.Add(next);
                texts.Add(text);
                next = NextElement(slide, slide.NextSiblingOf(next), out mixed);
                if (mixed) return -1;
            }

            // Field text counts for word boundaries and is never replaced.
            string? preceding = search.WholeWords ? FieldText(slide, previous) : null;
            string? following = search.WholeWords ? FieldText(slide, next) : null;
            string?[] updated = new string?[runs.Count];
            count += PowerPointRunReplace.Plan(texts, search.Text, search.Replacement, search.Comparison, search.WholeWords, preceding, following, updated);
            for (int index = 0; index < runs.Count; index++) {
                if (updated[index] is not { } value) continue;
                // A line break in the new text becomes a:br and more runs in the model; that path writes it instead.
                string written = LegacyPptXmlText.SanitizeAttributeValue(PowerPointTextContent.NormalizeBreaks(value))!;
                if (written.IndexOf('\n') >= 0) return -1;
                int text = slide.FindChild(runs[index], OfficeXmlNamespaces.DrawingML, "t");
                if (text == OfficeXmlTree.NoNode) return -1;
                edits.Add(new Edit(slide, text, written));
            }

            previous = runs[runs.Count - 1];
            child = next;
        }
        return count;
    }

    private static bool IsRun(OfficeXmlTree slide, int node) => slide.IsElement(node, OfficeXmlNamespaces.DrawingML, "r");

    // The text of the run's first a:t, as the model reads it; a:t holding elements is not text the model reads alike.
    private static bool TryGetRunText(OfficeXmlTree slide, int run, out string text) {
        text = string.Empty;
        int element = slide.FindChild(run, OfficeXmlNamespaces.DrawingML, "t");
        if (element == OfficeXmlTree.NoNode) return true;
        for (int child = slide.FirstChildOf(element); child != OfficeXmlTree.NoNode; child = slide.NextSiblingOf(child)) {
            if (slide.KindOf(child) != OfficeXmlTreeNodeKind.Text) return false;
        }
        text = slide.GetInnerText(element);
        return true;
    }

    private static string? FieldText(OfficeXmlTree slide, int node) {
        if (node == OfficeXmlTree.NoNode || !slide.IsElement(node, OfficeXmlNamespaces.DrawingML, "fld")) return null;
        int text = slide.FindChild(node, OfficeXmlNamespaces.DrawingML, "t");
        return text == OfficeXmlTree.NoNode ? null : slide.GetInnerText(text);
    }

    // The next element sibling. Whitespace between elements is not a node to the model; other text is mixed content,
    // which it reads differently, so the caller declines.
    private static int NextElement(OfficeXmlTree slide, int node, out bool mixed) {
        mixed = false;
        for (; node != OfficeXmlTree.NoNode; node = slide.NextSiblingOf(node)) {
            if (slide.KindOf(node) == OfficeXmlTreeNodeKind.Element) return node;
            if (!string.IsNullOrWhiteSpace(slide.GetText(node))) {
                mixed = true;
                return OfficeXmlTree.NoNode;
            }
        }
        return OfficeXmlTree.NoNode;
    }
}
