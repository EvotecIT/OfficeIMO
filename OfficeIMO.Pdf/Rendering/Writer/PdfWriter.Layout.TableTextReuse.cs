using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    // One table's preparation pass owns these plans. Rendering only reads the
    // line lists, and every retained plan has already recorded its font usage.
    // Rich runs, callbacks and diagnostic reports still take the ordinary path.
    private sealed class TableTextLayoutReuse {
        private const int MaximumEntries = 64;
        private const int MaximumTextLength = 256;
        private readonly PdfOptions options;
        private readonly Dictionary<(string Text, double Width, bool NoWrap, PdfStandardFont Font, double Size, double Leading), LinkedListNode<Entry>> entries = new();
        private readonly LinkedList<Entry> recent = new();

        public TableTextLayoutReuse(PdfOptions options) {
            this.options = options;
        }

        public TableCellTextLayout Create(TableCellLayout cell, double innerWidth, PdfStandardFont baseFont, double fontSize, double leading, double runFontSizeScale, double minimumShrinkFontSize) {
            if (!CanReuse(cell, runFontSizeScale)) {
                return CreateTableCellTextLayout(cell, innerWidth, baseFont, fontSize, leading, options, runFontSizeScale, minimumShrinkFontSize);
            }

            var key = (cell.Runs[0].Text, innerWidth, cell.NoWrap, baseFont, fontSize, leading);
            if (entries.TryGetValue(key, out LinkedListNode<Entry>? node)) {
                recent.Remove(node);
                recent.AddFirst(node);
                return node.Value.Layout;
            }

            TableCellTextLayout layout = CreateTableCellTextLayout(cell, innerWidth, baseFont, fontSize, leading, options, runFontSizeScale, minimumShrinkFontSize);
            node = recent.AddFirst(new Entry(key, layout));
            entries.Add(key, node);
            if (entries.Count > MaximumEntries) {
                LinkedListNode<Entry> oldest = recent.Last!;
                entries.Remove(oldest.Value.Key);
                recent.RemoveLast();
            }
            return layout;
        }

        private bool CanReuse(TableCellLayout cell, double runFontSizeScale) {
            if (runFontSizeScale < 0.999D || cell.Paragraphs.Count != 0 || cell.Runs.Count != 1 ||
                options.HasDiagnosticsReport || options.TextShapingProviderSnapshot != null ||
                options.TextHyphenationCallbackSnapshot != null || options.TextLineBreakCallbackSnapshot != null) {
                return false;
            }

            PdfTextRun run = cell.Runs[0];
            if (run.Bold || run.Italic || run.Underline || run.Strike || run.Color.HasValue ||
                run.BackgroundColor.HasValue || run.DecorationColor.HasValue || run.FontSize.HasValue || run.Font.HasValue ||
                run.FontFamily != null || run.LinkUri != null || run.LinkDestinationName != null || run.LinkContents != null ||
                run.Baseline != PdfTextBaseline.Normal || run.TabLeader != PdfTabLeaderStyle.None ||
                run.TabAlignment != PdfTabAlignment.Left || run.InlineElement != null || run.HorizontalOffset != 0D ||
                !run.FeatureSettings.IsDefault || run.TextDirection != OfficeTextDirection.Auto ||
                run.Text.Length == 0 || run.Text.Length > MaximumTextLength) {
                return false;
            }

            foreach (char character in run.Text) {
                if (character < ' ' || character > '~') return false;
            }
            return true;
        }

        private sealed class Entry {
            public Entry((string Text, double Width, bool NoWrap, PdfStandardFont Font, double Size, double Leading) key, TableCellTextLayout layout) {
                Key = key;
                Layout = layout;
            }

            public (string Text, double Width, bool NoWrap, PdfStandardFont Font, double Size, double Leading) Key { get; }
            public TableCellTextLayout Layout { get; }
        }
    }
}
