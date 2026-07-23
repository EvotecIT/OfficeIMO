namespace OfficeIMO.Rtf;

/// <summary>
/// Enforces the semantic table nesting boundary shared by RTF conversion surfaces.
/// </summary>
internal static class RtfTableTraversalGuard {
    internal const int MaximumDepth = 64;

    internal static void ValidateDocument(RtfDocument document) {
        if (document == null) throw new ArgumentNullException(nameof(document));

        var pending = new Stack<TableFrame>();
        PushTables(document.Blocks, 1, pending);
        foreach (RtfSection section in document.Sections) {
            PushTables(section.Blocks, 1, pending);
        }

        while (pending.Count > 0) {
            TableFrame frame = pending.Pop();
            EnsureDepth(frame.Depth);
            foreach (RtfTableRow row in frame.Table.Rows) {
                foreach (RtfTableCell cell in row.Cells) {
                    PushTables(cell.Blocks, frame.Depth + 1, pending);
                }
            }
        }
    }

    internal static void EnsureDepth(int depth) {
        if (depth <= MaximumDepth) return;

        throw new InvalidDataException(
            $"RTF table nesting depth exceeded the supported maximum of {MaximumDepth}.");
    }

    private static void PushTables(IReadOnlyList<IRtfBlock> blocks, int depth, Stack<TableFrame> pending) {
        for (int index = blocks.Count - 1; index >= 0; index--) {
            if (blocks[index] is RtfTable table) {
                pending.Push(new TableFrame(table, depth));
            }
        }
    }

    private readonly struct TableFrame {
        internal TableFrame(RtfTable table, int depth) {
            Table = table;
            Depth = depth;
        }

        internal RtfTable Table { get; }

        internal int Depth { get; }
    }
}
