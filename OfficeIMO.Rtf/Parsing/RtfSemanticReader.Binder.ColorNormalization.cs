using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed partial class Binder {
        // The semantic model reserves index zero for automatic color. Some producers start
        // with an explicit color instead. Normalize that variant only in the binder's view;
        // the read result's original syntax and bytes remain untouched.
        private RtfGroup NormalizeExplicitFirstColor(RtfGroup root) {
            RtfGroup? table = root.Children.OfType<RtfGroup>().FirstOrDefault(group => group.Destination == "colortbl");
            if (table == null) return root;
            bool explicitColor = false;
            foreach (RtfNode node in table.Children) {
                if (node is RtfControlWord control && control.Name != "colortbl") explicitColor = true;
                if (node is RtfText text && text.Text.Contains(";")) break;
            }
            if (!explicitColor) return root;
            return RemapColorGroup(root, table);
        }

        private RtfGroup RemapColorGroup(RtfGroup group, RtfGroup table) {
            _limits.CheckCancellation();
            var children = new List<RtfNode>(group.Children.Count + 1);
            foreach (RtfNode node in group.Children) {
                if (node is RtfGroup child) {
                    children.Add(RemapColorGroup(child, table));
                } else if (node is RtfControlWord control && IsColorReference(control.Name)) {
                    int value = control.Parameter ?? 0;
                    int mapped = value >= 0 && value < int.MaxValue ? value + 1 : value;
                    children.Add(new RtfControlWord(control.Position, control.Name, mapped, true, control.RawText));
                } else {
                    children.Add(node);
                }
                if (ReferenceEquals(group, table) && node is RtfControlWord destination && destination.Name == "colortbl") {
                    children.Add(new RtfText(destination.Position, ";", ";"));
                }
            }
            return new RtfGroup(group.Position, children);
        }

        private static bool IsColorReference(string name) => name is
            "cf" or "cb" or "highlight" or "ulc" or "chcbpat" or "chcfpat" or
            "cbpat" or "cfpat" or "brdrcf" or "trcbpat" or "trcfpat" or
            "clcbpat" or "clcfpat" or "pncf";
    }
}
