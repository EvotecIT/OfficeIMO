using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed partial class Binder {
        private static IReadOnlyList<RtfRevisionAuthor> ReadRevisionAuthors(RtfGroup root, int ansiCodePage, int unicodeSkipCount) {
            RtfGroup? revisionTable = root.Children.OfType<RtfGroup>().FirstOrDefault(group => group.Destination == "revtbl");
            if (revisionTable == null) return Array.Empty<RtfRevisionAuthor>();

            var authors = new List<RtfRevisionAuthor>();
            foreach (RtfGroup authorGroup in revisionTable.Children.OfType<RtfGroup>()) {
                string author = CollectPlainText(authorGroup, ansiCodePage, unicodeSkipCount).Trim().TrimEnd(';').Trim();
                authors.Add(new RtfRevisionAuthor(author));
            }

            return authors;
        }

        private static IReadOnlyList<int> ReadRevisionSaveIds(RtfGroup root) {
            RtfGroup? revisionSaveIdTable = root.Children.OfType<RtfGroup>().FirstOrDefault(group => group.Destination == "rsidtbl");
            if (revisionSaveIdTable == null) return Array.Empty<int>();

            var ids = new List<int>();
            foreach (RtfControlWord control in revisionSaveIdTable.Children.OfType<RtfControlWord>()) {
                if (control.Name == "rsid" && control.Parameter.HasValue && control.Parameter.Value >= 0) {
                    ids.Add(control.Parameter.Value);
                }
            }

            return ids;
        }

        private static int? ReadRevisionRootSaveId(RtfGroup root) {
            RtfGroup? revisionSaveIdTable = root.Children.OfType<RtfGroup>().FirstOrDefault(group => group.Destination == "rsidtbl");
            if (revisionSaveIdTable == null) return null;

            RtfControlWord? rootId = revisionSaveIdTable.Children.OfType<RtfControlWord>().FirstOrDefault(control => control.Name == "rsidroot");
            return rootId?.Parameter >= 0 ? rootId.Parameter : null;
        }
    }
}
