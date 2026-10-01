using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed partial class Binder {
        private static IReadOnlyList<RtfDocumentVariable> ReadDocumentVariables(RtfGroup root, int ansiCodePage, int unicodeSkipCount) {
            var variables = new List<RtfDocumentVariable>();
            foreach (RtfGroup documentVariableGroup in root.Children.OfType<RtfGroup>().Where(group => group.Destination == "docvar")) {
                RtfGroup[] valueGroups = documentVariableGroup.Children.OfType<RtfGroup>()
                    .Take(2)
                    .ToArray();
                if (valueGroups.Length < 2) continue;

                int scopedCount = GetUnicodeSkipCountBefore(root, documentVariableGroup);
                string name = CollectPlainText(valueGroups[0], ansiCodePage,
                    GetUnicodeSkipCountBefore(documentVariableGroup, valueGroups[0], scopedCount)).Trim();
                if (string.IsNullOrEmpty(name)) continue;

                string value = CollectPlainText(valueGroups[1], ansiCodePage,
                    GetUnicodeSkipCountBefore(documentVariableGroup, valueGroups[1], scopedCount)).Trim();
                variables.Add(new RtfDocumentVariable(name, value));
            }

            return variables;
        }
    }
}
