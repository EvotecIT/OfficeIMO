using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed partial class Binder {
        private static IReadOnlyList<RtfDocumentVariable> ReadDocumentVariables(RtfGroup root, int ansiCodePage, int unicodeSkipCount) {
            var variables = new List<RtfDocumentVariable>();
            foreach (RtfGroup documentVariableGroup in root.Children.OfType<RtfGroup>().Where(group => group.Destination == "docvar")) {
                RtfGroup[] valueGroups = documentVariableGroup.Children.OfType<RtfGroup>()
                    .Where(group => group.Destination == null)
                    .Take(2)
                    .ToArray();
                if (valueGroups.Length < 2) continue;

                string name = CollectPlainText(valueGroups[0], ansiCodePage, unicodeSkipCount).Trim();
                if (string.IsNullOrEmpty(name)) continue;

                string value = CollectPlainText(valueGroups[1], ansiCodePage, unicodeSkipCount).Trim();
                variables.Add(new RtfDocumentVariable(name, value));
            }

            return variables;
        }
    }
}
