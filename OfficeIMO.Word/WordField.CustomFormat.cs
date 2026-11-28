using OfficeIMO.OpenXml.Internal;
using System.Text.RegularExpressions;

namespace OfficeIMO.Word {
    // Provides custom format operations for fields.
    public partial class WordField {
        internal string GetCustomFormat() {
            var match = Regex.Match(Field, "\\\\@ \\\"([^\\\"]+)\\\"");
            return match.Success ? match.Groups[1].Value : string.Empty;
        }

        internal void SetCustomFormat(string format) {
            string instruction = Regex.Replace(Field, "\\\\@ \\\"[^\\\"]+\\\" ?", string.Empty);
            if (!string.IsNullOrWhiteSpace(format)) {
                instruction = instruction.Replace("\\* MERGEFORMAT", $"\\@ \"{format}\" \\* MERGEFORMAT");
            }
            ResolveNodes()?.MarkChanged();
            if (_simpleField.Exists) {
                _simpleField.SetAttribute(W, "instr", instruction);
                return;
            }
            foreach (OfficeXmlNode run in _runs) {
                OfficeXmlNode code = run.Child(W, "instrText");
                if (code.IsNone) continue;
                code.SetInnerText(instruction);
                break;
            }
        }
    }
}
