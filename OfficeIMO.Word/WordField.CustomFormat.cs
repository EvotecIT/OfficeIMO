using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides custom format operations for fields.
    /// </summary>
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
            if (_simpleField != null) {
                _simpleField.Instruction = instruction;
            } else if (_runs.Count > 0) {
                var fieldCode = _runs.Select(r => r.GetFirstChild<FieldCode>()).FirstOrDefault(fc => fc != null);
                if (fieldCode != null) {
                    fieldCode.Text = instruction;
                }
            }
        }
    }
}
