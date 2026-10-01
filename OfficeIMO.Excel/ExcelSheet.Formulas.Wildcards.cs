using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static string CreateFormulaWildcardPattern(string value) {
            var pattern = new StringBuilder();
            for (int index = 0; index < value.Length; index++) {
                char current = value[index];
                if (current == '~' && index + 1 < value.Length
                    && (value[index + 1] == '*' || value[index + 1] == '?' || value[index + 1] == '~')) {
                    pattern.Append(Regex.Escape(value[++index].ToString()));
                } else if (current == '*') pattern.Append(".*");
                else if (current == '?') pattern.Append('.');
                else pattern.Append(Regex.Escape(current.ToString()));
            }
            return pattern.ToString();
        }

        private static int FindFormulaWildcard(string pattern, string text, int start) {
            var expression = new Regex(CreateFormulaWildcardPattern(pattern),
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Singleline, FormulaRegexTimeout);
            Match match = expression.Match(text, start);
            return match.Success ? match.Index : -1;
        }
    }
}
