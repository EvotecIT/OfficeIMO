using System.Text;

namespace OfficeIMO.Core.Internal {
    /// <summary>Converts engine error codes such as <c>Yaml.Syntax.InvalidIndentation</c> to their catalog form.</summary>
    internal static class OfficeDiagnosticCodeText {
        /// <summary>
        /// Lowercases a dotted, Pascal-cased code and separates its words with hyphens:
        /// <c>Toml.Load.DuplicateKey</c> becomes <c>toml-load-duplicate-key</c>.
        /// </summary>
        public static string FromDottedCode(string code) {
            var text = new StringBuilder(code.Length + 8);
            for (int i = 0; i < code.Length; i++) {
                char ch = code[i];
                if (ch == '.') {
                    text.Append('-');
                } else if (char.IsUpper(ch)) {
                    if (i > 0 && char.IsLower(code[i - 1])) text.Append('-');
                    text.Append(char.ToLowerInvariant(ch));
                } else {
                    text.Append(ch);
                }
            }
            return text.ToString();
        }
    }
}
