using OfficeIMO.Drawing;
using Omd = OfficeIMO.Markdown;

namespace OfficeIMO.Word.Markdown {
    // TeX math from MathInline, MathBlock and math semantic fences becomes native, editable Office Math (OMML).
    internal partial class MarkdownToWordConverter {
        // Math fences carry TeX; AsciiMath fences from AsciiDoc stay code panels.
        private static bool IsMathSemanticBlock(Omd.SemanticFencedBlock block) =>
            string.Equals(block.SemanticKind, Omd.MarkdownSemanticKinds.Math, StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(block.Language, "asciimath", StringComparison.OrdinalIgnoreCase);

        private static bool TryCreateOmml(string tex, bool display, MarkdownWordExportOptions options, out string omml) {
            omml = string.Empty;
            try {
                omml = WordMathMarkup.ToOmml(OfficeMathMarkup.FromLatex(tex), display);
                return true;
            } catch (Exception ex) when (ex is OfficeMathParseException or FormatException or ArgumentException or NotSupportedException) {
                // NotSupportedException covers shared math nodes, such as gathered stacks, that Office Math cannot hold losslessly.
                options.OnWarning?.Invoke($"TeX math '{tex}' could not be converted to Office Math and is kept as text. {ex.Message}");
                return false;
            }
        }

        private static bool TryAddEquation(WordParagraph paragraph, string tex, bool display, MarkdownWordExportOptions options) {
            if (!TryCreateOmml(tex, display, options, out string omml)) return false;
            try {
                paragraph.AddEquation(omml);
                return true;
            } catch (Exception ex) when (ex is ArgumentException or System.Xml.XmlException) {
                options.OnWarning?.Invoke($"TeX math '{tex}' produced Office Math that Word cannot accept and is kept as text. {ex.Message}");
                return false;
            }
        }
    }
}
