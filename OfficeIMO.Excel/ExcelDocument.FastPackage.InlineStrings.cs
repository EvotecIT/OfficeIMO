using System.Linq;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static bool TryAppendPlainInlineString(StringBuilder builder, InlineString inline) {
            if (inline.HasAttributes || inline.NamespaceDeclarations.Any()
                || inline.FirstChild is not Text text || text.NextSibling() != null
                || text.NamespaceDeclarations.Any() || text.ExtendedAttributes.Any()
                || text.MCAttributes != null) {
                return false;
            }

            string? space = text.Space?.InnerText;
            if (space != null && space != "preserve" && space != "default") return false;
            string value = text.Text ?? string.Empty;
            // Keep XmlWriter's character validation; direct emission must not accept
            // malformed UTF-16 or XML control characters that OuterXml would reject.
            XmlConvert.VerifyXmlChars(value);
            builder.Append("<is><t");
            if (space != null) {
                builder.Append(" xml:space=\"").Append(space).Append('"');
            }
            builder.Append('>');
            AppendXmlEscaped(builder, value);
            builder.Append("</t></is>");
            return true;
        }
    }
}
