using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word {
    public partial class WordField {
        private static OfficeXmlNode CreateFieldRun(OfficeXmlNode parent) {
            OfficeXmlNode run = parent.CreateElement(W, "r");
            OfficeXmlNode properties = run.Append(run.CreateElement(W, "rPr"));
            properties.Append(properties.CreateElement(W, "noProof"));
            return run;
        }

        private static OfficeXmlNode AppendMarkerRun(OfficeXmlNode paragraph, string type) {
            OfficeXmlNode run = CreateFieldRun(paragraph);
            OfficeXmlNode marker = run.CreateElement(W, "fldChar");
            marker.SetAttribute(W, "fldCharType", type);
            run.Append(marker);
            return paragraph.Append(run);
        }

        private static OfficeXmlNode AppendTextRun(OfficeXmlNode parent, string localName, string value, bool preserveSpace) {
            OfficeXmlNode run = CreateFieldRun(parent);
            OfficeXmlNode text = run.CreateElement(W, localName);
            text.SetInnerText(value);
            if (preserveSpace) text.SetAttribute(OfficeXmlNamespaces.Xml, "space", "preserve");
            run.Append(text);
            return parent.Append(run);
        }

        private static string GenerateField(WordFieldType wordFieldType, WordFieldFormat? wordFieldFormat = null, string? customFormat = null, List<string>? parameters = null) {
            var fieldType = " " + wordFieldType.ToString().ToUpper() + " ";
            var fieldFormat = string.Empty;
            if (wordFieldFormat != null) {
                fieldFormat = @"\* " + WordFieldFormatCodec.ToInstruction(wordFieldFormat.Value) + " ";
            }
            if (!string.IsNullOrWhiteSpace(customFormat)) {
                fieldFormat += $"\\@ \"{customFormat}\" ";
            }

            var switchesList = " ";
            if (parameters != null && parameters.Count > 0) {
                switchesList += string.Join(" ", parameters.Select(s => s.Trim()));
            }

            return fieldType + switchesList + fieldFormat + @"\* MERGEFORMAT ";
        }
    }
}
