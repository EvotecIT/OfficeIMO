using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    internal static class PowerPointTableTextDefaults {
        internal const string Language = "en-US";

        /// <summary>A detached paragraph of a new cell: one run of text, or one empty run, and end properties with a language.</summary>
        internal static OfficeXmlNode CreateParagraph(OfficeXmlNode context, string text = "") {
            OfficeXmlNode paragraph = context.CreateElement(OfficeXmlNamespaces.DrawingML, "p");
            if (string.IsNullOrEmpty(text)) {
                OfficeXmlNode run = paragraph.Append(paragraph.CreateElement(OfficeXmlNamespaces.DrawingML, "r"));
                run.Append(CreateProperties(run, "rPr"));
                run.Append(run.CreateElement(OfficeXmlNamespaces.DrawingML, "t"));
            } else {
                PowerPointTextContent.InsertInlineContent(paragraph, text, CreateProperties(paragraph, "rPr"), default);
            }
            paragraph.Append(CreateProperties(paragraph, "endParaRPr"));
            return paragraph;
        }

        /// <summary>Detached run or end properties carrying only the default language.</summary>
        internal static OfficeXmlNode CreateProperties(OfficeXmlNode context, string localName) {
            OfficeXmlNode properties = context.CreateElement(OfficeXmlNamespaces.DrawingML, localName);
            properties.SetAttribute("lang", Language);
            return properties;
        }
    }
}
