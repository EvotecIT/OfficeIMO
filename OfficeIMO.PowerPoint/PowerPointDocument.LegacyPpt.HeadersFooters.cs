using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        /// <summary>
        /// Writes <c>p:hf</c> into a slide master, layout, notes master or handout master root, and the header and footer
        /// text into its placeholders.
        /// </summary>
        private static void ApplyLegacyHeaderFooter(OfficeXmlNode owner,
            OfficeXmlNode commonSlideData, LegacyPptHeaderFooterSettings? source,
            bool allowHeader) {
            if (source == null) return;
            owner.RemoveChildren(LegacyPptXml.P, "hf");
            OfficeXmlNode headerFooter = owner.CreateElement(LegacyPptXml.P, "hf")
                .Set("sldNum", source.ShowSlideNumber)
                .Set("hdr", allowHeader && source.ShowHeader)
                .Set("ftr", source.ShowFooter)
                .Set("dt", source.ShowDate);
            string[] successors = owner.LocalName switch {
                "sldMaster" => new[] { "txStyles", "extLst" },
                "sldLayout" => new[] { "timing", "transition", "extLst" },
                "notesMaster" => new[] { "notesStyle", "extLst" },
                "handoutMaster" => new[] { "extLst" },
                _ => Array.Empty<string>()
            };
            OfficeXmlNode successor = default;
            foreach (string name in successors) {
                successor = owner.Child(LegacyPptXml.P, name);
                if (!successor.IsNone) break;
            }
            owner.InsertBefore(headerFooter, successor);

            if (commonSlideData.IsNone) return;
            if (source.UseUserDate && source.UserDateText.Length > 0) {
                SetLegacyPlaceholderText(commonSlideData, PowerPointPlaceholderType.DateAndTime,
                    source.UserDateText);
            }
            if (allowHeader && source.HeaderText.Length > 0) {
                SetLegacyPlaceholderText(commonSlideData, PowerPointPlaceholderType.Header,
                    source.HeaderText);
            }
            if (source.FooterText.Length > 0) {
                SetLegacyPlaceholderText(commonSlideData, PowerPointPlaceholderType.Footer,
                    source.FooterText);
            }
        }

        private static void SetLegacyPlaceholderText(OfficeXmlNode commonSlideData,
            PowerPointPlaceholderType type, string text) {
            string token = type.ToOpenXmlToken();
            OfficeXmlNode shape = commonSlideData.Child(LegacyPptXml.P, "spTree").Children(LegacyPptXml.P, "sp").ToList()
                .FirstOrDefault(candidate => candidate.Path(LegacyPptXml.P, "nvSpPr", "nvPr", "ph").Attribute("type") == token);
            OfficeXmlNode textBody = shape.Child(LegacyPptXml.P, "txBody");
            if (textBody.IsNone) return;
            textBody.RemoveChildren(LegacyPptXml.A, "p");
            AppendLegacyPlainParagraph(textBody, text);
        }

        // A paragraph of plain text with empty end properties, as header and footer placeholders hold it.
        private static void AppendLegacyPlainParagraph(OfficeXmlNode textBody, string text) {
            OfficeXmlNode paragraph = textBody.AddA("p");
            PowerPointTextContent.InsertInlineContent(paragraph, text, default, default);
            paragraph.AddA("endParaRPr");
        }

        private static LegacyPptHeaderFooterSettings? GetEffectiveLegacySlideHeaderFooter(
            LegacyPptDocument legacy, LegacyPptMaster master) =>
            master.HeaderFooter ?? legacy.SlideHeaderFooterDefaults;
    }
}
