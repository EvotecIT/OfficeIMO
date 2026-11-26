using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Owns how plain text becomes DrawingML inline content: a:t holds one line of valid XML
    /// characters, so authored line breaks become a:br and characters XML cannot carry are dropped.
    /// </summary>
    internal static partial class PowerPointTextContent {
        /// <summary>Reduces every authored line break to <c>\n</c>.</summary>
        internal static string NormalizeBreaks(string? text) =>
            (text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Replace('\v', '\n');

        /// <summary>Single-line form for elements that cannot hold a break, such as a field.</summary>
        internal static string ToSingleLine(string? text) => string.Join(" ", SplitLines(text));

        /// <summary>Splits text that separates paragraphs on hard breaks; a vertical tab stays a soft break inside its paragraph.</summary>
        internal static string[] SplitParagraphs(string? text) =>
            (text ?? string.Empty).Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);

        internal static string[] SplitLines(string? text) =>
            LegacyPptXmlText.SanitizeAttributeValue(NormalizeBreaks(text))!.Split('\n');

        /// <summary>
        /// Inserts one run per line, with <c>a:br</c> between them, before <paramref name="before"/> (last when none),
        /// all carrying a copy of <paramref name="properties"/>, and returns what it inserted in order.
        /// </summary>
        internal static List<PowerPointXmlNode> InsertInlineContent(PowerPointXmlNode paragraph, string? text, PowerPointXmlNode properties, PowerPointXmlNode before) {
            var content = new List<PowerPointXmlNode>();
            foreach (string line in SplitLines(text)) {
                if (content.Count > 0) {
                    PowerPointXmlNode lineBreak = paragraph.InsertBefore(paragraph.CreateElement(OfficeXmlNamespaces.DrawingML, "br"), before);
                    ApplyProperties(lineBreak, properties);
                    content.Add(lineBreak);
                }
                PowerPointXmlNode run = paragraph.InsertBefore(paragraph.CreateElement(OfficeXmlNamespaces.DrawingML, "r"), before);
                run.Append(run.CreateElement(OfficeXmlNamespaces.DrawingML, "t")).SetInnerText(line);
                ApplyProperties(run, properties);
                content.Add(run);
            }
            return content;
        }

        /// <summary>Applies a change to the properties of every run and line break below <paramref name="container"/>.</summary>
        internal static void FormatRuns(PowerPointXmlNode container, Action<PowerPointXmlNode> change) {
            var runs = new List<PowerPointXmlNode>();
            var breaks = new List<PowerPointXmlNode>();
            Collect(container, runs, breaks);
            foreach (PowerPointXmlNode run in runs) change(run.GetOrPrependChild(OfficeXmlNamespaces.DrawingML, "rPr"));
            foreach (PowerPointXmlNode lineBreak in breaks) change(lineBreak.GetOrPrependChild(OfficeXmlNamespaces.DrawingML, "rPr"));
        }

        private static void Collect(PowerPointXmlNode element, List<PowerPointXmlNode> runs, List<PowerPointXmlNode> breaks) {
            foreach (PowerPointXmlNode child in element.Children()) {
                if (child.Is(OfficeXmlNamespaces.DrawingML, "r")) runs.Add(child);
                else if (child.Is(OfficeXmlNamespaces.DrawingML, "br")) breaks.Add(child);
                Collect(child, runs, breaks);
            }
        }

        /// <summary>Copies run properties to a run or line break; a break cannot be clicked, so it keeps no interaction.</summary>
        internal static void ApplyProperties(PowerPointXmlNode node, PowerPointXmlNode properties) {
            node.RemoveChildren(OfficeXmlNamespaces.DrawingML, "rPr");
            if (properties.IsNone) return;
            PowerPointXmlNode copy = node.InsertBefore(CopyFor(node, properties), node.FirstChild);
            if (!node.Is(OfficeXmlNamespaces.DrawingML, "br")) return;
            foreach (PowerPointXmlNode interaction in copy.Children().ToList()) {
                if (interaction.Is(OfficeXmlNamespaces.DrawingML, "hlinkClick") || interaction.Is(OfficeXmlNamespaces.DrawingML, "hlinkMouseOver")) interaction.Remove();
            }
        }

        // A detached copy of run properties, renamed a:rPr when they come from elsewhere (end or default properties).
        private static PowerPointXmlNode CopyFor(PowerPointXmlNode target, PowerPointXmlNode properties) {
            PowerPointXmlNode copy = properties.Clone();
            if (copy.Is(OfficeXmlNamespaces.DrawingML, "rPr")) return copy;
            PowerPointXmlNode renamed = target.CreateElement(OfficeXmlNamespaces.DrawingML, "rPr");
            foreach (string attribute in RunPropertyAttributes) {
                if (copy.Attribute(attribute) is { } value) renamed.SetAttribute(attribute, value);
            }
            foreach (PowerPointXmlNode child in copy.Children().ToList()) renamed.Append(child.Clone());
            return renamed;
        }

        private static readonly string[] RunPropertyAttributes = {
            "kumimoji", "lang", "altLang", "sz", "b", "i", "u", "strike", "kern", "cap", "spc", "normalizeH", "baseline",
            "noProof", "dirty", "err", "smtClean", "smtId", "bmk"
        };
    }
}
