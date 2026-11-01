using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Owns how plain text becomes DrawingML inline content: a:t holds one line of valid XML
    /// characters, so authored line breaks become a:br and characters XML cannot carry are dropped.
    /// </summary>
    internal static class PowerPointTextContent {
        /// <summary>Reduces every authored line break to <c>\n</c>.</summary>
        internal static string NormalizeBreaks(string? text) =>
            (text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Replace('\v', '\n');

        /// <summary>Single-line form for elements that cannot hold a break, such as a field.</summary>
        internal static string ToSingleLine(string? text) => string.Join(" ", SplitLines(text));

        /// <summary>Splits text that separates paragraphs on hard breaks; a vertical tab stays a soft break inside its paragraph.</summary>
        internal static string[] SplitParagraphs(string? text) =>
            (text ?? string.Empty).Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);

        /// <summary>Creates one run per line with <c>a:br</c> between them, all carrying <paramref name="properties"/>.</summary>
        internal static List<OpenXmlElement> CreateInlineContent(string? text, A.RunProperties? properties = null) {
            var content = new List<OpenXmlElement>();
            foreach (string line in SplitLines(text)) {
                if (content.Count > 0) {
                    var lineBreak = new A.Break();
                    ApplyProperties(lineBreak, properties);
                    content.Add(lineBreak);
                }
                var run = new A.Run(new A.Text(line));
                ApplyProperties(run, properties);
                content.Add(run);
            }
            return content;
        }

        /// <summary>Applies a change to the properties of every run and line break below <paramref name="container"/>.</summary>
        internal static void FormatRuns(OpenXmlElement container, Action<A.RunProperties> change) {
            foreach (A.Run run in container.Descendants<A.Run>()) change(run.RunProperties ??= new A.RunProperties());
            foreach (A.Break lineBreak in container.Descendants<A.Break>()) change(lineBreak.RunProperties ??= new A.RunProperties());
        }

        /// <summary>Copies run properties to a run or line break; a break cannot be clicked, so it keeps no interaction.</summary>
        internal static void ApplyProperties(OpenXmlElement node, A.RunProperties? properties) {
            var clone = properties?.CloneNode(true) as A.RunProperties;
            if (node is A.Run run) {
                run.RunProperties = clone;
            } else if (node is A.Break lineBreak) {
                foreach (A.HyperlinkType interaction in clone?.Elements<A.HyperlinkType>().ToArray() ?? Array.Empty<A.HyperlinkType>()) {
                    interaction.Remove();
                }
                lineBreak.RunProperties = clone;
            }
        }

        private static string[] SplitLines(string? text) =>
            LegacyPptXmlText.SanitizeAttributeValue(NormalizeBreaks(text))!.Split('\n');
    }
}
