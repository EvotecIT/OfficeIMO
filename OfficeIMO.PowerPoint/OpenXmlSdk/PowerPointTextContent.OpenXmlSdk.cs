using DocumentFormat.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    /// <summary>The SDK forms of the inline text helpers, for code that still builds SDK elements.</summary>
    internal static partial class PowerPointTextContent {
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
        internal static void ApplyProperties(OpenXmlElement node, A.RunProperties? properties) =>
            ApplyProperties(OfficeXmlNode.FromSdk(node), OfficeXmlNode.FromSdk(properties));
    }
}
