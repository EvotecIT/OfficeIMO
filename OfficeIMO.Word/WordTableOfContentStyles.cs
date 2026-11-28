using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using System.Globalization;

namespace OfficeIMO.Word {
    /// <summary>Defines the heading and entry styles referenced by generated tables of contents.</summary>
    internal static class WordTableOfContentStyles {
        private const int W = OfficeXmlNamespaces.WordprocessingML;
        private const int EntryIndentTwips = 220;
        private const int EntryLevels = 9;
        private const int DefaultRightTabTwips = 9350;

        // Called by the SDK default-styles initializer. Both paths use the same definitions.
        internal static IEnumerable<Style> Create(int rightTabTwips = DefaultRightTabTwips) {
            OfficeXmlNode parent = OfficeXmlNode.FromSdk(new Styles());
            foreach (OfficeXmlNode style in Create(parent, rightTabTwips)) yield return (Style)style.SdkElement!;
        }

        internal static void EnsureDefined(StyleDefinitionsPart? part) {
            if (part?.Styles is not { } styles) return;
            var document = part.OpenXmlPackage as WordprocessingDocument;
            OfficeXmlNode body = OfficeXmlNode.FromSdk(document?.MainDocumentPart?.Document?.Body);
            if (EnsureDefined(OfficeXmlNode.FromSdk(styles), body)) styles.Save();
        }

        /// <summary>Adds missing TOC styles while preserving the document's existing definitions.</summary>
        internal static bool EnsureDefined(OfficeXmlNode styles, OfficeXmlNode body) {
            var existing = new HashSet<string>(StringComparer.Ordinal);
            var existingNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (OfficeXmlNode style in styles.Children(W, "style")) {
                if (style.Attribute(W, "styleId") is { } id) existing.Add(id);
                if ((style.Attribute(W, "type") ?? "paragraph") == "paragraph" &&
                    style.Child(W, "name").Attribute(W, "val") is { } name) existingNames.Add(name.Trim());
            }

            bool changed = false;
            foreach (OfficeXmlNode style in Create(styles, ResolveRightTabTwips(body))) {
                string name = style.Child(W, "name").Attribute(W, "val")!;
                if (existingNames.Contains(name.Trim()) || !existing.Add(style.Attribute(W, "styleId")!)) continue;
                styles.Append(style);
                changed = true;
            }
            return changed;
        }

        private static int ResolveRightTabTwips(OfficeXmlNode body) {
            OfficeXmlNode section = default;
            foreach (OfficeXmlNode block in body.Children()) {
                section = block.Is(W, "sectPr") ? block
                    : block.Is(W, "p") ? block.Child(W, "pPr").Child(W, "sectPr") : default;
                if (section.Exists) break;
            }
            uint? width = OfficeXmlViewValues.ParseUInt32(section.Child(W, "pgSz").Attribute(W, "w"));
            OfficeXmlNode margin = section.Child(W, "pgMar");
            if (!width.HasValue || margin.IsNone) return DefaultRightTabTwips;
            long text = (long)width.Value - (OfficeXmlViewValues.ParseUInt32(margin.Attribute(W, "left")) ?? 1440U)
                - (OfficeXmlViewValues.ParseUInt32(margin.Attribute(W, "right")) ?? 1440U) - 10;
            return text > 1440 && text < 31680 ? (int)text : DefaultRightTabTwips;
        }

        private static IEnumerable<OfficeXmlNode> Create(OfficeXmlNode parent, int rightTabTwips) {
            OfficeXmlNode heading = CreateBaseStyle(parent, "TOCHeading", "TOC Heading", "Heading1");
            Append(heading, "uiPriority", "39");
            Append(heading, "unhideWhenUsed");
            Append(heading, "qFormat");
            // Outline level 9 is body text, so the heading never lists itself.
            Append(Append(heading, "pPr"), "outlineLvl", "9");
            yield return heading;

            for (int level = 1; level <= EntryLevels; level++) {
                string number = level.ToString(CultureInfo.InvariantCulture);
                OfficeXmlNode entry = CreateBaseStyle(parent, "TOC" + number, "toc " + number, "Normal");
                Append(entry, "autoRedefine");
                Append(entry, "uiPriority", "39");
                Append(entry, "unhideWhenUsed");
                OfficeXmlNode properties = Append(entry, "pPr");
                OfficeXmlNode tab = Append(Append(properties, "tabs"), "tab", "right");
                tab.SetAttribute(W, "leader", "dot");
                tab.SetAttribute(W, "pos", OfficeXmlViewValues.Format(rightTabTwips));
                Append(properties, "spacing").SetAttribute(W, "after", "100");
                if (level > 1) Append(properties, "ind").SetAttribute(W, "left", OfficeXmlViewValues.Format((level - 1) * EntryIndentTwips));
                yield return entry;
            }
        }

        private static OfficeXmlNode CreateBaseStyle(OfficeXmlNode parent, string id, string name, string baseStyle) {
            OfficeXmlNode style = parent.CreateElement(W, "style");
            style.SetAttribute(W, "type", "paragraph");
            style.SetAttribute(W, "styleId", id);
            Append(style, "name", name);
            Append(style, "basedOn", baseStyle);
            Append(style, "next", "Normal");
            return style;
        }

        private static OfficeXmlNode Append(OfficeXmlNode parent, string name, string? value = null) {
            OfficeXmlNode child = parent.CreateElement(W, name);
            if (value != null) child.SetAttribute(W, "val", value);
            return parent.Append(child);
        }
    }
}
