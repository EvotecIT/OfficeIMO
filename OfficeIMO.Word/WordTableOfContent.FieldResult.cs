using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using System.Globalization;

namespace OfficeIMO.Word {
    public partial class WordTableOfContent {
        private void ReplaceContentWithFieldResult(string instruction, string noContentText, IReadOnlyList<TocHeadingEntry> entries,
            ISet<int> pageNumberSuppressedLevels, string pageNumberSeparator) {
            const int W = TocNamespace;
            OfficeXmlNode content = ContentNode;
            if (content.IsNone) throw new InvalidOperationException("Table of contents content block is missing.");

            OfficeXmlNode title = GetTitleParagraphNode();
            if (title.Exists) title.Remove();
            while (content.FirstChild.Exists) content.FirstChild.Remove();
            if (title.Exists) content.Append(title);

            int? rightTabPosition = pageNumberSeparator == "\t" ? GetTextWidthTwips() : null;
            var paragraphs = new List<OfficeXmlNode>(entries.Count);
            foreach (TocHeadingEntry entry in entries) {
                OfficeXmlNode paragraph = content.CreateElement(W, "p");
                PopulateFieldEntry(paragraph, entry, !pageNumberSuppressedLevels.Contains(entry.Level), pageNumberSeparator, rightTabPosition);
                paragraphs.Add(paragraph);
            }
            AppendGeneratedFieldResult(content, instruction, paragraphs, noContentText);
        }

        private static void PopulateFieldEntry(OfficeXmlNode paragraph, TocHeadingEntry entry, bool includePageNumber,
            string pageNumberSeparator, int? rightTabPosition) {
            const int W = TocNamespace;
            OfficeXmlNode properties = AppendTocChild(paragraph, "pPr");
            AppendTocChild(properties, "pStyle").SetAttribute(W, "val", "TOC" + entry.Level.ToString(CultureInfo.InvariantCulture));
            if (includePageNumber && rightTabPosition.HasValue) {
                OfficeXmlNode tab = AppendTocChild(AppendTocChild(properties, "tabs"), "tab");
                tab.SetAttribute(W, "val", "right");
                tab.SetAttribute(W, "leader", "dot");
                tab.SetAttribute(W, "pos", OfficeXmlViewValues.Format(rightTabPosition));
            }

            OfficeXmlNode hyperlink = AppendTocChild(paragraph, "hyperlink");
            hyperlink.SetAttribute(W, "anchor", entry.BookmarkName);
            hyperlink.SetAttribute(W, "history", "true");
            hyperlink.Append(CreateTocTextRun(hyperlink, "t", entry.Text, preserveSpace: true));
            if (!includePageNumber) return;

            if (pageNumberSeparator == "\t") {
                AppendTocChild(AppendTocChild(hyperlink, "r"), "tab");
            } else if (pageNumberSeparator.Length > 0) {
                hyperlink.Append(CreateTocTextRun(hyperlink, "t", pageNumberSeparator, preserveSpace: true));
            }

            hyperlink.Append(CreateTocMarkerRun(hyperlink, "begin"));
            hyperlink.Append(CreateTocTextRun(hyperlink, "instrText", " PAGEREF " + entry.BookmarkName + " \\h ", preserveSpace: true));
            hyperlink.Append(CreateTocMarkerRun(hyperlink, "separate"));
            hyperlink.Append(CreateTocTextRun(hyperlink, "t", entry.PageNumber.ToString(CultureInfo.InvariantCulture), preserveSpace: false));
            hyperlink.Append(CreateTocMarkerRun(hyperlink, "end"));
        }

        /// <summary>Gets the text width of the section holding the table, for the page-number tab.</summary>
        private int GetTextWidthTwips() {
            const int W = TocNamespace;
            const int defaultTextWidth = 9360;
            OfficeXmlNode section = default;
            for (OfficeXmlNode paragraph = SdtNode.NextSiblingNamed(W, "p"); paragraph.Exists;
                 paragraph = paragraph.NextSiblingNamed(W, "p")) {
                section = paragraph.Child(W, "pPr").Child(W, "sectPr");
                if (section.Exists) break;
            }
            if (section.IsNone) section = SdtNode.AncestorOrSelf(W, "body").Children(W, "sectPr").Last();

            OfficeXmlNode margin = section.Child(W, "pgMar");
            uint? width = OfficeXmlViewValues.ParseUInt32(section.Child(W, "pgSz").Attribute(W, "w"));
            if (!width.HasValue) return defaultTextWidth;
            long left = OfficeXmlViewValues.ParseUInt32(margin.Attribute(W, "left")) ?? 1440U;
            long right = OfficeXmlViewValues.ParseUInt32(margin.Attribute(W, "right")) ?? 1440U;
            long gutter = OfficeXmlViewValues.ParseUInt32(margin.Attribute(W, "gutter")) ?? 0U;
            long textWidth = width.Value - left - right - gutter;
            return textWidth > 0 ? (int)textWidth : defaultTextWidth;
        }

        private void AppendGeneratedFieldResult(OfficeXmlNode content, string instruction, IEnumerable<OfficeXmlNode> entries,
            string noContentText) {
            PreservePlaceholderText(noContentText);
            List<OfficeXmlNode> paragraphs = entries.ToList();
            if (paragraphs.Count == 0) {
                OfficeXmlNode empty = content.CreateElement(TocNamespace, "p");
                empty.Append(CreateTocTextRun(empty, "t", noContentText, preserveSpace: true));
                paragraphs.Add(empty);
            }

            OfficeXmlNode first = paragraphs[0];
            OfficeXmlNode properties = first.Child(TocNamespace, "pPr");
            OfficeXmlNode reference = properties.Exists ? properties.NextSibling : first.FirstChild;
            first.InsertBefore(CreateTocMarkerRun(first, "begin"), reference);
            first.InsertBefore(CreateTocTextRun(first, "instrText", instruction, preserveSpace: true), reference);
            first.InsertBefore(CreateTocMarkerRun(first, "separate"), reference);

            OfficeXmlNode last = paragraphs[paragraphs.Count - 1];
            last.Append(CreateTocMarkerRun(last, "end"));
            foreach (OfficeXmlNode paragraph in paragraphs) {
                if (paragraph.Parent.Exists) paragraph.Remove();
                content.Append(paragraph);
            }
            LocalizeGeneratedStyles(content);
            RecordOwnedTocEdit();
        }

        private static OfficeXmlNode CreateTocMarkerRun(OfficeXmlNode parent, string type) {
            OfficeXmlNode run = parent.CreateElement(TocNamespace, "r");
            AppendTocChild(run, "fldChar").SetAttribute(TocNamespace, "fldCharType", type);
            return run;
        }

        private static OfficeXmlNode CreateTocTextRun(OfficeXmlNode parent, string name, string value, bool preserveSpace) {
            OfficeXmlNode run = parent.CreateElement(TocNamespace, "r");
            OfficeXmlNode text = AppendTocChild(run, name);
            text.SetInnerText(value);
            if (preserveSpace) text.SetAttribute(OfficeXmlNamespaces.Xml, "space", "preserve");
            return run;
        }

        private static OfficeXmlNode AppendTocChild(OfficeXmlNode parent, string name) =>
            parent.Append(parent.CreateElement(TocNamespace, name));
    }
}
