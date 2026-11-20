using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;

namespace OfficeIMO.Word {
    public partial class WordTableOfContent {
        private void ReplaceContentWithFieldResult(string instruction, string noContentText, IReadOnlyList<TocHeadingEntry> entries, ISet<int> pageNumberSuppressedLevels, string pageNumberSeparator) {
            SdtContentBlock content = _sdtBlock.SdtContentBlock
                ?? throw new InvalidOperationException("Table of contents content block is missing.");

            Paragraph titleParagraph = GetTitleParagraph() ?? CreateTitleParagraph("Table of Contents");
            RemoveIfAttached(titleParagraph);
            content.RemoveAllChildren();
            content.Append(titleParagraph);

            int? rightTabPosition = pageNumberSeparator == "\t" ? GetTextWidthTwips() : null;
            var paragraphs = new List<Paragraph>();
            foreach (TocHeadingEntry entry in entries) {
                paragraphs.Add(CreateFieldEntryParagraph(entry, !pageNumberSuppressedLevels.Contains(entry.Level), pageNumberSeparator, rightTabPosition));
            }

            AppendGeneratedFieldResult(content, instruction, paragraphs, noContentText);
        }

        private static Paragraph CreateFieldEntryParagraph(TocHeadingEntry entry, bool includePageNumber, string pageNumberSeparator, int? rightTabPosition) {
            var properties = new ParagraphProperties(new ParagraphStyleId { Val = "TOC" + entry.Level.ToString(CultureInfo.InvariantCulture) });
            if (includePageNumber && rightTabPosition != null) {
                properties.Tabs = new Tabs(new TabStop {
                    Val = TabStopValues.Right,
                    Leader = TabStopLeaderCharValues.Dot,
                    Position = rightTabPosition.Value
                });
            }

            var hyperlink = new Hyperlink {
                Anchor = entry.BookmarkName,
                History = true
            };
            hyperlink.Append(new Run(new Text(entry.Text) { Space = SpaceProcessingModeValues.Preserve }));
            if (includePageNumber) {
                if (pageNumberSeparator == "\t") {
                    hyperlink.Append(new Run(new TabChar()));
                } else if (pageNumberSeparator.Length > 0) {
                    hyperlink.Append(new Run(new Text(pageNumberSeparator) { Space = SpaceProcessingModeValues.Preserve }));
                }

                hyperlink.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }));
                hyperlink.Append(new Run(new FieldCode(" PAGEREF " + entry.BookmarkName + " \\h ") { Space = SpaceProcessingModeValues.Preserve }));
                hyperlink.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }));
                hyperlink.Append(new Run(new Text(entry.PageNumber.ToString(CultureInfo.InvariantCulture))));
                hyperlink.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
            }

            return new Paragraph(properties, hyperlink);
        }

        /// <summary>
        /// Gets the text width of the section that holds the table, for the right-aligned page number tab.
        /// </summary>
        private int GetTextWidthTwips() {
            const int defaultTextWidth = 9360;
            SectionProperties? section = _sdtBlock.ElementsAfter()
                .OfType<Paragraph>()
                .Select(_ => _.ParagraphProperties?.SectionProperties)
                .FirstOrDefault(_ => _ != null)
                ?? _sdtBlock.Ancestors<Body>().FirstOrDefault()?.Elements<SectionProperties>().LastOrDefault();
            PageSize? size = section?.GetFirstChild<PageSize>();
            PageMargin? margin = section?.GetFirstChild<PageMargin>();
            if (size?.Width?.Value is not uint width) {
                return defaultTextWidth;
            }

            long left = margin?.Left?.Value ?? 1440U;
            long right = margin?.Right?.Value ?? 1440U;
            long gutter = margin?.Gutter?.Value ?? 0U;
            long textWidth = width - left - right - gutter;
            if (textWidth > 0) {
                return (int)textWidth;
            }

            return defaultTextWidth;
        }

        private void AppendGeneratedFieldResult(SdtContentBlock content, string instruction, IEnumerable<Paragraph> entries, string noContentText) {
            PreservePlaceholderText(noContentText);
            var paragraphs = entries.ToList();
            if (paragraphs.Count == 0) {
                paragraphs.Add(new Paragraph(new Run(new Text(noContentText) { Space = SpaceProcessingModeValues.Preserve })));
            }

            Paragraph first = paragraphs[0];
            OpenXmlElement? anchor = first.ParagraphProperties;
            foreach (Run fieldStart in new[] {
                         new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                         new Run(new FieldCode(instruction) { Space = SpaceProcessingModeValues.Preserve }),
                         new Run(new FieldChar { FieldCharType = FieldCharValues.Separate })
                     }) {
                anchor = anchor == null ? first.PrependChild(fieldStart) : first.InsertAfter(fieldStart, anchor);
            }

            paragraphs[paragraphs.Count - 1].Append(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
            foreach (Paragraph paragraph in paragraphs) {
                RemoveIfAttached(paragraph);
                content.Append(paragraph);
            }
            _document.LocalizeBuiltInStyleReferences(content);
        }

    }
}
