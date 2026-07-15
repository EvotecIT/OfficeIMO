using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderSectionBlock(SectionBlock section) {
            encounteredSectionDefinitions.Add(section);
            if (section.Options.StartOnNewPage && (pageDirty || HasCurrentPageNonContentObjects())) {
                NewPage();
            }

            EnsurePage();
            AddNamedDestinationName(section.DestinationName, y);
            currentPage!.Sections.Add(new PageSection {
                DestinationName = section.DestinationName,
                Title = section.Title,
                Level = section.Options.Level,
                Y = y,
                Reference = section.Options.Reference
            });

            if (section.Options.IncludeHeading) {
                ProcessBlocks(new IPdfBlock[] {
                    new HeadingBlock(
                        section.Options.Level,
                        section.Title,
                        PdfAlign.Left,
                        color: null,
                        style: section.Options.HeadingStyle)
                });
            }

            ProcessBlocks(section.Blocks);
        }

        private void RenderTableOfContentsBlock(TableOfContentsBlock tableOfContents) {
            encounteredTableOfContents = true;
            PdfTableOfContentsOptions options = tableOfContents.Options;
            if (!string.IsNullOrWhiteSpace(options.Title)) {
                ProcessBlocks(new IPdfBlock[] { new HeadingBlock(1, options.Title!, PdfAlign.Left, color: null) });
            }

            var entries = new List<IPdfBlock>();
            for (int i = 0; i < sectionDefinitions.Count; i++) {
                SectionBlock section = sectionDefinitions[i];
                if (!section.Options.IncludeInTableOfContents ||
                    section.Options.Level < options.MinimumLevel ||
                    section.Options.Level > options.MaximumLevel) {
                    continue;
                }

                double leftIndent = (section.Options.Level - options.MinimumLevel) * options.IndentPerLevel;
                double tabPosition = Math.Max(12D, width - leftIndent - 1D);
                var style = new PdfParagraphStyle {
                    LeftIndent = leftIndent,
                    SpacingAfter = 2D,
                    KeepTogether = true
                };
                style.AddTabStop(tabPosition, PdfTabAlignment.Right, options.Leader);
                string pageText = sectionPageNumbers.TryGetValue(section.DestinationName, out int pageNumber)
                    ? FormatSectionPageNumber(options, pageNumber)
                    : "0000";
                entries.Add(new RichParagraphBlock(
                    new[] {
                        TextRun.LinkToBookmark(section.Title, section.DestinationName, underline: false),
                        TextRun.Tab(options.Leader, PdfTabAlignment.Right),
                        TextRun.Normal(pageText)
                    },
                    PdfAlign.Left,
                    defaultColor: null,
                    style));
            }

            ProcessBlocks(entries);
        }

        private static string FormatSectionPageNumber(PdfTableOfContentsOptions options, int pageNumber) {
            string result = options.PageNumberFormatter?.Invoke(pageNumber)
                ?? pageNumber.ToString(CultureInfo.InvariantCulture);
            if (string.IsNullOrWhiteSpace(result)) {
                throw new InvalidOperationException("TOC page number formatter returned an empty value.");
            }

            return result;
        }
    }
}
