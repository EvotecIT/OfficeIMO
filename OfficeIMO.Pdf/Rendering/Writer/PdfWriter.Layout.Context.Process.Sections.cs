using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderSectionBlock(SectionBlock section) {
            encounteredSectionDefinitions.Add(section);
            if (section.Options.StartOnNewPage && (pageDirty || HasCurrentPageNonContentObjects())) {
                pendingFloatingBookmarks.Clear();
                NewPage();
            }

            void CaptureSectionPlacement() {
                EnsurePage();
                AddNamedDestinationName(section.DestinationName, y);
                currentPage!.Sections.Add(new PageSection {
                    DestinationName = section.DestinationName,
                    Title = section.Title,
                    Level = section.Options.Level,
                    Y = y,
                    Reference = section.Options.Reference
                });
            }

            flowSemanticScopes.Add(new FlowSemanticScope(PdfSemanticRole.Section, alternativeText: null));
            try {
                if (section.Options.IncludeHeading) {
                    var heading = new HeadingBlock(
                            section.Options.Level,
                            section.Title,
                            PdfAlign.Left,
                            color: null,
                            style: section.Options.HeadingStyle);
                    if (HasFloatingTables) AvoidFloatingBlock(Math.Max(1, MeasureHeadingBlockHeight(heading, width)));
                    RenderHeadingFlowBlock(heading, null, new IPdfBlock[] { heading }, 0, CaptureSectionPlacement);
                } else CaptureSectionPlacement();

                ProcessBlocks(section.Blocks, section);
            } finally {
                flowSemanticScopes.RemoveAt(flowSemanticScopes.Count - 1);
            }
        }

        private void RenderTableOfContentsBlock(TableOfContentsBlock tableOfContents) {
            encounteredTableOfContents = true;
            PdfTableOfContentsOptions options = tableOfContents.Options;
            var entries = new List<IPdfBlock>();
            if (!string.IsNullOrWhiteSpace(options.Title))
                entries.Add(new HeadingBlock(1, options.Title!, PdfAlign.Left, color: null));
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
                        PdfTextRun.LinkToBookmark(section.Title, section.DestinationName, underline: false),
                        PdfTextRun.Tab(options.Leader, PdfTabAlignment.Right),
                        PdfTextRun.Normal(pageText)
                    },
                    PdfAlign.Left,
                    defaultColor: null,
                    style));
            }

            ProcessBlocks(entries, tableOfContents);
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
