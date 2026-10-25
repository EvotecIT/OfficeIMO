using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderSectionBlock(SectionBlock section) {
            encounteredSectionDefinitions.Add(section);
            if (section.Options.StartOnNewPage && (pageDirty || HasCurrentPageNonContentObjects() || HasPageFootnotes)) {
                pendingFloatingBookmarks.Clear();
                NewPageAfterFootnoteContinuations();
            }

            void CaptureSectionPlacement(double top) {
                EnsurePage();
                AddNamedDestinationName(section.DestinationName, top);
                currentPage!.Sections.Add(new PageSection {
                    DestinationName = section.DestinationName,
                    Title = section.Title,
                    Level = section.Options.Level,
                    Y = top,
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
                    if (HasFloatingTables) AvoidFloatingBlock(Math.Max(1, MeasureHeadingBlockHeight(heading, currentOpts.MarginLeft, width)));
                    RenderHeadingFlowBlock(heading, null, new IPdfBlock[] { heading }, 0, CaptureSectionPlacement);
                } else CaptureSectionPlacement(y);

                ProcessBlocks(section.Blocks, section);
            } finally {
                flowSemanticScopes.RemoveAt(flowSemanticScopes.Count - 1);
            }
        }

        private void RenderTableOfContentsBlock(TableOfContentsBlock tableOfContents) {
            encounteredTableOfContents = true;
            ProcessBlocks(CreateTableOfContentsBlocks(tableOfContents), tableOfContents);
        }

        private List<IPdfBlock> CreateTableOfContentsBlocks(TableOfContentsBlock tableOfContents) {
            PdfTableOfContentsOptions options = tableOfContents.Options;
            var entries = new List<IPdfBlock>();
            if (!string.IsNullOrWhiteSpace(options.Title))
                entries.Add(new HeadingBlock(1, options.Title!, PdfAlign.Left, color: null));
            IEnumerable<PdfTableOfContentsEntry> source = options.Entries ?? sectionDefinitions
                .Where(section => section.Options.IncludeInTableOfContents)
                .Select(section => new PdfTableOfContentsEntry(section.Title, section.DestinationName, section.Options.Level)).ToArray();
            foreach (PdfTableOfContentsEntry entry in source) {
                if (entry.Level < options.MinimumLevel || entry.Level > options.MaximumLevel) {
                    continue;
                }

                tableOfContentsDestinations.Add(entry.DestinationName);
                double leftIndent = (entry.Level - options.MinimumLevel) * options.IndentPerLevel;
                var style = new PdfParagraphStyle {
                    LeftIndent = leftIndent,
                    SpacingAfter = 2D,
                    KeepTogether = true
                };
                style.AddTabStop(1D, PdfTabAlignment.Right, options.Leader).IsFrameRightRelative = true;
                string pageText = sectionPageNumbers.TryGetValue(entry.DestinationName, out int pageNumber)
                    ? FormatSectionPageNumber(options, pageNumber)
                    : "0000";
                entries.Add(new RichParagraphBlock(
                    new[] {
                        PdfTextRun.LinkToBookmark(entry.Title, entry.DestinationName, underline: false, contents: entry.LinkContents),
                        PdfTextRun.Tab(options.Leader, PdfTabAlignment.Right),
                        PdfTextRun.Normal(pageText)
                    },
                    PdfAlign.Left,
                    defaultColor: null,
                    style));
            }

            return entries;
        }

        public void ValidateTableOfContentsDestinations(IReadOnlyDictionary<string, int> resolved) {
            foreach (string destination in tableOfContentsDestinations) {
                if (!resolved.ContainsKey(destination)) {
                    throw new InvalidOperationException("Table of contents destination '" + destination + "' was not emitted by the final layout.");
                }
            }
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
