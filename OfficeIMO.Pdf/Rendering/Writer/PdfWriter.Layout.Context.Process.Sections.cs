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

            flowSemanticScopes.Add(new FlowSemanticScope(MapSemanticStructureType(PdfSemanticRole.Section), alternativeText: null));
            try {
                if (section.Options.IncludeHeading) {
                    var heading = new HeadingBlock(
                            section.Options.Level,
                            section.Title,
                            OfficeAlignment.Left,
                            color: null,
                            style: section.Options.HeadingStyle);
                    if (HasFloatingTables) AvoidFloatingBlock(Math.Max(1, MeasureHeadingBlockHeight(heading, currentOpts.MarginLeft.ToPoints(), width)));
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
                entries.Add(new HeadingBlock(1, options.Title!, OfficeAlignment.Left, color: null));
            IEnumerable<PdfTableOfContentsEntry> source = options.Entries ?? sectionDefinitions
                .Where(section => section.Options.IncludeInTableOfContents)
                .Select(section => new PdfTableOfContentsEntry(section.Title, section.DestinationName, section.Options.Level)).ToArray();
            var items = new List<(int Level, IPdfBlock Paragraph)>();
            foreach (PdfTableOfContentsEntry entry in source) {
                if (entry.Level < options.MinimumLevel || entry.Level > options.MaximumLevel) {
                    continue;
                }

                tableOfContentsDestinations.Add(entry.DestinationName);
                double leftIndent = (entry.Level - options.MinimumLevel) * options.IndentPerLevel.ToPoints();
                var style = new PdfParagraphStyle {
                    LeftIndent = global::OfficeIMO.OfficeLength.Points(leftIndent),
                    SpacingAfter = global::OfficeIMO.OfficeLength.Points(2D),
                    KeepTogether = true
                };
                style.AddTabStop(global::OfficeIMO.OfficeLength.Points(1D), PdfTabAlignment.Right, options.Leader).IsFrameRightRelative = true;
                string pageText = sectionPageNumbers.TryGetValue(entry.DestinationName, out int pageNumber)
                    ? FormatSectionPageNumber(options, pageNumber)
                    : "0000";
                // Tagged output nests each entry as TOCI > Reference > (Link, page number);
                // the dot leader is painted as a layout artifact.
                items.Add((entry.Level, new RichParagraphBlock(
                    new[] {
                        PdfTextRun.LinkToBookmark(entry.Title, entry.DestinationName, underline: false, contents: entry.LinkContents),
                        PdfTextRun.Tab(options.Leader, PdfTabAlignment.Right),
                        PdfTextRun.Normal(pageText)
                    },
                    OfficeAlignment.Left,
                    defaultColor: null,
                    style).AsStructure("Reference")));
            }

            if (items.Count > 0) entries.Add(new SemanticBlock("TOC", CreateTableOfContentsItems(items, 0, items.Count)));
            return entries;
        }

        // A deeper entry belongs to a nested TOC inside the preceding shallower entry's TOCI (ISO 32000-1, 14.8.4.3.2).
        private static List<IPdfBlock> CreateTableOfContentsItems(List<(int Level, IPdfBlock Paragraph)> items, int start, int end) {
            var result = new List<IPdfBlock>();
            int index = start;
            while (index < end) {
                int next = index + 1;
                while (next < end && items[next].Level > items[index].Level) next++;
                var itemBlocks = new List<IPdfBlock> { items[index].Paragraph };
                if (next > index + 1) itemBlocks.Add(new SemanticBlock("TOC", CreateTableOfContentsItems(items, index + 1, next)));
                result.Add(new SemanticBlock("TOCI", itemBlocks));
                index = next;
            }

            return result;
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
