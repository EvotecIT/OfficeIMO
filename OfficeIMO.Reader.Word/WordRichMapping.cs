using OfficeIMO.Word;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeIMO.Reader.Word;

internal static class WordRichMapping {
    internal static OfficeDocumentReadResult Apply(
        WordDocumentSnapshot snapshot,
        IReadOnlyList<WordDocumentVisualSnapshot> pageSnapshots,
        ReaderOptions readerOptions,
        ReaderWordOptions options,
        OfficeDocumentReadResult result) {
        result.Source.Title = snapshot.Title;
        result.Source.Author = snapshot.Author;
        result.Source.Subject = snapshot.Subject;
        result.Source.Keywords = snapshot.Keywords;

        var blocks = new List<OfficeDocumentBlock>();
        var tables = new List<ReaderTable>();
        var links = new List<OfficeDocumentLink>();
        int sourceBlockIndex = 0;
        int tableIndex = 0;
        int linkIndex = 0;
        for (int sectionIndex = 0; sectionIndex < snapshot.Sections.Count; sectionIndex++) {
            WordSectionSnapshot section = snapshot.Sections[sectionIndex];
            ProjectWordElements(
                section.Elements,
                result.Source.Path,
                "section-" + (sectionIndex + 1).ToString("D4", CultureInfo.InvariantCulture),
                readerOptions,
                blocks,
                tables,
                links,
                ref sourceBlockIndex,
                ref tableIndex,
                ref linkIndex);

            ProjectWordHeaderFooter(section.DefaultHeader, sectionIndex, result.Source.Path, readerOptions, blocks, tables, links, ref sourceBlockIndex, ref tableIndex, ref linkIndex);
            ProjectWordHeaderFooter(section.FirstHeader, sectionIndex, result.Source.Path, readerOptions, blocks, tables, links, ref sourceBlockIndex, ref tableIndex, ref linkIndex);
            ProjectWordHeaderFooter(section.EvenHeader, sectionIndex, result.Source.Path, readerOptions, blocks, tables, links, ref sourceBlockIndex, ref tableIndex, ref linkIndex);
            ProjectWordHeaderFooter(section.DefaultFooter, sectionIndex, result.Source.Path, readerOptions, blocks, tables, links, ref sourceBlockIndex, ref tableIndex, ref linkIndex);
            ProjectWordHeaderFooter(section.FirstFooter, sectionIndex, result.Source.Path, readerOptions, blocks, tables, links, ref sourceBlockIndex, ref tableIndex, ref linkIndex);
            ProjectWordHeaderFooter(section.EvenFooter, sectionIndex, result.Source.Path, readerOptions, blocks, tables, links, ref sourceBlockIndex, ref tableIndex, ref linkIndex);
        }

        if (options.IncludeFootnotes) {
            ProjectWordNotes(snapshot, result.Source.Path, blocks, links, ref sourceBlockIndex, ref linkIndex);
        }

        var metadata = new[] {
            BuildCountMetadataEntry("word-section-count", "word.structure", "SectionCount", snapshot.Sections.Count)
        };
        IReadOnlyList<OfficeDocumentPage> pages = BuildWordPages(
            pageSnapshots,
            blocks,
            tables,
            links,
            result.Source.Path);
        IReadOnlyList<string> capabilities = options.IncludePageLocations
            ? new[] {
                "officeimo.word.inspection-snapshot",
                "officeimo.reader.word.rich-v5",
                "officeimo.reader.word.pages.computed"
            }
            : new[] {
                "officeimo.word.inspection-snapshot",
                "officeimo.reader.word.rich-v5"
            };
        return DocumentReaderEngine.EnrichDocumentResult(
            result,
            capabilities,
            blocks,
            tables,
            links,
            result.Visuals,
            pages,
            metadata);
    }

    private static IReadOnlyList<OfficeDocumentPage> BuildWordPages(
        IReadOnlyList<WordDocumentVisualSnapshot> pageSnapshots,
        IReadOnlyList<OfficeDocumentBlock> blocks,
        IReadOnlyList<ReaderTable> tables,
        IReadOnlyList<OfficeDocumentLink> links,
        string? sourcePath) {
        if (pageSnapshots.Count == 0) {
            return Array.Empty<OfficeDocumentPage>();
        }

        var blocksById = blocks
            .Where(static block => !string.IsNullOrEmpty(block.Id))
            .ToDictionary(static block => block.Id, StringComparer.Ordinal);
        var pages = new List<OfficeDocumentPage>(pageSnapshots.Count);
        foreach (WordDocumentVisualSnapshot snapshot in pageSnapshots) {
            int pageNumber = snapshot.PageIndex + 1;
            var pageBlocks = new List<OfficeDocumentBlock>();
            foreach (IGrouping<(int SectionIndex, int BlockIndex), WordDocumentVisualFragment> group in snapshot.Fragments
                         .GroupBy(static fragment => (fragment.SectionIndex, fragment.BlockIndex))) {
                string blockId = "word-section-" +
                    (group.Key.SectionIndex + 1).ToString("D4", CultureInfo.InvariantCulture) +
                    "-block-" +
                    group.Key.BlockIndex.ToString("D4", CultureInfo.InvariantCulture);
                if (!blocksById.TryGetValue(blockId, out OfficeDocumentBlock? sourceBlock)) {
                    continue;
                }

                string visibleText = CombineWordFragmentText(
                    group.Select(static fragment => fragment.Text));
                OfficeDocumentRegion? region = CombineWordRegions(group);
                pageBlocks.Add(CloneWordPageBlock(
                    sourceBlock,
                    pageNumber,
                    visibleText,
                    region));
            }

            string[] blockIds = pageBlocks.Select(static block => block.Id).ToArray();
            pages.Add(new OfficeDocumentPage {
                Number = pageNumber,
                Name = "Page " + pageNumber.ToString(CultureInfo.InvariantCulture),
                Width = snapshot.Width,
                Height = snapshot.Height,
                Location = new ReaderLocation {
                    Path = sourcePath,
                    Page = pageNumber,
                    SourceBlockKind = "page",
                    BlockAnchor = "word-page-" + pageNumber.ToString("D4", CultureInfo.InvariantCulture)
                },
                Blocks = pageBlocks.AsReadOnly(),
                Tables = tables.Where(table =>
                    table.Location?.BlockAnchor != null &&
                    blockIds.Contains(table.Location.BlockAnchor, StringComparer.Ordinal)).ToArray(),
                Links = links.Where(link =>
                    link.Location.BlockAnchor != null &&
                    blockIds.Any(blockId => link.Location.BlockAnchor.StartsWith(blockId, StringComparison.Ordinal))).ToArray()
            });
        }

        return pages.AsReadOnly();
    }

    internal static string CombineWordFragmentText(IEnumerable<string> fragmentText) {
        return string.Join(
            Environment.NewLine,
            fragmentText.Where(static text => !string.IsNullOrWhiteSpace(text)));
    }

    private static OfficeDocumentBlock CloneWordPageBlock(
        OfficeDocumentBlock source,
        int pageNumber,
        string text,
        OfficeDocumentRegion? region) {
        ReaderLocation location = CloneWordLocation(source.Location, source.Kind, source.Id);
        location.Page = pageNumber;
        return new OfficeDocumentBlock {
            Id = source.Id,
            Kind = source.Kind,
            Text = text,
            Level = source.Level,
            Marker = source.Marker,
            Location = location,
            Region = region
        };
    }

    private static OfficeDocumentRegion? CombineWordRegions(
        IEnumerable<WordDocumentVisualFragment> fragments) {
        bool hasRegion = false;
        double left = double.MaxValue;
        double top = double.MaxValue;
        double right = double.MinValue;
        double bottom = double.MinValue;
        foreach (WordDocumentVisualFragment fragment in fragments) {
            WordDocumentVisualRegion? region = fragment.Region;
            if (region == null) {
                continue;
            }
            left = Math.Min(left, region.X);
            top = Math.Min(top, region.Y);
            right = Math.Max(right, region.X + region.Width);
            bottom = Math.Max(bottom, region.Y + region.Height);
            hasRegion = true;
        }

        return hasRegion
            ? new OfficeDocumentRegion {
                X = left,
                Y = top,
                Width = Math.Max(0D, right - left),
                Height = Math.Max(0D, bottom - top)
            }
            : null;
    }

    private static void ProjectWordNotes(
        WordDocumentSnapshot snapshot,
        string? sourcePath,
        List<OfficeDocumentBlock> blocks,
        List<OfficeDocumentLink> links,
        ref int sourceBlockIndex,
        ref int linkIndex) {
        var emitted = new HashSet<string>(StringComparer.Ordinal);
        int noteIndex = 0;
        foreach (WordSectionSnapshot section in snapshot.Sections) {
            foreach (WordParagraphSnapshot paragraph in EnumerateWordParagraphs(section)) {
                foreach (WordRunSnapshot run in paragraph.Runs) {
                    if (run.Footnote != null) {
                        ProjectWordNote("footnote", run.Footnote.ReferenceId, run.Footnote.Paragraphs, sourcePath, blocks, links, emitted, ref sourceBlockIndex, ref noteIndex, ref linkIndex);
                    }
                    if (run.Endnote != null) {
                        ProjectWordNote("endnote", run.Endnote.ReferenceId, run.Endnote.Paragraphs, sourcePath, blocks, links, emitted, ref sourceBlockIndex, ref noteIndex, ref linkIndex);
                    }
                }
            }
        }
    }

    private static IEnumerable<WordParagraphSnapshot> EnumerateWordParagraphs(WordSectionSnapshot section) {
        foreach (WordParagraphSnapshot paragraph in EnumerateWordParagraphs(section.Elements)) yield return paragraph;
        foreach (WordHeaderFooterSnapshot? headerFooter in new[] {
            section.DefaultHeader, section.FirstHeader, section.EvenHeader,
            section.DefaultFooter, section.FirstFooter, section.EvenFooter
        }) {
            if (headerFooter == null) continue;
            foreach (WordParagraphSnapshot paragraph in EnumerateWordParagraphs(headerFooter.Elements)) yield return paragraph;
        }
    }

    private static IEnumerable<WordParagraphSnapshot> EnumerateWordParagraphs(IReadOnlyList<WordBlockSnapshot> elements) {
        foreach (WordBlockSnapshot element in elements) {
            if (element is WordParagraphSnapshot paragraph) {
                yield return paragraph;
            } else if (element is WordTableSnapshot table) {
                foreach (WordTableRowSnapshot row in table.Rows) {
                    foreach (WordTableCellSnapshot cell in row.Cells) {
                        foreach (WordParagraphSnapshot cellParagraph in cell.Paragraphs) yield return cellParagraph;
                    }
                }
            }
        }
    }

    private static void ProjectWordNote(
        string kind,
        long? referenceId,
        IReadOnlyList<WordParagraphSnapshot> noteParagraphs,
        string? sourcePath,
        List<OfficeDocumentBlock> blocks,
        List<OfficeDocumentLink> links,
        HashSet<string> emitted,
        ref int sourceBlockIndex,
        ref int noteIndex,
        ref int linkIndex) {
        string text = string.Join(
            Environment.NewLine,
            noteParagraphs.Select(static paragraph => paragraph.Text).Where(static value => !string.IsNullOrWhiteSpace(value)));
        if (string.IsNullOrWhiteSpace(text)) return;

        string identity = referenceId.HasValue
            ? kind + ":" + referenceId.Value.ToString(CultureInfo.InvariantCulture)
            : kind + ":ordinal:" + noteIndex.ToString(CultureInfo.InvariantCulture);
        if (!emitted.Add(identity)) return;

        string reference = referenceId.HasValue
            ? referenceId.Value.ToString(CultureInfo.InvariantCulture).Replace("-", "negative-")
            : noteIndex.ToString("D4", CultureInfo.InvariantCulture);
        string anchor = "word-" + kind + "-" + reference;
        var location = new ReaderLocation {
            Path = sourcePath,
            SourceBlockIndex = sourceBlockIndex,
            SourceBlockKind = kind,
            BlockAnchor = anchor
        };
        blocks.Add(new OfficeDocumentBlock {
            Id = anchor,
            Kind = kind,
            Text = text,
            Location = location
        });
        foreach (WordParagraphSnapshot paragraph in noteParagraphs) {
            AddWordParagraphLinks(paragraph, location, links, ref linkIndex);
        }
        sourceBlockIndex++;
        noteIndex++;
    }

    private static void ProjectWordHeaderFooter(
        WordHeaderFooterSnapshot? headerFooter,
        int sectionIndex,
        string? sourcePath,
        ReaderOptions options,
        List<OfficeDocumentBlock> blocks,
        List<ReaderTable> tables,
        List<OfficeDocumentLink> links,
        ref int sourceBlockIndex,
        ref int tableIndex,
        ref int linkIndex) {
        if (headerFooter == null) return;
        string owner = "section-" + (sectionIndex + 1).ToString("D4", CultureInfo.InvariantCulture)
            + "-" + headerFooter.Kind + "-" + headerFooter.Variant;
        ProjectWordElements(headerFooter.Elements, sourcePath, owner, options, blocks, tables, links, ref sourceBlockIndex, ref tableIndex, ref linkIndex);
    }

    private static void ProjectWordElements(
        IReadOnlyList<WordBlockSnapshot> elements,
        string? sourcePath,
        string owner,
        ReaderOptions options,
        List<OfficeDocumentBlock> blocks,
        List<ReaderTable> tables,
        List<OfficeDocumentLink> links,
        ref int sourceBlockIndex,
        ref int tableIndex,
        ref int linkIndex) {
        var headingStack = new List<(int Level, string Text)>();
        for (int elementIndex = 0; elementIndex < elements.Count; elementIndex++) {
            WordBlockSnapshot element = elements[elementIndex];
            string anchor = "word-" + owner + "-block-" + elementIndex.ToString("D4", CultureInfo.InvariantCulture);
            ReaderLocation location = new ReaderLocation {
                Path = sourcePath,
                SourceBlockIndex = sourceBlockIndex,
                SourceBlockKind = element.Kind,
                BlockAnchor = anchor
            };

            if (element is WordParagraphSnapshot paragraph) {
                int? headingLevel = ResolveWordHeadingLevel(paragraph);
                if (headingLevel.HasValue) {
                    for (int headingIndex = headingStack.Count - 1; headingIndex >= 0; headingIndex--) {
                        if (headingStack[headingIndex].Level >= headingLevel.Value) headingStack.RemoveAt(headingIndex);
                    }
                    headingStack.Add((headingLevel.Value, string.IsNullOrWhiteSpace(paragraph.Text) ? "Heading " + headingLevel.Value.ToString(CultureInfo.InvariantCulture) : paragraph.Text));
                }
                location.HeadingPath = BuildWordHeadingPath(headingStack);
                string kind = headingLevel.HasValue ? "heading" : paragraph.IsListItem ? "list-item" : "paragraph";
                location.SourceBlockKind = kind;
                blocks.Add(new OfficeDocumentBlock {
                    Id = anchor,
                    Kind = kind,
                    Text = paragraph.Text,
                    Level = headingLevel ?? paragraph.ListLevel,
                    Marker = paragraph.IsListItem ? (paragraph.IsOrderedList == true ? "1." : "-") : null,
                    Location = location
                });
                AddWordParagraphLinks(paragraph, location, links, ref linkIndex);
            } else if (element is WordTableSnapshot table) {
                location.HeadingPath = BuildWordHeadingPath(headingStack);
                ReaderTable mapped = MapWordTable(table, location, tableIndex++, options.MaxTableRows);
                blocks.Add(new OfficeDocumentBlock {
                    Id = anchor,
                    Kind = "table",
                    Text = DocumentReaderEngine.BuildRichTableText(mapped),
                    Location = location
                });
                tables.Add(mapped);
                foreach (WordTableRowSnapshot row in table.Rows) {
                    foreach (WordTableCellSnapshot cell in row.Cells) {
                        foreach (WordParagraphSnapshot cellParagraph in cell.Paragraphs) {
                            AddWordParagraphLinks(cellParagraph, location, links, ref linkIndex);
                        }
                    }
                }
            }
            sourceBlockIndex++;
        }
    }

    private static void AddWordParagraphLinks(
        WordParagraphSnapshot paragraph,
        ReaderLocation ownerLocation,
        List<OfficeDocumentLink> links,
        ref int linkIndex) {
        for (int runIndex = 0; runIndex < paragraph.Runs.Count; runIndex++) {
            WordRunSnapshot run = paragraph.Runs[runIndex];
            if (!run.IsHyperlink || (string.IsNullOrWhiteSpace(run.HyperlinkUri) && string.IsNullOrWhiteSpace(run.HyperlinkAnchor))) continue;
            links.Add(new OfficeDocumentLink {
                Id = "word-link-" + linkIndex.ToString("D4", CultureInfo.InvariantCulture),
                Kind = string.IsNullOrWhiteSpace(run.HyperlinkUri) ? "internal" : "uri",
                Uri = run.HyperlinkUri,
                DestinationName = run.HyperlinkAnchor,
                Text = run.Text,
                Location = CloneWordLocation(ownerLocation, "hyperlink", ownerLocation.BlockAnchor + "-link-" + runIndex.ToString("D4", CultureInfo.InvariantCulture))
            });
            linkIndex++;
        }
    }

    private static ReaderTable MapWordTable(WordTableSnapshot table, ReaderLocation location, int tableIndex, int maxRows) {
        ReaderLocation tableLocation = CloneWordLocation(location, "table", location.BlockAnchor);
        tableLocation.TableIndex = tableIndex;
        return WordTableProjection.Map(table, tableLocation, tableIndex, maxRows);
    }

    private static int? ResolveWordHeadingLevel(WordParagraphSnapshot paragraph) {
        string style = paragraph.StyleName ?? paragraph.StyleId ?? string.Empty;
        if (style.IndexOf("heading", StringComparison.OrdinalIgnoreCase) < 0) return null;
        for (int i = style.Length - 1; i >= 0; i--) {
            if (style[i] >= '1' && style[i] <= '9') return style[i] - '0';
        }
        return 1;
    }

    private static string? BuildWordHeadingPath(IReadOnlyList<(int Level, string Text)> headings) {
        return ReaderHeadingPath.Combine(headings.Select(static heading => heading.Text));
    }

    private static ReaderLocation CloneWordLocation(ReaderLocation source, string kind, string? anchor) {
        return new ReaderLocation {
            Path = source.Path,
            SourceBlockIndex = source.SourceBlockIndex,
            HeadingPath = source.HeadingPath,
            SourceBlockKind = kind,
            BlockAnchor = anchor
        };
    }

    private static OfficeDocumentMetadataEntry BuildCountMetadataEntry(string id, string category, string name, int count) {
        return new OfficeDocumentMetadataEntry {
            Id = id,
            Category = category,
            Name = name,
            Value = count.ToString(CultureInfo.InvariantCulture),
            ValueType = "count"
        };
    }
}
