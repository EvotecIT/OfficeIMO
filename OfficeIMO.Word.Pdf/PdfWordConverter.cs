using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    internal static partial class PdfWordConverter {
        private const string ConverterName = "OfficeIMO.Word.Pdf";
        private const double DefaultEditableTableFontSizePoints = 11D;

        public static WordDocument Convert(PdfCore.PdfDocumentReadResult source, PdfToWordOptions? options) {
            if (source == null) {
                throw new ArgumentNullException(nameof(source));
            }

            PdfToWordOptions readOptions = options ?? new PdfToWordOptions();
            readOptions.CancellationToken.ThrowIfCancellationRequested();
            WordDocument document = WordDocument.Create();
            try {
                ImportInto(source, document, readOptions);
                return document;
            } catch {
                document.Dispose();
                throw;
            }
        }

        public static void ImportInto(PdfCore.PdfDocumentReadResult source, WordDocument target, PdfToWordOptions options) {
            if (source == null) {
                throw new ArgumentNullException(nameof(source));
            }

            if (target == null) {
                throw new ArgumentNullException(nameof(target));
            }

            if (options == null) {
                throw new ArgumentNullException(nameof(options));
            }

            if (options.IncludeMetadata) {
                CopyMetadata(source.Metadata, target);
            }
            ReportDocumentReconstructionBoundaries(source, options);

            bool emittedContent = false;
            WordList? bulletList = null;
            WordList? numberedList = null;
            ImportNavigationMap navigation = BuildNavigationMap(source, options);

            for (int pageIndex = 0; pageIndex < source.Pages.Count; pageIndex++) {
                options.CancellationToken.ThrowIfCancellationRequested();
                PdfCore.PdfLogicalPage page = source.Pages[pageIndex];
                ReportPageReconstructionBoundaries(page, options);
                List<ImportItem> items = BuildImportItems(page, options, navigation);
                bool hasNavigationAnchor = navigation.HasAnchorsForPage(page.PageNumber);
                bool sourcePageSizeApplied = false;
                if (pageIndex > 0 && options.PreservePageBreaks && (items.Count > 0 || options.IncludeEmptyPages || hasNavigationAnchor)) {
                    if (options.PreserveSourcePageSize) {
                        sourcePageSizeApplied = ConfigureEditablePageSection(target.AddSection(WordSectionBreakType.NextPage), page, options);
                    } else {
                        target.AddPageBreak();
                    }
                } else if (pageIndex == 0 && options.PreserveSourcePageSize) {
                    sourcePageSizeApplied = ConfigureEditablePageSection(target.Sections[0], page, options);
                }
                double typographyScale = GetEditableTypographyScale(page, sourcePageSizeApplied);

                if (AddNavigationBookmarks(target, page, navigation)) {
                    emittedContent = true;
                }

                if (items.Count == 0) {
                    if (options.IncludeEmptyPages) {
                        target.AddParagraph();
                        emittedContent = true;
                    }

                    continue;
                }

                items.Sort(CompareImportItems);
                for (int itemIndex = 0; itemIndex < items.Count; itemIndex++) {
                    options.CancellationToken.ThrowIfCancellationRequested();
                    ImportItem item = items[itemIndex];
                    bool itemEmitted = true;
                    switch (item.Kind) {
                        case ImportItemKind.Heading:
                            AddHeading(target, item.Heading!, item.Link, item.LinkText, options, navigation, typographyScale);
                            break;
                        case ImportItemKind.Paragraph:
                            AddParagraph(target, item.Paragraph!, item.Link, item.LinkText, options, navigation, typographyScale);
                            break;
                        case ImportItemKind.TextBlock:
                            AddTextBlock(target, item.TextBlock!, item.Link, item.LinkText, options, navigation, typographyScale);
                            break;
                        case ImportItemKind.ListItem:
                            AddListItem(target, item.ListItem!, ref bulletList, ref numberedList, options, typographyScale);
                            break;
                        case ImportItemKind.Table:
                            AddTable(target, item.TableExtraction!, options, typographyScale);
                            break;
                        case ImportItemKind.Image:
                            itemEmitted = AddImage(target, page, item.Image!, item.ImagePlacement, sourcePageSizeApplied, options);
                            break;
                        case ImportItemKind.FormWidget:
                            AddFormWidgetPlaceholder(target, item.FormWidget!, options);
                            break;
                        case ImportItemKind.Link:
                            AddStandaloneLink(target, item.Link!, item.LinkText, options, navigation);
                            break;
                    }

                    emittedContent |= itemEmitted;
                }
            }

            ReportNonReconstructedLinks(source, options, navigation);
            options.CancellationToken.ThrowIfCancellationRequested();
            if (!emittedContent) {
                target.AddParagraph(string.IsNullOrWhiteSpace(options.EmptyDocumentMessage)
                    ? "No supported PDF content detected."
                    : options.EmptyDocumentMessage);
            }
        }

        private static List<ImportItem> BuildImportItems(PdfCore.PdfLogicalPage page, PdfToWordOptions options, ImportNavigationMap navigation) {
            var items = new List<ImportItem>();
            var consumedLinks = new HashSet<PdfCore.PdfLogicalLinkAnnotation>();
            IReadOnlyDictionary<(PdfCore.PdfLogicalReadingOrderKind Kind, int SourceIndex, int PlacementIndex), int> readingOrder =
                BuildReadingOrder(page, options.UseSharedPageReadingOrder);
            int sequence = 0;

            if (options.ImportHeadings) {
                for (int i = 0; i < page.Headings.Count; i++) {
                    PdfCore.PdfLogicalHeading heading = page.Headings[i];
                    PdfCore.PdfLogicalLinkAnnotation? link = FindOverlappingImportableLink(page, heading.Line, options, navigation, consumedLinks);
                    items.Add(ImportItem.ForHeading(heading, heading.Line.BaselineY, sequence++, GetReadingOrder(readingOrder, PdfCore.PdfLogicalReadingOrderKind.Heading, i), link, link == null ? null : heading.Text));
                }
            }

            if (options.ImportParagraphs) {
                for (int i = 0; i < page.Paragraphs.Count; i++) {
                    PdfCore.PdfLogicalParagraph paragraph = page.Paragraphs[i];
                    PdfCore.PdfLogicalLinkAnnotation? link = FindOverlappingImportableLink(page, paragraph, options, navigation, consumedLinks);
                    items.Add(ImportItem.ForParagraph(paragraph, paragraph.YTop, sequence++, GetReadingOrder(readingOrder, PdfCore.PdfLogicalReadingOrderKind.Paragraph, i), link, link == null ? null : paragraph.Text));
                }

                for (int i = 0; i < page.TextBlocks.Count; i++) {
                    PdfCore.PdfLogicalTextBlock block = page.TextBlocks[i];
                    if (block.Kind is not (PdfCore.PdfLogicalElementKind.Header or
                        PdfCore.PdfLogicalElementKind.Footer or
                        PdfCore.PdfLogicalElementKind.Caption or
                        PdfCore.PdfLogicalElementKind.Footnote)) {
                        continue;
                    }

                    PdfCore.PdfLogicalLinkAnnotation? link = FindOverlappingImportableLink(page, block, options, navigation, consumedLinks);
                    items.Add(ImportItem.ForTextBlock(
                        block,
                        block.BaselineY,
                        sequence++,
                        GetReadingOrder(readingOrder, PdfCore.PdfLogicalReadingOrderKind.TextBlock, i),
                        link,
                        link == null ? null : block.Text));
                }
            }

            if (options.ImportLists) {
                for (int i = 0; i < page.ListItems.Count; i++) {
                    PdfCore.PdfLogicalListItem listItem = page.ListItems[i];
                    items.Add(ImportItem.ForListItem(listItem, listItem.Line.BaselineY, sequence++, GetReadingOrder(readingOrder, PdfCore.PdfLogicalReadingOrderKind.ListItem, i)));
                }
            }

            if (options.ImportTables) {
                IReadOnlyList<PdfCore.PdfLogicalTableExtraction> tables = PdfCore.PdfLogicalTableAnalysis.ExtractTables(page, options.MaxTableRows);
                for (int i = 0; i < tables.Count; i++) {
                    PdfCore.PdfLogicalTableExtraction table = tables[i];
                    items.Add(ImportItem.ForTable(table, table.Table.YTop, sequence++, GetReadingOrder(readingOrder, PdfCore.PdfLogicalReadingOrderKind.Table, i)));
                }
            }

            AddLinkItems(page, options, navigation, consumedLinks, items, readingOrder, ref sequence);

            if (options.ImportImages || options.IncludeImagePlaceholders) {
                for (int i = 0; i < page.Images.Count; i++) {
                    PdfCore.PdfLogicalImage image = page.Images[i];
                    if (image.Placements.Count == 0) {
                        if (ShouldQueueImage(page, image, null, options)) {
                            items.Add(ImportItem.ForImage(image, null, GetImageSortY(image), sequence++, GetReadingOrder(readingOrder, PdfCore.PdfLogicalReadingOrderKind.Image, i, -1)));
                        }
                        continue;
                    }

                    for (int placementIndex = 0; placementIndex < image.Placements.Count; placementIndex++) {
                        PdfCore.PdfImagePlacement placement = image.Placements[placementIndex];
                        if (ShouldQueueImage(page, image, placement, options)) {
                            items.Add(ImportItem.ForImage(image, placement, placement.Y + placement.Height, sequence++, GetReadingOrder(readingOrder, PdfCore.PdfLogicalReadingOrderKind.Image, i, placementIndex)));
                        }
                    }
                }
            } else {
                int visibleImageCount = page.Images.Count(image =>
                    PdfCore.PdfImagePlacementImportPolicy.HasVisiblePlacement(page, image));
                if (visibleImageCount > 0) {
                    AddWarning(
                        options,
                        "PdfImageSkipped",
                        "Page " + page.PageNumber.ToString(CultureInfo.InvariantCulture) + "/Image",
                        "PDF image content was not imported because IncludeImagePlaceholders is false.",
                        PdfCore.PdfConversionWarningSeverity.Warning,
                        OfficeConversionLossKind.Omission,
                        new Dictionary<string, string> {
                            ["ImageCount"] = visibleImageCount.ToString(CultureInfo.InvariantCulture)
                        });
                }
            }

            if (options.IncludeFormFieldPlaceholders) {
                for (int i = 0; i < page.FormWidgets.Count; i++) {
                    PdfCore.PdfLogicalFormWidget widget = page.FormWidgets[i];
                    items.Add(ImportItem.ForFormWidget(widget, widget.Y2, sequence++, GetReadingOrder(readingOrder, PdfCore.PdfLogicalReadingOrderKind.FormWidget, i)));
                    AddWarning(
                        options,
                        "PdfFormWidgetPlaceholder",
                        "Page " + widget.PageNumber.ToString(CultureInfo.InvariantCulture) + "/FormWidget",
                        "PDF form widget content is represented as editable Word placeholder text; interactive form reconstruction is not part of the semantic import contract.",
                        PdfCore.PdfConversionWarningSeverity.Warning,
                        OfficeConversionLossKind.Omission,
                        new Dictionary<string, string> {
                            ["FieldName"] = widget.FieldName ?? string.Empty,
                            ["FieldType"] = widget.FieldType ?? string.Empty
                        });
                }
            } else if (page.FormWidgets.Count > 0) {
                AddWarning(
                    options,
                    "PdfFormWidgetSkipped",
                    "Page " + page.PageNumber.ToString(CultureInfo.InvariantCulture) + "/FormWidget",
                    "PDF form widgets were not imported because IncludeFormFieldPlaceholders is false.",
                    PdfCore.PdfConversionWarningSeverity.Warning,
                    OfficeConversionLossKind.Omission,
                    new Dictionary<string, string> {
                        ["FormWidgetCount"] = page.FormWidgets.Count.ToString(CultureInfo.InvariantCulture)
                    });
            }

            return items;
        }

        private static void AddLinkItems(
            PdfCore.PdfLogicalPage page,
            PdfToWordOptions options,
            ImportNavigationMap navigation,
            HashSet<PdfCore.PdfLogicalLinkAnnotation> consumedLinks,
            List<ImportItem> items,
            IReadOnlyDictionary<(PdfCore.PdfLogicalReadingOrderKind Kind, int SourceIndex, int PlacementIndex), int> readingOrder,
            ref int sequence) {
            if (!options.ImportUriLinks && !options.ImportInternalLinks) {
                int uriLinkCount = page.Links.Count(link => !string.IsNullOrWhiteSpace(link.Uri));
                int internalLinkCount = page.Links.Count(link => link.IsInternalDestinationLink);
                if (uriLinkCount > 0) {
                    AddWarning(
                        options,
                        "PdfUriLinkSkipped",
                        "Page " + page.PageNumber.ToString(CultureInfo.InvariantCulture) + "/LinkAnnotation",
                        "PDF URI link annotations were not imported because ImportUriLinks is false.",
                        PdfCore.PdfConversionWarningSeverity.Information,
                        OfficeConversionLossKind.Omission,
                        new Dictionary<string, string> {
                            ["LinkCount"] = uriLinkCount.ToString(CultureInfo.InvariantCulture)
                        });
                }

                if (internalLinkCount > 0) {
                    AddWarning(
                        options,
                        "PdfInternalLinkSkipped",
                        "Page " + page.PageNumber.ToString(CultureInfo.InvariantCulture) + "/LinkAnnotation",
                        "PDF internal link annotations were not imported because ImportInternalLinks is false.",
                        PdfCore.PdfConversionWarningSeverity.Information,
                        OfficeConversionLossKind.Omission,
                        new Dictionary<string, string> {
                            ["LinkCount"] = internalLinkCount.ToString(CultureInfo.InvariantCulture)
                        });
                }

                return;
            }

            for (int i = 0; i < page.Links.Count; i++) {
                PdfCore.PdfLogicalLinkAnnotation link = page.Links[i];
                if (consumedLinks.Contains(link)) {
                    continue;
                }

                if (!TryResolveWordLinkTarget(link, options, navigation, out _)) {
                    if (!string.IsNullOrWhiteSpace(link.Uri)) {
                        ReportSkippedUriLink(link, options);
                    } else if (link.IsInternalDestinationLink) {
                        ReportSkippedInternalLink(link, options);
                    }

                    continue;
                }

                string displayText = GetOverlappingText(page, link);
                if (string.IsNullOrWhiteSpace(displayText)) {
                    displayText = GetLinkDisplayText(link, null);
                }

                consumedLinks.Add(link);
                items.Add(ImportItem.ForLink(link, GetLinkSortY(link), sequence++, GetReadingOrder(readingOrder, PdfCore.PdfLogicalReadingOrderKind.Link, i), displayText));
            }
        }

        private static PdfCore.PdfLogicalLinkAnnotation? FindOverlappingImportableLink(
            PdfCore.PdfLogicalPage page,
            PdfCore.PdfLogicalTextBlock textBlock,
            PdfToWordOptions options,
            ImportNavigationMap navigation,
            HashSet<PdfCore.PdfLogicalLinkAnnotation> consumedLinks) {
            if (!options.ImportUriLinks && !options.ImportInternalLinks) {
                return null;
            }

            for (int i = 0; i < page.Links.Count; i++) {
                PdfCore.PdfLogicalLinkAnnotation link = page.Links[i];
                if (consumedLinks.Contains(link) || !TryResolveWordLinkTarget(link, options, navigation, out _)) {
                    continue;
                }

                if (OverlapsTextBlock(link, textBlock)) {
                    consumedLinks.Add(link);
                    return link;
                }
            }

            return null;
        }

        private static PdfCore.PdfLogicalLinkAnnotation? FindOverlappingImportableLink(
            PdfCore.PdfLogicalPage page,
            PdfCore.PdfLogicalParagraph paragraph,
            PdfToWordOptions options,
            ImportNavigationMap navigation,
            HashSet<PdfCore.PdfLogicalLinkAnnotation> consumedLinks) {
            if (!options.ImportUriLinks && !options.ImportInternalLinks) {
                return null;
            }

            for (int lineIndex = 0; lineIndex < paragraph.Lines.Count; lineIndex++) {
                PdfCore.PdfLogicalLinkAnnotation? link = FindOverlappingImportableLink(page, paragraph.Lines[lineIndex], options, navigation, consumedLinks);
                if (link != null) {
                    return link;
                }
            }

            return null;
        }

        private static bool OverlapsTextBlock(PdfCore.PdfLogicalLinkAnnotation link, PdfCore.PdfLogicalTextBlock textBlock) {
            const double tolerance = 2D;
            bool yOverlaps = textBlock.BaselineY >= link.Y1 - tolerance && textBlock.BaselineY <= link.Y2 + tolerance;
            bool xOverlaps = Math.Min(link.X2, textBlock.XEnd) - Math.Max(link.X1, textBlock.XStart) > tolerance;
            return yOverlaps && xOverlaps;
        }

        private static string GetOverlappingText(PdfCore.PdfLogicalPage page, PdfCore.PdfLogicalLinkAnnotation link) {
            return string.Join(" ", page.TextBlocks
                .Where(textBlock => OverlapsTextBlock(link, textBlock))
                .OrderByDescending(textBlock => textBlock.BaselineY)
                .ThenBy(textBlock => textBlock.XStart)
                .Select(textBlock => textBlock.Text)
                .Where(text => !string.IsNullOrWhiteSpace(text)));
        }

        private static string GetLinkDisplayText(PdfCore.PdfLogicalLinkAnnotation link, string? preferredText) {
            if (!string.IsNullOrWhiteSpace(preferredText)) {
                return preferredText!;
            }

            if (!string.IsNullOrWhiteSpace(link.Contents)) {
                return link.Contents!;
            }

            if (link.IsInternalDestinationLink) {
                return GetInternalLinkDisplayText(link);
            }

            return string.IsNullOrWhiteSpace(link.Uri) ? "PDF link" : link.Uri!;
        }

        private static double GetLinkSortY(PdfCore.PdfLogicalLinkAnnotation link) => link.Y2;

        private static bool TryCreateWordHyperlinkUri(
            PdfCore.PdfLogicalLinkAnnotation link,
            PdfToWordOptions options,
            out Uri? uri) {
            uri = null;
            string? target = link.Uri;
            if (string.IsNullOrWhiteSpace(target) || !Uri.TryCreate(target, UriKind.Absolute, out Uri? parsed)) {
                return false;
            }

            if (!options.AllowedHyperlinkUriSchemes.Contains(parsed.Scheme)) {
                return false;
            }

            uri = parsed;
            return true;
        }

        private static void ReportSkippedUriLink(PdfCore.PdfLogicalLinkAnnotation link, PdfToWordOptions options) {
            AddWarning(
                options,
                "PdfUriLinkSkippedUnsafe",
                "Page " + link.PageNumber.ToString(CultureInfo.InvariantCulture) + "/LinkAnnotation",
                "PDF URI link annotation was kept inert because it is not an absolute URI with an allowed Word hyperlink scheme.",
                PdfCore.PdfConversionWarningSeverity.Warning,
                OfficeConversionLossKind.Omission,
                new Dictionary<string, string> {
                    ["Uri"] = link.Uri ?? string.Empty
                });
        }

        private static void ReportSkippedInternalLink(PdfCore.PdfLogicalLinkAnnotation link, PdfToWordOptions options) {
            AddWarning(
                options,
                "PdfInternalLinkNotReconstructed",
                "Page " + link.PageNumber.ToString(CultureInfo.InvariantCulture) + "/LinkAnnotation",
                "PDF internal link annotation could not be resolved to an imported Word bookmark.",
                PdfCore.PdfConversionWarningSeverity.Information,
                OfficeConversionLossKind.Omission,
                new Dictionary<string, string> {
                    ["DestinationName"] = link.DestinationName ?? string.Empty,
                    ["DestinationPageNumber"] = link.DestinationPageNumber?.ToString(CultureInfo.InvariantCulture) ?? string.Empty
                });
        }

        private static int CompareImportItems(ImportItem left, ImportItem right) {
            if (left.ReadingOrderIndex.HasValue && right.ReadingOrderIndex.HasValue) {
                int orderComparison = left.ReadingOrderIndex.Value.CompareTo(right.ReadingOrderIndex.Value);
                if (orderComparison != 0) return orderComparison;
            } else if (left.ReadingOrderIndex.HasValue != right.ReadingOrderIndex.HasValue) {
                return left.ReadingOrderIndex.HasValue ? -1 : 1;
            }
            int yComparison = right.Y.CompareTo(left.Y);
            return yComparison != 0 ? yComparison : left.Sequence.CompareTo(right.Sequence);
        }

        private static IReadOnlyDictionary<(PdfCore.PdfLogicalReadingOrderKind Kind, int SourceIndex, int PlacementIndex), int> BuildReadingOrder(
            PdfCore.PdfLogicalPage page,
            bool enabled) {
            if (!enabled) return new Dictionary<(PdfCore.PdfLogicalReadingOrderKind Kind, int SourceIndex, int PlacementIndex), int>();
            return PdfCore.PdfLogicalReadingOrderAnalysis.Analyze(
                page,
                PdfCore.PdfLogicalReadingOrderScope.PageContent).ToDictionary(
                static item => (item.Kind, item.SourceIndex, item.PlacementIndex),
                static item => item.OrderIndex);
        }

        private static int? GetReadingOrder(
            IReadOnlyDictionary<(PdfCore.PdfLogicalReadingOrderKind Kind, int SourceIndex, int PlacementIndex), int> readingOrder,
            PdfCore.PdfLogicalReadingOrderKind kind,
            int sourceIndex,
            int placementIndex = -1) => readingOrder.TryGetValue((kind, sourceIndex, placementIndex), out int index) ? index : null;

        private static void AddHeading(
            WordDocument document,
            PdfCore.PdfLogicalHeading heading,
            PdfCore.PdfLogicalLinkAnnotation? link,
            string? linkText,
            PdfToWordOptions options,
            ImportNavigationMap navigation,
            double typographyScale) {
            WordParagraph paragraph = link == null
                ? AddStyledParagraph(document, new[] { heading.Line }, options, typographyScale)
                : AddHyperlinkParagraph(
                    document,
                    link,
                    string.IsNullOrWhiteSpace(linkText) ? heading.Text : linkText!,
                    options,
                    navigation,
                    heading.FontSize > 0D ? heading.FontSize * typographyScale : null);
            paragraph.SetStyle(MapHeadingStyle(heading.Level));
            paragraph.KeepWithNext = true;
            if (heading.FontSize > 0) {
                paragraph.FontSizePoints = Math.Max(0.5D, heading.FontSize * typographyScale);
            }
        }

        private static void AddParagraph(
            WordDocument document,
            PdfCore.PdfLogicalParagraph paragraph,
            PdfCore.PdfLogicalLinkAnnotation? link,
            string? linkText,
            PdfToWordOptions options,
            ImportNavigationMap navigation,
            double typographyScale) {
            if (link == null) {
                AddStyledParagraph(document, paragraph.Lines, options, typographyScale);
                return;
            }

            AddHyperlinkParagraph(
                document,
                link,
                string.IsNullOrWhiteSpace(linkText) ? paragraph.Text : linkText!,
                options,
                navigation,
                GetFirstPositiveFontSize(paragraph.Lines) * typographyScale);
        }

        private static void AddTextBlock(
            WordDocument document,
            PdfCore.PdfLogicalTextBlock block,
            PdfCore.PdfLogicalLinkAnnotation? link,
            string? linkText,
            PdfToWordOptions options,
            ImportNavigationMap navigation,
            double typographyScale) {
            if (link == null) {
                AddStyledParagraph(document, new[] { block }, options, typographyScale);
                return;
            }

            AddHyperlinkParagraph(
                document,
                link,
                string.IsNullOrWhiteSpace(linkText) ? block.Text : linkText!,
                options,
                navigation,
                GetFirstPositiveFontSize(new[] { block }) * typographyScale);
        }

        private static WordParagraph AddStyledParagraph(
            WordDocument document,
            IReadOnlyList<PdfCore.PdfLogicalTextBlock> lines,
            PdfToWordOptions options,
            double typographyScale) {
            WordParagraph paragraph = document.AddParagraph();
            for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
                if (lineIndex > 0) {
                    paragraph.AddText(" ");
                }

                PdfCore.PdfLogicalTextBlock line = lines[lineIndex];
                AppendStyledRuns(paragraph, line.Runs, line.Text, line.FontSize, typographyScale);
            }

            ApplySourceParagraphSpacing(paragraph, options);

            return paragraph;
        }

        private static void AppendStyledRuns(
            WordParagraph paragraph,
            IReadOnlyList<PdfCore.PdfLogicalTextRun> runs,
            string fallbackText,
            double fallbackFontSize,
            double typographyScale) {
            if (runs.Count == 0) {
                WordParagraph fallbackRun = paragraph.AddText(fallbackText);
                if (fallbackFontSize > 0D) {
                    fallbackRun.FontSizePoints = Math.Max(0.5D, fallbackFontSize * typographyScale);
                }
                return;
            }

            for (int runIndex = 0; runIndex < runs.Count; runIndex++) {
                PdfCore.PdfLogicalTextRun source = runs[runIndex];
                WordParagraph run = paragraph.AddText(source.Text);
                if (source.IsBold) run.SetBold();
                if (source.IsItalic) run.SetItalic();
                double fontSize = source.FontSize > 0D ? source.FontSize : fallbackFontSize;
                if (fontSize > 0D) {
                    run.FontSizePoints = Math.Max(0.5D, fontSize * typographyScale);
                }
                if (source.Color.HasValue && source.Color.Value.A > 0) {
                    run.SetColorHex(source.Color.Value.ToRgbHex());
                }
            }
        }

        private static void AddStandaloneLink(
            WordDocument document,
            PdfCore.PdfLogicalLinkAnnotation link,
            string? linkText,
            PdfToWordOptions options,
            ImportNavigationMap navigation) {
            AddHyperlinkParagraph(document, link, GetLinkDisplayText(link, linkText), options, navigation);
        }

        private static WordParagraph AddHyperlinkParagraph(
            WordDocument document,
            PdfCore.PdfLogicalLinkAnnotation link,
            string text,
            PdfToWordOptions options,
            ImportNavigationMap navigation,
            double? fontSizePoints = null) {
            if (!TryResolveWordLinkTarget(link, options, navigation, out WordLinkTarget target)) {
                WordParagraph fallback = document.AddParagraph(text);
                ApplySourceParagraphSpacing(fallback, options);
                return fallback;
            }

            WordParagraph paragraph = document.AddParagraph();
            ApplySourceParagraphSpacing(paragraph, options);
            WordParagraph hyperlink;
            if (target.IsUri) {
                hyperlink = paragraph.AddHyperLink(text, target.Uri!, addStyle: true, tooltip: "Imported PDF link from page " + link.PageNumber.ToString(CultureInfo.InvariantCulture));
                if (fontSizePoints > 0D) hyperlink.FontSizePoints = Math.Max(0.5D, fontSizePoints.Value);
                AddWarning(
                    options,
                    "PdfUriLinkReconstructed",
                    "Page " + link.PageNumber.ToString(CultureInfo.InvariantCulture) + "/LinkAnnotation",
                    "PDF URI link annotation was reconstructed as an editable Word hyperlink.",
                    PdfCore.PdfConversionWarningSeverity.Information,
                    new Dictionary<string, string> {
                        ["Uri"] = link.Uri ?? string.Empty,
                        ["Text"] = text
                    });
                return paragraph;
            }

            hyperlink = paragraph.AddHyperLink(text, target.Anchor!, addStyle: true, tooltip: "Imported PDF internal link from page " + link.PageNumber.ToString(CultureInfo.InvariantCulture));
            if (fontSizePoints > 0D) hyperlink.FontSizePoints = Math.Max(0.5D, fontSizePoints.Value);
            AddWarning(
                options,
                "PdfInternalLinkReconstructed",
                "Page " + link.PageNumber.ToString(CultureInfo.InvariantCulture) + "/LinkAnnotation",
                "PDF internal link annotation was reconstructed as an editable Word bookmark hyperlink.",
                PdfCore.PdfConversionWarningSeverity.Information,
                new Dictionary<string, string> {
                    ["Anchor"] = target.Anchor!,
                    ["Text"] = text,
                    ["DestinationName"] = link.DestinationName ?? string.Empty,
                    ["DestinationPageNumber"] = link.DestinationPageNumber?.ToString(CultureInfo.InvariantCulture) ?? string.Empty
                });
            return paragraph;
        }

        private static double? GetFirstPositiveFontSize(IReadOnlyList<PdfCore.PdfLogicalTextBlock> lines) {
            for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
                IReadOnlyList<PdfCore.PdfLogicalTextRun> runs = lines[lineIndex].Runs;
                for (int runIndex = 0; runIndex < runs.Count; runIndex++) {
                    if (runs[runIndex].FontSize > 0D) return runs[runIndex].FontSize;
                }
                if (lines[lineIndex].FontSize > 0D) return lines[lineIndex].FontSize;
            }

            return null;
        }

        private static WordParagraphStyles MapHeadingStyle(int level) {
            switch (Math.Max(1, Math.Min(6, level))) {
                case 1:
                    return WordParagraphStyles.Heading1;
                case 2:
                    return WordParagraphStyles.Heading2;
                case 3:
                    return WordParagraphStyles.Heading3;
                case 4:
                    return WordParagraphStyles.Heading4;
                case 5:
                    return WordParagraphStyles.Heading5;
                default:
                    return WordParagraphStyles.Heading6;
            }
        }

        private static void AddListItem(
            WordDocument document,
            PdfCore.PdfLogicalListItem item,
            ref WordList? bulletList,
            ref WordList? numberedList,
            PdfToWordOptions options,
            double typographyScale) {
            bool bullet = IsBulletMarker(item.Marker);
            WordList list = bullet
                ? bulletList ??= document.AddListBulleted()
                : numberedList ??= document.AddListNumbered();
            WordParagraph paragraph = list.AddItem((string?)null, Math.Max(0, item.Level - 1));
            AppendStyledRuns(paragraph, item.Runs, item.Text, item.Line.FontSize, typographyScale);
            ApplySourceParagraphSpacing(paragraph, options);
        }

        private static void ApplySourceParagraphSpacing(WordParagraph paragraph, PdfToWordOptions options) {
            if (!options.PreserveCompactSourceSpacing) {
                return;
            }

            paragraph.LineSpacingBeforePoints = 0D;
            paragraph.LineSpacingAfterPoints = 0D;
        }

        private static bool IsBulletMarker(string marker) {
            string trimmed = marker.Trim();
            if (trimmed.Length == 0) {
                return false;
            }

            return trimmed == "\u2022" ||
                trimmed == "\u25CF" ||
                trimmed == "-" ||
                trimmed == "*" ||
                trimmed == "\u00B7";
        }

        private static void AddTable(
            WordDocument document,
            PdfCore.PdfLogicalTableExtraction extraction,
            PdfToWordOptions options,
            double typographyScale) {
            PdfCore.PdfLogicalTableData data = extraction.Data;
            if (data.Truncated) {
                AddWarning(
                    options,
                    "PdfTableTruncated",
                    "Page " + extraction.PageNumber.ToString(CultureInfo.InvariantCulture) + "/Table " + (extraction.TableIndex + 1).ToString(CultureInfo.InvariantCulture),
                    "PDF table rows were truncated by MaxTableRows.",
                    PdfCore.PdfConversionWarningSeverity.Warning,
                    OfficeConversionLossKind.Omission,
                    new Dictionary<string, string> {
                        ["ImportedRowCount"] = data.Rows.Count.ToString(CultureInfo.InvariantCulture),
                        ["TotalRowCount"] = data.TotalRowCount.ToString(CultureInfo.InvariantCulture)
                    });
            }
            bool headerRowIncluded = HasHeaderRow(data);
            int columnCount = data.Columns.Count;
            int rowCount = data.Rows.Count + (headerRowIncluded ? 1 : 0);
            if (columnCount == 0 || rowCount == 0) {
                return;
            }

            WordTable table = document.AddTable(rowCount, columnCount, options.TableStyle);
            PopulateTable(table, extraction.Table, data, headerRowIncluded, options, typographyScale);
        }

        private static bool HasHeaderRow(PdfCore.PdfLogicalTableData data) {
            return data.Columns.Count > 0
                && data.Structure.HasHeaderRow
                && data.Columns.Any(column => !string.IsNullOrWhiteSpace(column));
        }

        private static void PopulateTable(
            WordTable table,
            PdfCore.PdfLogicalTable sourceTable,
            PdfCore.PdfLogicalTableData data,
            bool headerRowIncluded,
            PdfToWordOptions options,
            double typographyScale) {
            List<WordTableRow> rows = table.Rows;
            int rowOffset = headerRowIncluded ? 1 : 0;

            if (headerRowIncluded) {
                WriteRow(rows[0], data.Columns, data, alignNumericColumns: false, options, typographyScale);
                if (options.RepeatHeaderRows) {
                    rows[0].RepeatHeaderRowAtTheTopOfEachPage = true;
                }
            }

            for (int rowIndex = 0; rowIndex < data.Rows.Count; rowIndex++) {
                WriteRow(rows[rowIndex + rowOffset], data.Rows[rowIndex], data, options.AlignNumericColumns, options, typographyScale);
            }

            if (options.FitTablesToPageWidth) {
                int[] columnWeights = BuildColumnWeights(sourceTable.Columns);
                if (columnWeights.Length == data.Columns.Count) {
                    table.SetColumnWidthsPercentage(columnWeights);
                } else {
                    table.WidthType = WordTableWidthUnit.Pct;
                    table.Width = 5000;
                    table.DistributeColumnsEvenly();
                }
            }
        }

        private static int[] BuildColumnWeights(IReadOnlyList<PdfCore.PdfLogicalTableColumn> columns) {
            double[] widths = columns
                .Select(static column => Math.Abs(column.To - column.From))
                .ToArray();
            double total = widths.Where(IsFinitePositive).Sum();
            if (!IsFinitePositive(total)) {
                return Enumerable.Repeat(1, columns.Count).ToArray();
            }

            return widths
                .Select(width => IsFinitePositive(width)
                    ? Math.Max(1, (int)Math.Round(width / total * 10_000D))
                    : 1)
                .ToArray();
        }

        private static bool IsFinitePositive(double value) =>
            value > 0D && !double.IsNaN(value) && !double.IsInfinity(value);

        private static void WriteRow(
            WordTableRow row,
            IReadOnlyList<string> values,
            PdfCore.PdfLogicalTableData data,
            bool alignNumericColumns,
            PdfToWordOptions options,
            double typographyScale) {
            List<WordTableCell> cells = row.Cells;
            for (int columnIndex = 0; columnIndex < cells.Count; columnIndex++) {
                string value = columnIndex < values.Count ? values[columnIndex] : string.Empty;
                WordParagraph paragraph = cells[columnIndex].AddParagraph(value ?? string.Empty, removeExistingParagraphs: true);
                ApplySourceParagraphSpacing(paragraph, options);
                if (typographyScale != 1D) {
                    paragraph.FontSizePoints = Math.Max(0.5D, DefaultEditableTableFontSizePoints * typographyScale);
                }
                if (alignNumericColumns && data.IsNumericColumn(columnIndex)) {
                    paragraph.ParagraphAlignment = WordParagraphAlignment.Right;
                }
            }
        }

        private static void AddFormWidgetPlaceholder(WordDocument document, PdfCore.PdfLogicalFormWidget widget, PdfToWordOptions options) {
            string name = string.IsNullOrWhiteSpace(widget.FieldName) ? "(unnamed)" : widget.FieldName!;
            string type = string.IsNullOrWhiteSpace(widget.FieldType) ? "field" : widget.FieldType!;
            string value = string.IsNullOrWhiteSpace(widget.Value) ? string.Empty : " = " + widget.Value;
            WordParagraph paragraph = document.AddParagraph("[PDF form " + type + ": " + name + value + "]").SetItalic();
            ApplySourceParagraphSpacing(paragraph, options);
        }

    }
}
