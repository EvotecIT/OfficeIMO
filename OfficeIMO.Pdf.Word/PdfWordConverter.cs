using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Word {
    internal static partial class PdfWordConverter {
        private const string ConverterName = "OfficeIMO.Pdf.Word";
        private const double DefaultEditableTableFontSizePoints = 11D;
        private const double MinimumEditableFontSizePoints = 0.5D;
        private const double MaximumEditableFontSizePoints = 4000D;

        public static WordDocument Convert(PdfCore.PdfDocumentReadResult source, WordPdfImportOptions? options) {
            if (source == null) {
                throw new ArgumentNullException(nameof(source));
            }

            WordPdfImportOptions readOptions = options ?? new WordPdfImportOptions();
            readOptions.CancellationToken.ThrowIfCancellationRequested();
            PdfCore.PdfLogicalContentProjection.RequireComplete(source, readOptions.CancellationToken);
            WordDocument document = WordDocument.Create();
            try {
                ImportInto(source, document, readOptions);
                return document;
            } catch {
                document.Dispose();
                throw;
            }
        }

        public static void ImportInto(PdfCore.PdfDocumentReadResult source, WordDocument target, WordPdfImportOptions options) {
            if (source == null) {
                throw new ArgumentNullException(nameof(source));
            }

            if (target == null) {
                throw new ArgumentNullException(nameof(target));
            }

            if (options == null) {
                throw new ArgumentNullException(nameof(options));
            }
            if (options.MaxImageTextOverlapComparisons <= 0)
                throw new ArgumentOutOfRangeException(nameof(options.MaxImageTextOverlapComparisons));

            if (options.IncludeMetadata) {
                CopyMetadata(source.Metadata, target);
            }
            ReportDocumentReconstructionBoundaries(source, options);

            bool emittedContent = false;
            WordList? bulletList = null;
            WordList? numberedList = null;
            ImportNavigationMap navigation = BuildNavigationMap(source, options);
            var imageOverlapBudget = new ImageOverlapBudget(options);
            var pageItems = new List<ImportItem>[source.Pages.Count];
            for (int pageIndex = 0; pageIndex < source.Pages.Count; pageIndex++) {
                options.CancellationToken.ThrowIfCancellationRequested();
                PdfCore.PdfLogicalPage page = source.Pages[pageIndex];
                ReportPageReconstructionBoundaries(page, options);
                List<ImportItem> pageImportItems = BuildImportItems(page, options, navigation);
                pageImportItems.Sort(CompareImportItems);
                pageItems[pageIndex] = pageImportItems;
            }
            navigation = PlaceNavigationAnchors(source, pageItems, navigation);

            for (int pageIndex = 0; pageIndex < source.Pages.Count; pageIndex++) {
                options.CancellationToken.ThrowIfCancellationRequested();
                PdfCore.PdfLogicalPage page = source.Pages[pageIndex];
                List<ImportItem> items = pageItems[pageIndex];
                var imageContext = new ImageImportContext(page, imageOverlapBudget);
                bool hasNavigationAnchor = navigation.HasAnchorsForPage(page.PageNumber);
                bool hasPageOutputCandidate = items.Count > 0 || options.IncludeEmptyPages || hasNavigationAnchor;
                WordParagraph? provisionalPageBreak = null;
                WordSection? provisionalSection = null;
                bool sourcePageSizeApplied = false;
                if (emittedContent && options.PreservePageBreaks && hasPageOutputCandidate) {
                    if (options.PreserveSourcePageSize) {
                        provisionalSection = target.AddSection(WordSectionBreakType.NextPage);
                        sourcePageSizeApplied = ConfigureEditablePageSection(provisionalSection, page, options);
                    } else {
                        provisionalPageBreak = target.AddPageBreak();
                    }
                } else if (!emittedContent && options.PreserveSourcePageSize && hasPageOutputCandidate) {
                    sourcePageSizeApplied = ConfigureEditablePageSection(target.Sections[0], page, options);
                }
                double typographyScale = GetEditableTypographyScale(page, sourcePageSizeApplied);

                bool pageEmitted = AddNavigationBookmarks(target, page, navigation);

                if (items.Count == 0) {
                    if (options.IncludeEmptyPages) {
                        target.AddParagraph();
                        pageEmitted = true;
                    }

                    if (!pageEmitted) {
                        provisionalPageBreak?.Remove();
                        provisionalSection?.RemoveSection();
                    }
                    emittedContent |= pageEmitted;
                    continue;
                }

                WordParagraph? formRowParagraph = null;
                double formRowY = 0D;
                for (int itemIndex = 0; itemIndex < items.Count; itemIndex++) {
                    options.CancellationToken.ThrowIfCancellationRequested();
                    ImportItem item = items[itemIndex];
                    bool itemEmitted = true;
                    WordParagraph? emittedParagraph = null;
                    switch (item.Kind) {
                        case ImportItemKind.Heading:
                            emittedParagraph = AddHeading(target, item.Heading!, item.Link, item.LinkText, options, navigation, typographyScale);
                            break;
                        case ImportItemKind.Paragraph:
                            emittedParagraph = AddParagraph(target, item.Paragraph!, item.Link, item.LinkText, options, navigation, typographyScale);
                            break;
                        case ImportItemKind.TextBlock:
                            emittedParagraph = AddTextBlock(target, item.TextBlock!, item.Link, item.LinkText, options, navigation, typographyScale);
                            break;
                        case ImportItemKind.ListItem:
                            emittedParagraph = AddListItem(target, item.ListItem!, ref bulletList, ref numberedList, options, typographyScale);
                            break;
                        case ImportItemKind.Table:
                            AddTable(target, item.TableExtraction!, options, typographyScale);
                            break;
                        case ImportItemKind.Image:
                            itemEmitted = AddImage(target, page, item.Image!, item.ImagePlacement, sourcePageSizeApplied, options, imageContext);
                            break;
                        case ImportItemKind.FormWidget:
                            PdfCore.PdfLogicalFormWidget widget = item.FormWidget!;
                            bool sameRow = formRowParagraph != null && Math.Abs(formRowY - widget.Y2) <= Math.Max(4D, widget.Height * 0.5D);
                            WordParagraph? widgetParagraph = AddFormWidget(target, widget, options, sameRow ? formRowParagraph : null);
                            itemEmitted = widgetParagraph != null;
                            if (widgetParagraph != null) {
                                formRowParagraph = widgetParagraph;
                                formRowY = widget.Y2;
                            }
                            break;
                        case ImportItemKind.Figure:
                            itemEmitted = AddVectorFigure(target, page, item.Figure!, options);
                            break;
                        case ImportItemKind.Link:
                            emittedParagraph = AddStandaloneLink(target, item.Link!, item.LinkText, options, navigation);
                            break;
                    }
                    if (item.Kind != ImportItemKind.FormWidget) formRowParagraph = null;
                    if (emittedParagraph != null && navigation.TryGetItemAnchor(page.PageNumber, itemIndex, out string? itemAnchor)) {
                        emittedParagraph.AddBookmark(itemAnchor!);
                    }

                    pageEmitted |= itemEmitted;
                }
                if (!pageEmitted) {
                    provisionalPageBreak?.Remove();
                    provisionalSection?.RemoveSection();
                }
                emittedContent |= pageEmitted;
            }

            ReportNonReconstructedLinks(source, options, navigation);
            options.CancellationToken.ThrowIfCancellationRequested();
            if (!emittedContent) {
                target.AddParagraph(string.IsNullOrWhiteSpace(options.EmptyDocumentMessage)
                    ? "No supported PDF content detected."
                    : options.EmptyDocumentMessage);
            }
        }

        private static List<ImportItem> BuildImportItems(PdfCore.PdfLogicalPage page, WordPdfImportOptions options, ImportNavigationMap navigation) {
            var items = new List<ImportItem>();
            var consumedLinks = new HashSet<PdfCore.PdfLogicalLinkAnnotation>();
            IReadOnlyList<PdfCore.PdfLogicalReadingOrderItem> projection = PdfCore.PdfLogicalReadingOrderAnalysis.Analyze(
                page, PdfCore.PdfLogicalReadingOrderScope.PageContent);
            var projectedLists = new HashSet<int>(projection.Where(static item => item.Kind == PdfCore.PdfLogicalReadingOrderKind.ListItem)
                .Select(static item => item.SourceIndex));
            var listOwnedLines = new HashSet<PdfCore.PdfLogicalTextBlock>(options.ImportLists
                ? projectedLists.SelectMany(index => page.ListItems[index].Lines)
                : Enumerable.Empty<PdfCore.PdfLogicalTextBlock>());
            var projectedTextBlocks = new HashSet<int>(projection.Where(static item => item.Kind == PdfCore.PdfLogicalReadingOrderKind.TextBlock)
                .Select(static item => item.SourceIndex));
            var projectedParagraphs = new HashSet<int>(projection.Where(static item => item.Kind == PdfCore.PdfLogicalReadingOrderKind.Paragraph)
                .Select(static item => item.SourceIndex));
            IReadOnlyDictionary<(PdfCore.PdfLogicalReadingOrderKind Kind, int SourceIndex, int PlacementIndex), int> readingOrder =
                BuildReadingOrder(projection, options.UseSharedPageReadingOrder);
            int sequence = 0;
            // Labels drawn as part of a chart travel with the picture of the chart.
            var figureLabels = new HashSet<PdfCore.PdfLogicalTextBlock>(options.ImportImages
                ? page.VectorFigures.SelectMany(static figure => figure.Labels)
                : Enumerable.Empty<PdfCore.PdfLogicalTextBlock>());
            var figureOrder = new Dictionary<PdfCore.PdfLogicalVectorFigure, int>();
            void NoteFigureLabel(PdfCore.PdfLogicalTextBlock line, int? order) {
                if (!order.HasValue) return;
                foreach (PdfCore.PdfLogicalVectorFigure figure in page.VectorFigures) {
                    if (!figure.Labels.Contains(line)) continue;
                    figureOrder[figure] = figureOrder.TryGetValue(figure, out int existing) ? Math.Min(existing, order.Value) : order.Value;
                }
            }

            if (options.ImportHeadings) {
                for (int i = 0; i < page.Headings.Count; i++) {
                    PdfCore.PdfLogicalHeading heading = page.Headings[i];
                    if (figureLabels.Contains(heading.Line)) {
                        NoteFigureLabel(heading.Line, GetReadingOrder(readingOrder, PdfCore.PdfLogicalReadingOrderKind.Heading, i));
                        continue;
                    }
                    PdfCore.PdfLogicalLinkAnnotation? link = FindOverlappingImportableLink(page, heading.Line, options, navigation, consumedLinks);
                    items.Add(ImportItem.ForHeading(heading, heading.Line.BaselineY, sequence++, GetReadingOrder(readingOrder, PdfCore.PdfLogicalReadingOrderKind.Heading, i), link, link == null ? null : heading.Text));
                }
            }

            if (options.ImportParagraphs) {
                for (int i = 0; i < page.Paragraphs.Count; i++) {
                    PdfCore.PdfLogicalParagraph paragraph = page.Paragraphs[i];
                    if (paragraph.Lines.Count > 0 && paragraph.Lines.All(listOwnedLines.Contains)) continue;
                    if (paragraph.Lines.Count > 0 && paragraph.Lines.All(figureLabels.Contains)) {
                        NoteFigureLabel(paragraph.Lines[0], GetReadingOrder(readingOrder, PdfCore.PdfLogicalReadingOrderKind.Paragraph, i));
                        continue;
                    }
                    if (options.ImportTables && !projectedParagraphs.Contains(i) &&
                        paragraph.Lines.Count > 0 && paragraph.Lines.All(static line => line.IsTableContent)) continue;
                    PdfCore.PdfLogicalLinkAnnotation? link = FindOverlappingImportableLink(page, paragraph, options, navigation, consumedLinks);
                    items.Add(ImportItem.ForParagraph(paragraph, paragraph.YTop, sequence++, GetReadingOrder(readingOrder, PdfCore.PdfLogicalReadingOrderKind.Paragraph, i), link, link == null ? null : paragraph.Text));
                }

                for (int i = 0; i < page.TextBlocks.Count; i++) {
                    PdfCore.PdfLogicalTextBlock block = page.TextBlocks[i];
                    if (figureLabels.Contains(block)) {
                        NoteFigureLabel(block, GetReadingOrder(readingOrder, PdfCore.PdfLogicalReadingOrderKind.TextBlock, i));
                        continue;
                    }
                    if (!projectedTextBlocks.Contains(i) && block.Kind is not (PdfCore.PdfLogicalElementKind.Header or
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
                    if (!projectedLists.Contains(i)) continue;
                    PdfCore.PdfLogicalListItem listItem = page.ListItems[i];
                    items.Add(ImportItem.ForListItem(listItem, listItem.Line.BaselineY, sequence++, GetReadingOrder(readingOrder, PdfCore.PdfLogicalReadingOrderKind.ListItem, i)));
                }
            }

            if (options.ImportTables) {
                IReadOnlyList<PdfCore.PdfLogicalTableExtraction> tables = PdfCore.PdfLogicalTableAnalysis.ExtractTables(page, options.MaxTableRows);
                for (int i = 0; i < tables.Count; i++) {
                    PdfCore.PdfLogicalTableExtraction table = tables[i];
                    // Dot-leader rows (tables of contents, price lists) are also projected as paragraphs, which
                    // become tabbed leader lines. A second, table copy of the same lines would duplicate them.
                    if (options.ImportParagraphs && IsLeaderTableProjectedAsText(page, table)) continue;
                    items.Add(ImportItem.ForTable(table, table.Table.YTop, sequence++, GetReadingOrder(readingOrder, PdfCore.PdfLogicalReadingOrderKind.Table, i)));
                }
            }

            if (options.ImportImages) {
                (double _, double pageHeight) = page.GetVisualPageSize();
                foreach (PdfCore.PdfLogicalVectorFigure figure in page.VectorFigures) {
                    int? order = figureOrder.TryGetValue(figure, out int labelOrder) ? labelOrder : null;
                    items.Add(ImportItem.ForFigure(figure, pageHeight - figure.Top, sequence++, order));
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
                        "pdf-image-skipped",
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
                        "pdf-form-widget-placeholder",
                        "Page " + widget.PageNumber.ToString(CultureInfo.InvariantCulture) + "/FormWidget",
                        "PDF form widget is represented as a Word content control holding its value; PDF field actions, calculations and appearance streams are not carried over.",
                        PdfCore.PdfConversionWarningSeverity.Warning,
                        OfficeConversionLossKind.Approximation,
                        new Dictionary<string, string> {
                            ["FieldName"] = widget.FieldName ?? string.Empty,
                            ["FieldType"] = widget.FieldType ?? string.Empty
                        });
                }
            } else if (page.FormWidgets.Count > 0) {
                AddWarning(
                    options,
                    "pdf-form-widget-skipped",
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
            WordPdfImportOptions options,
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
                        "pdf-uri-link-skipped",
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
                        "pdf-internal-link-skipped",
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
            WordPdfImportOptions options,
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
            WordPdfImportOptions options,
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
            WordPdfImportOptions options,
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

        private static void ReportSkippedUriLink(PdfCore.PdfLogicalLinkAnnotation link, WordPdfImportOptions options) {
            AddWarning(
                options,
                "pdf-uri-link-skipped-unsafe",
                "Page " + link.PageNumber.ToString(CultureInfo.InvariantCulture) + "/LinkAnnotation",
                "PDF URI link annotation was kept inert because it is not an absolute URI with an allowed Word hyperlink scheme.",
                PdfCore.PdfConversionWarningSeverity.Warning,
                OfficeConversionLossKind.Omission,
                new Dictionary<string, string> {
                    ["Uri"] = link.Uri ?? string.Empty
                });
        }

        private static void ReportSkippedInternalLink(PdfCore.PdfLogicalLinkAnnotation link, WordPdfImportOptions options) {
            AddWarning(
                options,
                "pdf-internal-link-not-reconstructed",
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
            IReadOnlyList<PdfCore.PdfLogicalReadingOrderItem> projection,
            bool enabled) {
            if (!enabled) return new Dictionary<(PdfCore.PdfLogicalReadingOrderKind Kind, int SourceIndex, int PlacementIndex), int>();
            return projection.ToDictionary(
                static item => (item.Kind, item.SourceIndex, item.PlacementIndex),
                static item => item.OrderIndex);
        }

        private static int? GetReadingOrder(
            IReadOnlyDictionary<(PdfCore.PdfLogicalReadingOrderKind Kind, int SourceIndex, int PlacementIndex), int> readingOrder,
            PdfCore.PdfLogicalReadingOrderKind kind,
            int sourceIndex,
            int placementIndex = -1) => readingOrder.TryGetValue((kind, sourceIndex, placementIndex), out int index) ? index : null;

        private static WordParagraph AddHeading(
            WordDocument document,
            PdfCore.PdfLogicalHeading heading,
            PdfCore.PdfLogicalLinkAnnotation? link,
            string? linkText,
            WordPdfImportOptions options,
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
                    heading.FontSize > 0D ? ScaleEditableFontSize(heading.FontSize, typographyScale) : null);
            paragraph.SetStyle(MapHeadingStyle(heading.Level));
            paragraph.KeepWithNext = true;
            if (heading.FontSize > 0) {
                paragraph.FontSizePoints = ScaleEditableFontSize(heading.FontSize, typographyScale);
            }
            return paragraph;
        }

        private static WordParagraph AddParagraph(
            WordDocument document,
            PdfCore.PdfLogicalParagraph paragraph,
            PdfCore.PdfLogicalLinkAnnotation? link,
            string? linkText,
            WordPdfImportOptions options,
            ImportNavigationMap navigation,
            double typographyScale) {
            if (link == null) {
                return AddStyledParagraph(document, paragraph.Lines, options, typographyScale);
            }

            return AddHyperlinkParagraph(
                document,
                link,
                string.IsNullOrWhiteSpace(linkText) ? paragraph.Text : linkText!,
                options,
                navigation,
                ScaleEditableFontSize(GetFirstPositiveFontSize(paragraph.Lines), typographyScale));
        }

        private static WordParagraph AddTextBlock(
            WordDocument document,
            PdfCore.PdfLogicalTextBlock block,
            PdfCore.PdfLogicalLinkAnnotation? link,
            string? linkText,
            WordPdfImportOptions options,
            ImportNavigationMap navigation,
            double typographyScale) {
            if (link == null) {
                return AddStyledParagraph(document, new[] { block }, options, typographyScale);
            }

            return AddHyperlinkParagraph(
                document,
                link,
                string.IsNullOrWhiteSpace(linkText) ? block.Text : linkText!,
                options,
                navigation,
                ScaleEditableFontSize(GetFirstPositiveFontSize(new[] { block }), typographyScale));
        }

        private static readonly System.Text.RegularExpressions.Regex LeaderLinePattern = new System.Text.RegularExpressions.Regex(
            @"^(?<label>.*?\S)\s*(?:\.\s*){4,}(?<value>\S{1,12})$",
            System.Text.RegularExpressions.RegexOptions.CultureInvariant,
            TimeSpan.FromSeconds(1));

        private static bool IsLeaderTableProjectedAsText(PdfCore.PdfLogicalPage page, PdfCore.PdfLogicalTableExtraction table) {
            if (!string.Equals(table.DetectionKind, "leaders", StringComparison.OrdinalIgnoreCase)) return false;
            foreach (IReadOnlyList<string> row in table.Table.Rows) {
                string label = row.Count > 0 ? row[0]?.Trim() ?? string.Empty : string.Empty;
                if (label.Length == 0) continue;
                bool projected = page.Paragraphs.Any(paragraph => !paragraph.Lines.All(static line => line.IsTableContent) &&
                        paragraph.Text.IndexOf(label, StringComparison.Ordinal) >= 0) ||
                    page.TextBlocks.Any(block => !block.IsTableContent && block.Text.IndexOf(label, StringComparison.Ordinal) >= 0);
                if (!projected) return false;
            }
            return true;
        }

        /// <summary>
        /// Writes "Title ........ 12" as "Title", a tab and "12" over a right-aligned dot-leader tab stop, so the
        /// leader stays aligned however the text reflows.
        /// </summary>
        private static bool TryAddLeaderParagraph(
            WordDocument document,
            IReadOnlyList<PdfCore.PdfLogicalTextBlock> lines,
            WordPdfImportOptions options,
            double typographyScale,
            out WordParagraph paragraph) {
            paragraph = null!;
            if (lines.Count != 1) return false;
            System.Text.RegularExpressions.Match match = LeaderLinePattern.Match(lines[0].Text.Trim());
            if (!match.Success) return false;
            WordSection section = document.Sections[document.Sections.Count - 1];
            long textWidth = (long)(section.PageSettings.Width ?? 12240U) - section.Margins.Left - section.Margins.Right;
            if (textWidth <= 720) return false;
            paragraph = document.AddParagraph();
            paragraph.AddTabStop((int)Math.Min(int.MaxValue, textWidth), WordTabAlignment.Right, WordTabLeader.Dot);
            double fontSize = lines[0].FontSize;
            PdfCore.PdfLogicalTextRun? firstRun = lines[0].Runs.Count > 0 ? lines[0].Runs[0] : null;
            WordParagraph label = paragraph.AddText(match.Groups["label"].Value);
            paragraph.AddTab();
            WordParagraph value = paragraph.AddText(match.Groups["value"].Value);
            foreach (WordParagraph run in new[] { label, value }) {
                if (fontSize > 0D) run.FontSizePoints = ScaleEditableFontSize(fontSize, typographyScale);
                if (firstRun?.IsBold == true) run.SetBold();
                string? family = PdfCore.PdfFontFamilyNames.ToFamilyName(firstRun?.BaseFont);
                if (!string.IsNullOrWhiteSpace(family)) run.SetFontFamily(family!);
            }
            ApplySourceParagraphSpacing(paragraph, options);
            return true;
        }

        private static WordParagraph AddStyledParagraph(
            WordDocument document,
            IReadOnlyList<PdfCore.PdfLogicalTextBlock> lines,
            WordPdfImportOptions options,
            double typographyScale) {
            if (TryAddLeaderParagraph(document, lines, options, typographyScale, out WordParagraph leader)) return leader;
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
                    fallbackRun.FontSizePoints = ScaleEditableFontSize(fallbackFontSize, typographyScale);
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
                    run.FontSizePoints = ScaleEditableFontSize(fontSize, typographyScale);
                }
                // Light text was drawn on a coloured panel that a reflowed paragraph does not keep; on the white
                // page it would be invisible, so it keeps the document text colour.
                if (source.Color.HasValue && source.Color.Value.A > 0 && !IsNearWhite(source.Color.Value)) {
                    run.SetColorHex(source.Color.Value.ToRgbHex());
                }
                string? fontFamily = PdfCore.PdfFontFamilyNames.ToFamilyName(source.BaseFont);
                if (!string.IsNullOrWhiteSpace(fontFamily)) {
                    run.SetFontFamily(fontFamily!);
                }
            }
        }

        private static WordParagraph AddStandaloneLink(
            WordDocument document,
            PdfCore.PdfLogicalLinkAnnotation link,
            string? linkText,
            WordPdfImportOptions options,
            ImportNavigationMap navigation) {
            return AddHyperlinkParagraph(document, link, GetLinkDisplayText(link, linkText), options, navigation);
        }

        private static WordParagraph AddHyperlinkParagraph(
            WordDocument document,
            PdfCore.PdfLogicalLinkAnnotation link,
            string text,
            WordPdfImportOptions options,
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
                if (fontSizePoints > 0D) hyperlink.FontSizePoints = BoundEditableFontSize(fontSizePoints.Value);
                AddWarning(
                    options,
                    "pdf-uri-link-reconstructed",
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
            if (fontSizePoints > 0D) hyperlink.FontSizePoints = BoundEditableFontSize(fontSizePoints.Value);
            AddWarning(
                options,
                "pdf-internal-link-reconstructed",
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

        private static double? ScaleEditableFontSize(double? fontSize, double scale) =>
            fontSize.HasValue ? ScaleEditableFontSize(fontSize.Value, scale) : null;

        private static double ScaleEditableFontSize(double fontSize, double scale) =>
            BoundEditableFontSize(fontSize * scale);

        private static double BoundEditableFontSize(double fontSize) {
            if (double.IsNaN(fontSize) || fontSize <= MinimumEditableFontSizePoints) {
                return MinimumEditableFontSizePoints;
            }

            return double.IsInfinity(fontSize) || fontSize >= MaximumEditableFontSizePoints
                ? MaximumEditableFontSizePoints
                : fontSize;
        }

        private static WordParagraphStyles MapHeadingStyle(int level) {
            switch (Math.Max(1, Math.Min(9, level))) {
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
                case 6:
                    return WordParagraphStyles.Heading6;
                case 7:
                    return WordParagraphStyles.Heading7;
                case 8:
                    return WordParagraphStyles.Heading8;
                default:
                    return WordParagraphStyles.Heading9;
            }
        }

        private static WordParagraph AddListItem(
            WordDocument document,
            PdfCore.PdfLogicalListItem item,
            ref WordList? bulletList,
            ref WordList? numberedList,
            WordPdfImportOptions options,
            double typographyScale) {
            bool bullet = IsBulletMarker(item.Marker);
            WordList list = bullet
                ? bulletList ??= document.AddListBulleted()
                : numberedList ??= document.AddListNumbered();
            WordParagraph paragraph = list.AddItem((string?)null, Math.Max(0, item.Level - 1));
            AppendStyledRuns(paragraph, item.Runs, item.Text, item.Line.FontSize, typographyScale);
            ApplyCompactSpacing(paragraph);
            return paragraph;
        }

        private static void ApplySourceParagraphSpacing(WordParagraph paragraph, WordPdfImportOptions options) {
            if (!options.PreserveCompactSourceSpacing) {
                return;
            }

            ApplyCompactSpacing(paragraph);
        }

        private static void ApplyCompactSpacing(WordParagraph paragraph) {
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
            WordPdfImportOptions options,
            double typographyScale) {
            PdfCore.PdfLogicalTableData data = extraction.Data;
            if (data.Truncated) {
                AddWarning(
                    options,
                    "pdf-table-truncated",
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
            WordPdfImportOptions options,
            double typographyScale) {
            List<WordTableRow> rows = table.Rows;
            table.RepeatHeaderRowAtTheTopOfEachPage = headerRowIncluded && options.RepeatHeaderRows;
            int rowOffset = headerRowIncluded ? 1 : 0;

            IReadOnlyList<IReadOnlyList<PdfCore.PdfLogicalTableCellStyle?>>? styles = sourceTable.CellStyles;
            // Table data keeps source rows in order; a promoted header is source row zero.
            const int sourceRowOffset = 0;
            if (headerRowIncluded) {
                WriteRow(rows[0], data.Columns, data, alignNumericColumns: false, options, typographyScale, GetRowStyles(styles, sourceRowOffset));
            }

            for (int rowIndex = 0; rowIndex < data.Rows.Count; rowIndex++) {
                WriteRow(rows[rowIndex + rowOffset], data.Rows[rowIndex], data, options.AlignNumericColumns, options, typographyScale,
                    GetRowStyles(styles, sourceRowOffset + rowOffset + rowIndex));
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

        internal static int[] BuildColumnWeights(IReadOnlyList<PdfCore.PdfLogicalTableColumn> columns) {
            bool horizontalProgression = PdfCore.PdfTableColumnGeometry.HasHorizontalProgression(
                columns,
                static column => column.VisualBounds,
                fallback: true);
            double[] widths = columns
                .Select(column => PdfCore.PdfTableColumnGeometry.GetProgressionLength(column, horizontalProgression))
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

        private static IReadOnlyList<PdfCore.PdfLogicalTableCellStyle?>? GetRowStyles(
            IReadOnlyList<IReadOnlyList<PdfCore.PdfLogicalTableCellStyle?>>? styles,
            int sourceRowIndex) =>
            styles != null && sourceRowIndex >= 0 && sourceRowIndex < styles.Count ? styles[sourceRowIndex] : null;

        private static void WriteRow(
            WordTableRow row,
            IReadOnlyList<string> values,
            PdfCore.PdfLogicalTableData data,
            bool alignNumericColumns,
            WordPdfImportOptions options,
            double typographyScale,
            IReadOnlyList<PdfCore.PdfLogicalTableCellStyle?>? styles = null) {
            List<WordTableCell> cells = row.Cells;
            for (int columnIndex = 0; columnIndex < cells.Count; columnIndex++) {
                string value = columnIndex < values.Count ? values[columnIndex] : string.Empty;
                WordParagraph paragraph = cells[columnIndex].AddParagraph(value ?? string.Empty, removeExistingParagraphs: true);
                ApplyCompactSpacing(paragraph);
                PdfCore.PdfLogicalTableCellStyle? style = styles != null && columnIndex < styles.Count ? styles[columnIndex] : null;
                if (style?.FontSize is double sourceFontSize && sourceFontSize > 0D) {
                    paragraph.FontSizePoints = ScaleEditableFontSize(sourceFontSize, typographyScale);
                } else if (typographyScale != 1D) {
                    paragraph.FontSizePoints = ScaleEditableFontSize(DefaultEditableTableFontSizePoints, typographyScale);
                }
                if (style != null) ApplyCellStyle(cells[columnIndex], paragraph, style);
                if (alignNumericColumns && data.IsNumericColumn(columnIndex)) {
                    paragraph.ParagraphAlignment = WordParagraphAlignment.Right;
                }
            }
        }

        /// <summary>
        /// Carries the source cell fill and text styling. The fill and text colour travel together so light text
        /// written on a dark source fill stays readable.
        /// </summary>
        private static void ApplyCellStyle(WordTableCell cell, WordParagraph paragraph, PdfCore.PdfLogicalTableCellStyle style) {
            if (style.FillColor.HasValue) {
                cell.ShadingFillColorHex = style.FillColor.Value.ToRgbHex();
            }
            if (paragraph.Text.Length == 0) return;
            if (style.IsBold) paragraph.SetBold();
            if (style.IsItalic) paragraph.SetItalic();
            if (style.TextColor.HasValue && style.TextColor.Value.A > 0 && !IsBlack(style.TextColor.Value)) {
                paragraph.SetColorHex(style.TextColor.Value.ToRgbHex());
            }
            string? fontFamily = PdfCore.PdfFontFamilyNames.ToFamilyName(style.FontName);
            if (!string.IsNullOrWhiteSpace(fontFamily)) paragraph.SetFontFamily(fontFamily!);
        }

        private static bool IsBlack(OfficeIMO.OfficeColor color) => color.R == 0 && color.G == 0 && color.B == 0;

        private static bool IsNearWhite(OfficeIMO.OfficeColor color) =>
            0.2126D * color.R + 0.7152D * color.G + 0.0722D * color.B >= 0.88D * 255D;

        /// <summary>
        /// Rebuilds a form widget as a Word content control: check boxes stay check boxes, choices become drop-down
        /// lists and text fields become text controls holding the current value. Widgets on the same source row share
        /// one paragraph, so a form keeps its rows instead of growing one paragraph per field.
        /// </summary>
        private static WordParagraph? AddFormWidget(
            WordDocument document,
            PdfCore.PdfLogicalFormWidget widget,
            WordPdfImportOptions options,
            WordParagraph? rowParagraph) {
            PdfCore.PdfFormField field = widget.Field;
            if (field.IsPushButton || widget.IsHidden || widget.IsInvisible) return null;
            string alias = GetFormFieldLabel(field);
            WordParagraph paragraph = rowParagraph ?? document.AddParagraph();
            if (rowParagraph == null) {
                ApplySourceParagraphSpacing(paragraph, options);
            } else {
                paragraph.AddText(" ");
            }
            if (field.IsButtonField) {
                bool isChecked = !string.IsNullOrEmpty(widget.AppearanceState) &&
                    !string.Equals(widget.AppearanceState, "Off", StringComparison.Ordinal);
                paragraph.AddCheckBox(isChecked, alias, field.Name);
            } else if (field.IsChoiceField && field.HasOptions) {
                var items = field.Options.Select(static option => option.DisplayText).Where(static text => !string.IsNullOrEmpty(text)).Distinct().ToArray();
                if (items.Length > 0) {
                    paragraph.AddDropDownList(items, alias, field.Name);
                } else {
                    paragraph.AddStructuredDocumentTag(GetFormFieldText(field), alias, field.Name);
                }
            } else {
                paragraph.AddStructuredDocumentTag(GetFormFieldText(field), alias, field.Name);
            }
            return paragraph;
        }

        private static string GetFormFieldText(PdfCore.PdfFormField field) {
            string? value = field.IsRichText && !string.IsNullOrEmpty(field.RichValuePlainText) ? field.RichValuePlainText : field.Value;
            // An empty control would be invisible; keep a short blank the reader can click into.
            return string.IsNullOrEmpty(value) ? new string('\u2002', 8) : value!;
        }

        private static string GetFormFieldLabel(PdfCore.PdfFormField field) {
            if (!string.IsNullOrWhiteSpace(field.AlternateName)) return field.AlternateName!.Trim();
            string? name = field.PartialName ?? field.Name;
            if (string.IsNullOrWhiteSpace(name)) return "Form field";
            // Generated names such as "f1_01[0]" carry no meaning for a reader beyond the field itself.
            return name!.Trim();
        }

    }
}
