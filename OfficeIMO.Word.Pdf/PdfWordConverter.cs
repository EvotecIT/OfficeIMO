using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    internal static partial class PdfWordConverter {
        private const string ConverterName = "OfficeIMO.Word.Pdf";

        public static WordDocument Convert(PdfCore.PdfLogicalDocument source, PdfWordReadOptions? options) {
            if (source == null) {
                throw new ArgumentNullException(nameof(source));
            }

            PdfWordReadOptions readOptions = options ?? new PdfWordReadOptions();
            readOptions.ResetImportState();
            WordDocument document = WordDocument.Create();
            ImportInto(source, document, readOptions);
            return document;
        }

        public static void ImportInto(PdfCore.PdfLogicalDocument source, WordDocument target, PdfWordReadOptions options) {
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

            bool emittedContent = false;
            WordList? bulletList = null;
            WordList? numberedList = null;
            ImportNavigationMap navigation = BuildNavigationMap(source, options);

            for (int pageIndex = 0; pageIndex < source.Pages.Count; pageIndex++) {
                PdfCore.PdfLogicalPage page = source.Pages[pageIndex];
                List<ImportItem> items = BuildImportItems(page, options, navigation);
                bool hasNavigationAnchor = navigation.HasAnchorsForPage(page.PageNumber);
                if (pageIndex > 0 && options.PreservePageBreaks && (items.Count > 0 || options.IncludeEmptyPages || hasNavigationAnchor)) {
                    target.AddPageBreak();
                }

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
                    ImportItem item = items[itemIndex];
                    switch (item.Kind) {
                        case ImportItemKind.Heading:
                            AddHeading(target, item.Heading!, item.Link, item.LinkText, options, navigation);
                            break;
                        case ImportItemKind.Paragraph:
                            AddParagraph(target, item.Paragraph!, item.Link, item.LinkText, options, navigation);
                            break;
                        case ImportItemKind.ListItem:
                            AddListItem(target, item.ListItem!, ref bulletList, ref numberedList);
                            break;
                        case ImportItemKind.Table:
                            AddTable(target, item.TableExtraction!, options);
                            break;
                        case ImportItemKind.Image:
                            AddImage(target, item.Image!, item.ImagePlacement, options);
                            break;
                        case ImportItemKind.FormWidget:
                            AddFormWidgetPlaceholder(target, item.FormWidget!, options);
                            break;
                        case ImportItemKind.Link:
                            AddStandaloneLink(target, item.Link!, item.LinkText, options, navigation);
                            break;
                    }

                    emittedContent = true;
                }
            }

            ReportNonReconstructedLinks(source, options, navigation);
            if (!emittedContent) {
                target.AddParagraph(string.IsNullOrWhiteSpace(options.EmptyDocumentMessage)
                    ? "No supported PDF content detected."
                    : options.EmptyDocumentMessage);
            }
        }

        private static List<ImportItem> BuildImportItems(PdfCore.PdfLogicalPage page, PdfWordReadOptions options, ImportNavigationMap navigation) {
            var items = new List<ImportItem>();
            var consumedLinks = new HashSet<PdfCore.PdfLogicalLinkAnnotation>();
            int sequence = 0;

            if (options.ImportHeadings) {
                for (int i = 0; i < page.Headings.Count; i++) {
                    PdfCore.PdfLogicalHeading heading = page.Headings[i];
                    PdfCore.PdfLogicalLinkAnnotation? link = FindOverlappingImportableLink(page, heading.Line, options, navigation, consumedLinks);
                    items.Add(ImportItem.ForHeading(heading, heading.Line.BaselineY, sequence++, link, link == null ? null : heading.Text));
                }
            }

            if (options.ImportParagraphs) {
                for (int i = 0; i < page.Paragraphs.Count; i++) {
                    PdfCore.PdfLogicalParagraph paragraph = page.Paragraphs[i];
                    PdfCore.PdfLogicalLinkAnnotation? link = FindOverlappingImportableLink(page, paragraph, options, navigation, consumedLinks);
                    items.Add(ImportItem.ForParagraph(paragraph, paragraph.YTop, sequence++, link, link == null ? null : paragraph.Text));
                }
            }

            if (options.ImportLists) {
                for (int i = 0; i < page.ListItems.Count; i++) {
                    PdfCore.PdfLogicalListItem listItem = page.ListItems[i];
                    items.Add(ImportItem.ForListItem(listItem, listItem.Line.BaselineY, sequence++));
                }
            }

            if (options.ImportTables) {
                IReadOnlyList<PdfCore.PdfLogicalTableExtraction> tables = PdfCore.PdfLogicalTableAnalysis.ExtractTables(page, options.MaxTableRows);
                for (int i = 0; i < tables.Count; i++) {
                    PdfCore.PdfLogicalTableExtraction table = tables[i];
                    items.Add(ImportItem.ForTable(table, table.Table.YTop, sequence++));
                }
            }

            AddLinkItems(page, options, navigation, consumedLinks, items, ref sequence);

            if (options.ImportImages || options.IncludeImagePlaceholders) {
                for (int i = 0; i < page.Images.Count; i++) {
                    PdfCore.PdfLogicalImage image = page.Images[i];
                    if (image.Placements.Count == 0) {
                        items.Add(ImportItem.ForImage(image, null, GetImageSortY(image), sequence++));
                        continue;
                    }

                    for (int placementIndex = 0; placementIndex < image.Placements.Count; placementIndex++) {
                        PdfCore.PdfImagePlacement placement = image.Placements[placementIndex];
                        items.Add(ImportItem.ForImage(image, placement, placement.Y + placement.Height, sequence++));
                    }
                }
            } else if (page.Images.Count > 0) {
                AddWarning(
                    options,
                    "PdfImageSkipped",
                    "Page " + page.PageNumber.ToString(CultureInfo.InvariantCulture) + "/Image",
                    "PDF image content was not imported because IncludeImagePlaceholders is false.",
                    PdfCore.PdfConversionWarningSeverity.Warning,
                    new Dictionary<string, string> {
                        ["ImageCount"] = page.Images.Count.ToString(CultureInfo.InvariantCulture)
                    });
            }

            if (options.IncludeFormFieldPlaceholders) {
                for (int i = 0; i < page.FormWidgets.Count; i++) {
                    PdfCore.PdfLogicalFormWidget widget = page.FormWidgets[i];
                    items.Add(ImportItem.ForFormWidget(widget, widget.Y2, sequence++));
                    AddWarning(
                        options,
                        "PdfFormWidgetPlaceholder",
                        "Page " + widget.PageNumber.ToString(CultureInfo.InvariantCulture) + "/FormWidget",
                        "PDF form widget content is represented as editable Word placeholder text; interactive form reconstruction is not part of the semantic import contract.",
                        PdfCore.PdfConversionWarningSeverity.Information,
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
                    new Dictionary<string, string> {
                        ["FormWidgetCount"] = page.FormWidgets.Count.ToString(CultureInfo.InvariantCulture)
                    });
            }

            return items;
        }

        private static void AddLinkItems(
            PdfCore.PdfLogicalPage page,
            PdfWordReadOptions options,
            ImportNavigationMap navigation,
            HashSet<PdfCore.PdfLogicalLinkAnnotation> consumedLinks,
            List<ImportItem> items,
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
                items.Add(ImportItem.ForLink(link, GetLinkSortY(link), sequence++, displayText));
            }
        }

        private static PdfCore.PdfLogicalLinkAnnotation? FindOverlappingImportableLink(
            PdfCore.PdfLogicalPage page,
            PdfCore.PdfLogicalTextBlock textBlock,
            PdfWordReadOptions options,
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
            PdfWordReadOptions options,
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
            PdfWordReadOptions options,
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

        private static void ReportSkippedUriLink(PdfCore.PdfLogicalLinkAnnotation link, PdfWordReadOptions options) {
            AddWarning(
                options,
                "PdfUriLinkSkippedUnsafe",
                "Page " + link.PageNumber.ToString(CultureInfo.InvariantCulture) + "/LinkAnnotation",
                "PDF URI link annotation was kept inert because it is not an absolute URI with an allowed Word hyperlink scheme.",
                PdfCore.PdfConversionWarningSeverity.Warning,
                new Dictionary<string, string> {
                    ["Uri"] = link.Uri ?? string.Empty
                });
        }

        private static void ReportSkippedInternalLink(PdfCore.PdfLogicalLinkAnnotation link, PdfWordReadOptions options) {
            AddWarning(
                options,
                "PdfInternalLinkNotReconstructed",
                "Page " + link.PageNumber.ToString(CultureInfo.InvariantCulture) + "/LinkAnnotation",
                "PDF internal link annotation could not be resolved to an imported Word bookmark.",
                PdfCore.PdfConversionWarningSeverity.Information,
                new Dictionary<string, string> {
                    ["DestinationName"] = link.DestinationName ?? string.Empty,
                    ["DestinationPageNumber"] = link.DestinationPageNumber?.ToString(CultureInfo.InvariantCulture) ?? string.Empty
                });
        }

        private static int CompareImportItems(ImportItem left, ImportItem right) {
            int yComparison = right.Y.CompareTo(left.Y);
            return yComparison != 0 ? yComparison : left.Sequence.CompareTo(right.Sequence);
        }

        private static void AddHeading(
            WordDocument document,
            PdfCore.PdfLogicalHeading heading,
            PdfCore.PdfLogicalLinkAnnotation? link,
            string? linkText,
            PdfWordReadOptions options,
            ImportNavigationMap navigation) {
            WordParagraph paragraph = link == null
                ? document.AddParagraph(heading.Text)
                : AddHyperlinkParagraph(document, link, string.IsNullOrWhiteSpace(linkText) ? heading.Text : linkText!, options, navigation);
            paragraph.SetStyle(MapHeadingStyle(heading.Level));
            paragraph.KeepWithNext = true;
            if (heading.FontSize > 0) {
                paragraph.SetFontSize((int)Math.Round(heading.FontSize, MidpointRounding.AwayFromZero));
            }
        }

        private static void AddParagraph(
            WordDocument document,
            PdfCore.PdfLogicalParagraph paragraph,
            PdfCore.PdfLogicalLinkAnnotation? link,
            string? linkText,
            PdfWordReadOptions options,
            ImportNavigationMap navigation) {
            if (link == null) {
                document.AddParagraph(paragraph.Text);
                return;
            }

            AddHyperlinkParagraph(document, link, string.IsNullOrWhiteSpace(linkText) ? paragraph.Text : linkText!, options, navigation);
        }

        private static void AddStandaloneLink(
            WordDocument document,
            PdfCore.PdfLogicalLinkAnnotation link,
            string? linkText,
            PdfWordReadOptions options,
            ImportNavigationMap navigation) {
            AddHyperlinkParagraph(document, link, GetLinkDisplayText(link, linkText), options, navigation);
        }

        private static WordParagraph AddHyperlinkParagraph(
            WordDocument document,
            PdfCore.PdfLogicalLinkAnnotation link,
            string text,
            PdfWordReadOptions options,
            ImportNavigationMap navigation) {
            if (!TryResolveWordLinkTarget(link, options, navigation, out WordLinkTarget target)) {
                return document.AddParagraph(text);
            }

            WordParagraph paragraph = document.AddParagraph();
            if (target.IsUri) {
                paragraph.AddHyperLink(text, target.Uri!, addStyle: true, tooltip: "Imported PDF link from page " + link.PageNumber.ToString(CultureInfo.InvariantCulture));
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

            paragraph.AddHyperLink(text, target.Anchor!, addStyle: true, tooltip: "Imported PDF internal link from page " + link.PageNumber.ToString(CultureInfo.InvariantCulture));
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

        private static void AddListItem(WordDocument document, PdfCore.PdfLogicalListItem item, ref WordList? bulletList, ref WordList? numberedList) {
            bool bullet = IsBulletMarker(item.Marker);
            WordList list = bullet
                ? bulletList ??= document.AddListBulleted()
                : numberedList ??= document.AddListNumbered();
            list.AddItem(item.Text, Math.Max(0, item.Level - 1));
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

        private static void AddTable(WordDocument document, PdfCore.PdfLogicalTableExtraction extraction, PdfWordReadOptions options) {
            PdfCore.PdfLogicalTableData data = extraction.Data;
            bool headerRowIncluded = HasHeaderRow(data);
            int columnCount = data.Columns.Count;
            int rowCount = data.Rows.Count + (headerRowIncluded ? 1 : 0);
            if (columnCount == 0 || rowCount == 0) {
                return;
            }

            WordTable table = document.AddTable(rowCount, columnCount, options.TableStyle);
            PopulateTable(table, data, headerRowIncluded, options);
        }

        private static bool HasHeaderRow(PdfCore.PdfLogicalTableData data) {
            return data.Columns.Count > 0
                && (data.Structure.HasHeaderRow || data.Structure.IsKeyValueTable)
                && data.Columns.Any(column => !string.IsNullOrWhiteSpace(column));
        }

        private static void PopulateTable(
            WordTable table,
            PdfCore.PdfLogicalTableData data,
            bool headerRowIncluded,
            PdfWordReadOptions options) {
            List<WordTableRow> rows = table.Rows;
            int rowOffset = headerRowIncluded ? 1 : 0;

            if (headerRowIncluded) {
                WriteRow(rows[0], data.Columns, data, alignNumericColumns: false);
                if (options.RepeatHeaderRows) {
                    rows[0].RepeatHeaderRowAtTheTopOfEachPage = true;
                }
            }

            for (int rowIndex = 0; rowIndex < data.Rows.Count; rowIndex++) {
                WriteRow(rows[rowIndex + rowOffset], data.Rows[rowIndex], data, options.AlignNumericColumns);
            }

            if (options.FitTablesToPageWidth) {
                table.WidthType = TableWidthUnitValues.Pct;
                table.Width = 5000;
                table.DistributeColumnsEvenly();
            }
        }

        private static void WriteRow(
            WordTableRow row,
            IReadOnlyList<string> values,
            PdfCore.PdfLogicalTableData data,
            bool alignNumericColumns) {
            List<WordTableCell> cells = row.Cells;
            for (int columnIndex = 0; columnIndex < cells.Count; columnIndex++) {
                string value = columnIndex < values.Count ? values[columnIndex] : string.Empty;
                WordParagraph paragraph = cells[columnIndex].AddParagraph(value ?? string.Empty, removeExistingParagraphs: true);
                if (alignNumericColumns && data.IsNumericColumn(columnIndex)) {
                    paragraph.ParagraphAlignment = JustificationValues.Right;
                }
            }
        }

        private static void AddImage(
            WordDocument document,
            PdfCore.PdfLogicalImage image,
            PdfCore.PdfImagePlacement? placement,
            PdfWordReadOptions options) {
            if (options.ImportImages && TryAddEmbeddedImage(document, image, placement, options)) {
                return;
            }

            if (options.IncludeImagePlaceholders) {
                AddImagePlaceholder(document, image, options.ImportImages ? "unsupported-image-payload" : "image-import-disabled");
            }
        }

        private static bool TryAddEmbeddedImage(
            WordDocument document,
            PdfCore.PdfLogicalImage image,
            PdfCore.PdfImagePlacement? placement,
            PdfWordReadOptions options) {
            PdfCore.PdfExtractedImage source = image.SourceImage;
            if (!source.IsImageFile || source.Bytes.Length == 0) {
                AddImageSkippedWarning(image, "PDF image stream is not exposed as a complete image file payload.");
                return false;
            }

            string extension = ResolveImageExtension(source);
            if (string.IsNullOrWhiteSpace(extension)) {
                AddImageSkippedWarning(image, "PDF image file extension could not be resolved for Word embedding.");
                return false;
            }

            string fileName = BuildImageFileName(image, extension);
            double? width = null;
            double? height = null;
            if (options.PreserveImagePlacementSize && placement != null && placement.Width > 0 && placement.Height > 0) {
                width = PdfPointsToWordPixels(placement.Width);
                height = PdfPointsToWordPixels(placement.Height);
            }

            try {
                using var stream = new MemoryStream(source.Bytes);
                document.AddParagraph().AddImage(stream, fileName, width, height, description: "Imported PDF image " + image.ResourceName + " from page " + image.PageNumber.ToString(CultureInfo.InvariantCulture));
                AddWarning(
                    options,
                    "PdfImageEmbedded",
                    "Page " + image.PageNumber.ToString(CultureInfo.InvariantCulture) + "/Image",
                    "PDF image content was embedded as a native Word image.",
                    PdfCore.PdfConversionWarningSeverity.Information,
                    new Dictionary<string, string> {
                        ["ResourceName"] = image.ResourceName,
                        ["Width"] = image.Width.ToString(CultureInfo.InvariantCulture),
                        ["Height"] = image.Height.ToString(CultureInfo.InvariantCulture),
                        ["MimeType"] = image.MimeType ?? string.Empty,
                        ["PlacementWidth"] = placement?.Width.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
                        ["PlacementHeight"] = placement?.Height.ToString(CultureInfo.InvariantCulture) ?? string.Empty
                    });
                if (source.HasUnresolvedTransparencyMask) {
                    AddWarning(
                        options,
                        "PdfImageTransparencyMaskNotResolved",
                        "Page " + image.PageNumber.ToString(CultureInfo.InvariantCulture) + "/Image",
                        "PDF image content was embedded as a native Word image, but its PDF transparency mask is not represented by the embedded image payload.",
                        PdfCore.PdfConversionWarningSeverity.Warning,
                        new Dictionary<string, string> {
                            ["ResourceName"] = image.ResourceName,
                            ["MaskKind"] = source.TransparencyMaskKind ?? string.Empty,
                            ["MimeType"] = image.MimeType ?? string.Empty
                        });
                }
                return true;
            } catch (Exception ex) when (ex is ArgumentException || ex is InvalidOperationException || ex is NotSupportedException) {
                AddImageSkippedWarning(image, "Word image embedding rejected the extracted PDF image payload: " + ex.Message);
                return false;
            }

            void AddImageSkippedWarning(PdfCore.PdfLogicalImage skippedImage, string message) {
                AddWarning(
                    options,
                    "PdfImageEmbeddingSkipped",
                    "Page " + skippedImage.PageNumber.ToString(CultureInfo.InvariantCulture) + "/Image",
                    message,
                    PdfCore.PdfConversionWarningSeverity.Warning,
                    new Dictionary<string, string> {
                        ["ResourceName"] = skippedImage.ResourceName,
                        ["MimeType"] = skippedImage.MimeType ?? string.Empty,
                        ["IsImageFile"] = skippedImage.SourceImage.IsImageFile ? "true" : "false"
                    });
            }
        }

        private static void AddImagePlaceholder(WordDocument document, PdfCore.PdfLogicalImage image, string reason) {
            string text = "[PDF image: page "
                + image.PageNumber.ToString(CultureInfo.InvariantCulture)
                + ", resource "
                + image.ResourceName
                + ", "
                + image.Width.ToString(CultureInfo.InvariantCulture)
                + "x"
                + image.Height.ToString(CultureInfo.InvariantCulture)
                + (image.MimeType == null ? string.Empty : ", " + image.MimeType)
                + ", "
                + reason
                + "]";
            document.AddParagraph(text).SetItalic();
        }

        private static string ResolveImageExtension(PdfCore.PdfExtractedImage image) {
            if (!string.IsNullOrWhiteSpace(image.FileExtension)) {
                return image.FileExtension!.TrimStart('.');
            }

            switch (image.MimeType?.ToLowerInvariant()) {
                case "image/jpeg":
                    return "jpg";
                case "image/png":
                    return "png";
                case "image/gif":
                    return "gif";
                case "image/bmp":
                    return "bmp";
                case "image/tiff":
                    return "tif";
                default:
                    return string.Empty;
            }
        }

        private static string BuildImageFileName(PdfCore.PdfLogicalImage image, string extension) {
            string resourceName = string.IsNullOrWhiteSpace(image.ResourceName) ? "image" : image.ResourceName;
            var safe = new char[resourceName.Length];
            for (int i = 0; i < resourceName.Length; i++) {
                char ch = resourceName[i];
                safe[i] = char.IsLetterOrDigit(ch) || ch == '-' || ch == '_' ? ch : '_';
            }

            return "pdf-page-"
                + image.PageNumber.ToString(CultureInfo.InvariantCulture)
                + "-"
                + new string(safe)
                + "."
                + extension;
        }

        private static double PdfPointsToWordPixels(double points) => points * 96D / 72D;

        private static void AddFormWidgetPlaceholder(WordDocument document, PdfCore.PdfLogicalFormWidget widget, PdfWordReadOptions options) {
            string name = string.IsNullOrWhiteSpace(widget.FieldName) ? "(unnamed)" : widget.FieldName!;
            string type = string.IsNullOrWhiteSpace(widget.FieldType) ? "field" : widget.FieldType!;
            string value = string.IsNullOrWhiteSpace(widget.Value) ? string.Empty : " = " + widget.Value;
            document.AddParagraph("[PDF form " + type + ": " + name + value + "]").SetItalic();
        }

        private static void ReportNonReconstructedLinks(PdfCore.PdfLogicalDocument source, PdfWordReadOptions options, ImportNavigationMap navigation) {
            int linkCount = source.Links.Count(link => !TryResolveWordLinkTarget(link, options, navigation, out _));
            if (linkCount == 0) {
                return;
            }

            AddWarning(
                options,
                "PdfLinkAnnotationNotReconstructed",
                "LinkAnnotation",
                "PDF link annotations that are remote, named viewer actions, unsafe, or unresolved are reported as diagnostics.",
                PdfCore.PdfConversionWarningSeverity.Information,
                new Dictionary<string, string> {
                    ["LinkCount"] = linkCount.ToString(CultureInfo.InvariantCulture)
                });
        }

        private static void CopyMetadata(PdfCore.PdfMetadata source, WordDocument target) {
            target.BuiltinDocumentProperties.Title = source.Title;
            target.BuiltinDocumentProperties.Creator = source.Author;
            target.BuiltinDocumentProperties.Subject = source.Subject;
            target.BuiltinDocumentProperties.Keywords = source.Keywords;
        }

        private static double GetImageSortY(PdfCore.PdfLogicalImage image) {
            if (image.PlacedY.HasValue || image.PlacedHeight.HasValue) {
                return image.PlacedY.GetValueOrDefault() + image.PlacedHeight.GetValueOrDefault();
            }

            return 0D;
        }

        private static void AddWarning(
            PdfWordReadOptions options,
            string code,
            string source,
            string message,
            PdfCore.PdfConversionWarningSeverity severity,
            IReadOnlyDictionary<string, string>? details = null) {
            options.ConversionReport.Add(new PdfCore.PdfConversionWarning(
                ConverterName,
                code,
                source,
                message,
                severity,
                details: details));
        }

        private sealed class ImportItem {
            private ImportItem(ImportItemKind kind, double y, int sequence) {
                Kind = kind;
                Y = y;
                Sequence = sequence;
            }

            public ImportItemKind Kind { get; }

            public double Y { get; }

            public int Sequence { get; }

            public PdfCore.PdfLogicalHeading? Heading { get; private set; }

            public PdfCore.PdfLogicalParagraph? Paragraph { get; private set; }

            public PdfCore.PdfLogicalListItem? ListItem { get; private set; }

            public PdfCore.PdfLogicalTableExtraction? TableExtraction { get; private set; }

            public PdfCore.PdfLogicalImage? Image { get; private set; }

            public PdfCore.PdfImagePlacement? ImagePlacement { get; private set; }

            public PdfCore.PdfLogicalFormWidget? FormWidget { get; private set; }

            public PdfCore.PdfLogicalLinkAnnotation? Link { get; private set; }

            public string? LinkText { get; private set; }

            public static ImportItem ForHeading(PdfCore.PdfLogicalHeading heading, double y, int sequence, PdfCore.PdfLogicalLinkAnnotation? link = null, string? linkText = null) =>
                new ImportItem(ImportItemKind.Heading, y, sequence) { Heading = heading, Link = link, LinkText = linkText };

            public static ImportItem ForParagraph(PdfCore.PdfLogicalParagraph paragraph, double y, int sequence, PdfCore.PdfLogicalLinkAnnotation? link = null, string? linkText = null) =>
                new ImportItem(ImportItemKind.Paragraph, y, sequence) { Paragraph = paragraph, Link = link, LinkText = linkText };

            public static ImportItem ForListItem(PdfCore.PdfLogicalListItem listItem, double y, int sequence) =>
                new ImportItem(ImportItemKind.ListItem, y, sequence) { ListItem = listItem };

            public static ImportItem ForTable(PdfCore.PdfLogicalTableExtraction table, double y, int sequence) =>
                new ImportItem(ImportItemKind.Table, y, sequence) { TableExtraction = table };

            public static ImportItem ForImage(PdfCore.PdfLogicalImage image, PdfCore.PdfImagePlacement? placement, double y, int sequence) =>
                new ImportItem(ImportItemKind.Image, y, sequence) { Image = image, ImagePlacement = placement };

            public static ImportItem ForFormWidget(PdfCore.PdfLogicalFormWidget widget, double y, int sequence) =>
                new ImportItem(ImportItemKind.FormWidget, y, sequence) { FormWidget = widget };

            public static ImportItem ForLink(PdfCore.PdfLogicalLinkAnnotation link, double y, int sequence, string? linkText) =>
                new ImportItem(ImportItemKind.Link, y, sequence) { Link = link, LinkText = linkText };
        }

        private enum ImportItemKind {
            Heading,
            Paragraph,
            ListItem,
            Table,
            Image,
            FormWidget,
            Link
        }
    }
}
