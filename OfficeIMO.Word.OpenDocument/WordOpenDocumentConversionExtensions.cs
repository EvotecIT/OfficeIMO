using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.OpenDocument;
using OfficeIMO.Word;

namespace OfficeIMO.Word.OpenDocument;

/// <summary>Explicit conversions between OfficeIMO Word and native OpenDocument text models.</summary>
public static class WordOpenDocumentConversionExtensions {
    /// <summary>Converts a Word document to an in-memory ODT document and reports every lossy mapping.</summary>
    public static OdfConversionResult<OdtDocument> ToOpenDocument(this WordDocument source,
        WordOpenDocumentConversionOptions? options = null) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        WordOpenDocumentConversionOptions effective = options ?? new WordOpenDocumentConversionOptions();
        WordDocumentSnapshot snapshot = source.CreateInspectionSnapshot();
        OdtDocument target = OdtDocument.Create();
        var report = new OdfConversionReport("DOCX", "ODT");

        int paragraphs = 0, headings = 0, lists = 0, tables = 0, hyperlinks = 0, images = 0, bookmarks = 0;
        int unsupportedFootnotes = 0;
        IReadOnlyList<WordParagraphSnapshot> sourceParagraphs = EnumerateParagraphs(snapshot).ToList();
        int paragraphFormatting = sourceParagraphs.Count(HasUnsupportedParagraphFormatting);
        int runFormatting = sourceParagraphs.SelectMany(paragraph => paragraph.Runs).Count(HasUnsupportedRunFormatting);
        int tableFormatting = snapshot.Sections.SelectMany(section => section.Elements).OfType<WordTableSnapshot>().Count(HasUnsupportedTableFormatting);
        int imageLayout = sourceParagraphs.SelectMany(paragraph => paragraph.Runs).Count(run => run.InlineImage != null &&
            (!string.IsNullOrWhiteSpace(run.InlineImage.Description) || !string.IsNullOrWhiteSpace(run.InlineImage.Title) ||
             (!run.InlineImage.IsInline && !string.IsNullOrWhiteSpace(run.InlineImage.WrapText))));
        if (snapshot.Sections.Count > 0) ApplyWordPageLayout(snapshot.Sections[0], target.PageLayout);
        foreach (WordSectionSnapshot section in snapshot.Sections) {
            OdtList? currentList = null;
            bool? currentOrdered = null;
            foreach (WordBlockSnapshot block in section.Elements.OrderBy(item => item.Order)) {
                if (block is WordParagraphSnapshot paragraph) {
                    if (paragraph.IsListItem) {
                        bool ordered = paragraph.IsOrderedList == true;
                        if (currentList == null || currentOrdered != ordered) {
                            currentList = target.AddList(ordered);
                            currentOrdered = ordered;
                            lists++;
                        }
                        OdtParagraph listParagraph = currentList.AddItem().Paragraphs[0];
                        CopyParagraph(paragraph, listParagraph, effective, ref hyperlinks, ref images, ref bookmarks, ref unsupportedFootnotes);
                        paragraphs++;
                        continue;
                    }

                    currentList = null;
                    currentOrdered = null;
                    int headingLevel = GetHeadingLevel(paragraph);
                    OdtParagraph converted = headingLevel > 0 ? target.AddHeading(string.Empty, headingLevel) : target.AddParagraph();
                    CopyParagraph(paragraph, converted, effective, ref hyperlinks, ref images, ref bookmarks, ref unsupportedFootnotes);
                    if (headingLevel > 0) headings++; else paragraphs++;
                } else if (block is WordTableSnapshot table) {
                    currentList = null;
                    currentOrdered = null;
                    ConvertTable(table, target);
                    tables++;
                }
            }
        }

        int headerFooterBlocks = snapshot.Sections.Sum(CountHeaderFooterBlocks);
        if (effective.IncludeHeadersAndFooters && snapshot.Sections.Count > 0) {
            WordSectionSnapshot first = snapshot.Sections[0];
            CopyHeaderFooter(first.DefaultHeader, target.PageLayout.Header);
            CopyHeaderFooter(first.DefaultFooter, target.PageLayout.Footer);
            int alternate = snapshot.Sections.Sum(section =>
                (section.FirstHeader == null ? 0 : 1) + (section.FirstFooter == null ? 0 : 1) +
                (section.EvenHeader == null ? 0 : 1) + (section.EvenFooter == null ? 0 : 1));
            if (alternate > 0) report.Add("alternate-headers-footers", OdfConversionMappingStatus.Unsupported, alternate,
                "ODT conversion currently maps only the first section's default header and footer.");
        } else if (headerFooterBlocks > 0) {
            report.Add("headers-footers", OdfConversionMappingStatus.Skipped, headerFooterBlocks,
                "Header and footer content was omitted because IncludeHeadersAndFooters is disabled.");
        }

        AddCount(report, "paragraphs", paragraphs);
        AddCount(report, "headings", headings);
        AddCount(report, "lists", lists);
        AddCount(report, "tables", tables);
        AddCount(report, "hyperlinks", hyperlinks);
        AddCount(report, "images", images);
        AddCount(report, "bookmarks", bookmarks);
        if (snapshot.Sections.Count > 0) report.Add("page-layout", OdfConversionMappingStatus.Converted, 1);
        if (snapshot.Sections.Count > 1) report.Add("sections", OdfConversionMappingStatus.Approximated, snapshot.Sections.Count,
            "Section content is retained in order, but section-specific layout is collapsed to one ODT page layout.");
        if (paragraphFormatting > 0) report.Add("paragraph-formatting", OdfConversionMappingStatus.Approximated, paragraphFormatting,
            "Alignment, indentation, spacing, borders, shading, tab stops, and pagination controls outside the shared subset are omitted.");
        if (runFormatting > 0) report.Add("run-formatting", OdfConversionMappingStatus.Approximated, runFormatting,
            "Underline, strike-through, highlight, capitalization, vertical alignment, and other Word-only run details are omitted.");
        if (tableFormatting > 0) report.Add("table-formatting", OdfConversionMappingStatus.Approximated, tableFormatting,
            "Table text and merges are retained; widths, borders, shading, styles, and repeated-header behavior are not fully mapped.");
        if (imageLayout > 0) report.Add("image-layout", OdfConversionMappingStatus.Approximated, imageLayout,
            "Image descriptions, titles, and advanced wrapping are not represented by the current ODT adapter.");
        if (unsupportedFootnotes > 0) report.Add("footnotes", OdfConversionMappingStatus.Unsupported, unsupportedFootnotes,
            "Footnote references are omitted from the current ODT adapter.");
        AddUnmappedWordFindings(source.InspectFeatures(), report, images, hyperlinks, bookmarks);
        return new OdfConversionResult<OdtDocument>(target, report);
    }

    /// <summary>Converts an ODT document to an in-memory Word document and reports every lossy mapping.</summary>
    public static OdfConversionResult<WordDocument> ToWordDocument(this OdtDocument source,
        WordOpenDocumentConversionOptions? options = null) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        WordOpenDocumentConversionOptions effective = options ?? new WordOpenDocumentConversionOptions();
        WordDocument target = WordDocument.Create();
        var report = new OdfConversionReport("ODT", "DOCX");
        int paragraphs = 0, headings = 0, lists = 0, tables = 0, hyperlinks = 0, images = 0, approximatedRuns = 0;
        int sourceImages = source.ContentBlocks.Where(block => block.Paragraph != null).Sum(block => block.Paragraph!.Images.Count);
        WordList? currentList = null;
        bool? currentOrdered = null;

        foreach (OdtContentBlock block in source.ContentBlocks) {
            if (block.Table != null) {
                currentList = null;
                currentOrdered = null;
                ConvertTable(block.Table, target);
                tables++;
                continue;
            }

            OdtParagraph paragraph = block.Paragraph!;
            WordParagraph converted;
            if (block.IsListItem) {
                bool ordered = block.IsOrderedList == true;
                if (currentList == null || currentOrdered != ordered) {
                    currentList = ordered ? target.AddListNumbered() : target.AddListBulleted();
                    currentOrdered = ordered;
                    lists++;
                }
                converted = currentList.AddItem(null, Math.Max(0, Math.Min(8, block.ListLevel)));
                paragraphs++;
            } else {
                currentList = null;
                currentOrdered = null;
                converted = target.AddParagraph();
                if (block.Kind == OdtContentBlockKind.Heading) {
                    converted.Style = HeadingStyle(paragraph.HeadingLevel ?? 1);
                    headings++;
                } else {
                    paragraphs++;
                }
            }

            CopyParagraph(paragraph, converted, effective, ref hyperlinks, ref images, ref approximatedRuns);
        }

        ApplyOdtPageLayout(source.PageLayout, target.Sections[0]);
        report.Add("page-layout", OdfConversionMappingStatus.Converted, 1);

        if (effective.IncludeHeadersAndFooters &&
            (source.PageLayout.Header.Paragraphs.Count > 0 || source.PageLayout.Footer.Paragraphs.Count > 0)) {
            target.AddHeadersAndFooters();
            foreach (OdtParagraph paragraph in source.PageLayout.Header.Paragraphs) target.Header!.Default!.AddParagraph(paragraph.Text);
            foreach (OdtParagraph paragraph in source.PageLayout.Footer.Paragraphs) target.Footer!.Default!.AddParagraph(paragraph.Text);
            report.Add("headers-footers", OdfConversionMappingStatus.Converted,
                source.PageLayout.Header.Paragraphs.Count + source.PageLayout.Footer.Paragraphs.Count);
        } else if (!effective.IncludeHeadersAndFooters &&
            (source.PageLayout.Header.Paragraphs.Count > 0 || source.PageLayout.Footer.Paragraphs.Count > 0)) {
            report.Add("headers-footers", OdfConversionMappingStatus.Skipped,
                source.PageLayout.Header.Paragraphs.Count + source.PageLayout.Footer.Paragraphs.Count,
                "Header and footer content was omitted because IncludeHeadersAndFooters is disabled.");
        }

        AddCount(report, "paragraphs", paragraphs);
        AddCount(report, "headings", headings);
        AddCount(report, "lists", lists);
        AddCount(report, "tables", tables);
        AddCount(report, "hyperlinks", hyperlinks);
        AddCount(report, "images", images);
        if (approximatedRuns > 0) report.Add("inline-formatting", OdfConversionMappingStatus.Approximated, approximatedRuns,
            "Mixed plain text, spans, and links are flattened when their exact inline order is not exposed by the typed ODT surface.");
        if (sourceImages > images) report.Add("images", OdfConversionMappingStatus.Skipped, sourceImages - images,
            "Images were omitted because IncludeImages is disabled or their source bytes were unavailable.");
        AddUnmappedOdfFindings(source.InspectFeatures(), report, hyperlinks);
        target = Normalize(target);
        return new OdfConversionResult<WordDocument>(target, report);
    }

    private static void CopyParagraph(WordParagraphSnapshot source, OdtParagraph target,
        WordOpenDocumentConversionOptions options, ref int hyperlinks, ref int images, ref int bookmarks, ref int unsupportedFootnotes) {
        bool wrote = false;
        foreach (WordRunSnapshot run in source.Runs) {
            if (!string.IsNullOrEmpty(run.Text)) {
                if (run.IsHyperlink && (!string.IsNullOrWhiteSpace(run.HyperlinkUri) || !string.IsNullOrWhiteSpace(run.HyperlinkAnchor))) {
                    target.AddHyperlink(run.Text, run.HyperlinkUri ?? "#" + run.HyperlinkAnchor);
                    hyperlinks++;
                } else {
                    OdtSpan span = target.AddSpan(run.Text);
                    span.Bold = run.Bold ? true : (bool?)null;
                    span.Italic = run.Italic ? true : (bool?)null;
                    if (run.FontSize.HasValue) span.FontSize = OdfLength.Points(run.FontSize.Value);
                    if (!string.IsNullOrWhiteSpace(run.ColorHex)) span.Color = OdfColor.Parse("#" + run.ColorHex!.TrimStart('#'));
                }
                wrote = true;
            }
            if (options.IncludeImages && run.InlineImage?.Bytes is { Length: > 0 } bytes) {
                WordInlineImageSnapshot image = run.InlineImage;
                target.AddImage(bytes, image.FileName ?? "image.png",
                    OdfLength.Points(image.Width ?? 72D), OdfLength.Points(image.Height ?? 72D),
                    image.IsInline ? OdtImageAnchor.Inline : OdtImageAnchor.Paragraph);
                images++;
                wrote = true;
            }
            if (run.Footnote != null) unsupportedFootnotes++;
        }
        if (!wrote && source.Text.Length > 0) target.Text = source.Text;
        target.PageBreakBefore = source.PageBreakBefore;
        if (!string.IsNullOrWhiteSpace(source.BookmarkName)) { target.AddBookmark(source.BookmarkName!); bookmarks++; }
    }

    private static void CopyParagraph(OdtParagraph source, WordParagraph target,
        WordOpenDocumentConversionOptions options, ref int hyperlinks, ref int images, ref int approximatedRuns) {
        IReadOnlyList<OdtSpan> spans = source.Spans;
        IReadOnlyList<OdtHyperlink> links = source.Hyperlinks;
        bool exactSpans = links.Count == 0 && spans.Count > 0 && string.Equals(string.Concat(spans.Select(span => span.Text)), source.Text, StringComparison.Ordinal);
        if (exactSpans) {
            foreach (OdtSpan span in spans) {
                WordParagraph run = target.AddText(span.Text);
                run.Bold = span.Bold == true;
                run.Italic = span.Italic == true;
                if (span.FontSize.HasValue) run.FontSize = checked((int)Math.Round(span.FontSize.Value.ToPoints()));
                if (span.Color.HasValue) run.ColorHex = span.Color.Value.ToString();
            }
        } else if (links.Count == 1 && string.Equals(links[0].Text, source.Text, StringComparison.Ordinal)) {
            string href = links[0].Href;
            if (href.StartsWith("#", StringComparison.Ordinal)) target.AddHyperLink(links[0].Text, href.Substring(1), addStyle: true);
            else if (Uri.TryCreate(href, UriKind.Absolute, out Uri? uri)) target.AddHyperLink(links[0].Text, uri, addStyle: true);
            else target.AddText(source.Text);
            hyperlinks++;
        } else {
            target.AddText(source.Text);
            if (spans.Count > 0 || links.Count > 0) approximatedRuns++;
        }

        target.PageBreakBefore = source.PageBreakBefore;
        target.Bold = source.Bold == true;
        target.Italic = source.Italic == true;
        if (source.FontSize.HasValue) target.FontSize = checked((int)Math.Round(source.FontSize.Value.ToPoints()));
        if (source.Color.HasValue) target.ColorHex = source.Color.Value.ToString();
        if (options.IncludeImages) {
            foreach (OdtImage image in source.Images) {
                using var stream = new MemoryStream(image.GetImageBytes(), writable: false);
                target.AddImage(stream, Path.GetFileName(image.Path), image.Width.ToPoints(), image.Height.ToPoints());
                images++;
            }
        }
    }

    private static void ConvertTable(WordTableSnapshot source, OdtDocument targetDocument) {
        int rows = Math.Max(1, source.RowCount);
        int columns = Math.Max(1, source.ColumnCount);
        OdtTable target = targetDocument.AddTable(rows, columns, source.Title);
        var covered = new bool[rows, columns];
        foreach (WordTableRowSnapshot row in source.Rows) {
            foreach (WordTableCellSnapshot cell in row.Cells) {
                int column = cell.ColumnIndex;
                if (row.RowIndex < 0 || row.RowIndex >= rows || column < 0 || column >= columns || covered[row.RowIndex, column]) continue;
                target.Cell(row.RowIndex, column).Text = string.Join("\n", cell.Paragraphs.Select(paragraph => paragraph.Text));
                int rowSpan = Math.Min(cell.RowSpan, rows - row.RowIndex);
                int columnSpan = Math.Min(cell.ColumnSpan, columns - column);
                if (rowSpan > 1 || columnSpan > 1) {
                    target.Merge(row.RowIndex, column, rowSpan, columnSpan);
                    for (int y = 0; y < rowSpan; y++) for (int x = 0; x < columnSpan; x++)
                        if (x != 0 || y != 0) covered[row.RowIndex + y, column + x] = true;
                }
            }
        }
    }

    private static void ConvertTable(OdtTable source, WordDocument targetDocument) {
        int rows = Math.Max(1, source.Rows.Count);
        int columns = Math.Max(1, source.Rows.Select(row => row.Cells.Count).DefaultIfEmpty(1).Max());
        WordTable target = targetDocument.AddTable(rows, columns);
        var merges = new List<(int Row, int Column, int RowSpan, int ColumnSpan)>();
        for (int row = 0; row < source.Rows.Count; row++) {
            IReadOnlyList<OdtTableCell> cells = source.Rows[row].Cells;
            for (int column = 0; column < cells.Count && column < columns; column++) {
                OdtTableCell cell = cells[column];
                if (cell.IsCovered) continue;
                WordTableCell targetCell = target.Rows[row].Cells[column];
                targetCell.Paragraphs[0].Text = cell.Text;
                if (cell.RowSpan > 1 || cell.ColumnSpan > 1) merges.Add((row, column, cell.RowSpan, cell.ColumnSpan));
            }
        }
        foreach (var merge in merges) {
            int rowSpan = Math.Min(merge.RowSpan, rows - merge.Row);
            int columnSpan = Math.Min(merge.ColumnSpan, columns - merge.Column);
            target.MergeCells(merge.Row, merge.Column, rowSpan, columnSpan);
        }
    }

    private static void CopyHeaderFooter(WordHeaderFooterSnapshot? source, OdtHeaderFooter target) {
        if (source == null) return;
        foreach (WordParagraphSnapshot paragraph in source.Paragraphs) target.AddParagraph(paragraph.Text);
    }

    private static int GetHeadingLevel(WordParagraphSnapshot paragraph) {
        string value = paragraph.StyleId ?? paragraph.StyleName ?? string.Empty;
        if (!value.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)) return 0;
        return int.TryParse(value.Substring(7), out int level) ? Math.Max(1, Math.Min(9, level)) : 0;
    }

    private static WordParagraphStyles HeadingStyle(int level) {
        switch (Math.Max(1, Math.Min(9, level))) {
            case 1: return WordParagraphStyles.Heading1;
            case 2: return WordParagraphStyles.Heading2;
            case 3: return WordParagraphStyles.Heading3;
            case 4: return WordParagraphStyles.Heading4;
            case 5: return WordParagraphStyles.Heading5;
            case 6: return WordParagraphStyles.Heading6;
            case 7: return WordParagraphStyles.Heading7;
            case 8: return WordParagraphStyles.Heading8;
            default: return WordParagraphStyles.Heading9;
        }
    }

    private static void AddCount(OdfConversionReport report, string feature, int count) {
        if (count > 0) report.Add(feature, OdfConversionMappingStatus.Converted, count);
    }

    private static IEnumerable<WordParagraphSnapshot> EnumerateParagraphs(WordDocumentSnapshot snapshot) {
        foreach (WordSectionSnapshot section in snapshot.Sections) {
            foreach (WordBlockSnapshot block in section.Elements) {
                if (block is WordParagraphSnapshot paragraph) yield return paragraph;
                else if (block is WordTableSnapshot table) {
                    foreach (WordParagraphSnapshot nested in table.Rows.SelectMany(row => row.Cells).SelectMany(cell => cell.Paragraphs)) yield return nested;
                }
            }
        }
    }

    private static bool HasUnsupportedParagraphFormatting(WordParagraphSnapshot paragraph) =>
        paragraph.Alignment != null || paragraph.IndentStartPoints.HasValue || paragraph.IndentEndPoints.HasValue ||
        paragraph.IndentFirstLinePoints.HasValue || paragraph.SpaceAbovePoints.HasValue || paragraph.SpaceBelowPoints.HasValue ||
        paragraph.LineSpacingValue.HasValue || paragraph.LineSpacingRule != null || paragraph.ShadingFillColorHex != null ||
        paragraph.LeftBorder != null || paragraph.RightBorder != null || paragraph.TopBorder != null || paragraph.BottomBorder != null ||
        paragraph.IsRightToLeft || paragraph.KeepWithNext || paragraph.KeepLinesTogether || paragraph.AvoidWidowAndOrphan || paragraph.TabStops.Count > 0;

    private static bool HasUnsupportedRunFormatting(WordRunSnapshot run) => run.Underline || run.Strike ||
        !string.IsNullOrWhiteSpace(run.FontFamily) || !string.IsNullOrWhiteSpace(run.HighlightColor) ||
        !string.IsNullOrWhiteSpace(run.VerticalTextAlignment) || !string.IsNullOrWhiteSpace(run.CapsStyle);

    private static bool HasUnsupportedTableFormatting(WordTableSnapshot table) => table.StyleName != null ||
        table.Description != null || table.RepeatHeaderRow || table.ColumnWidthPoints.Count > 0 ||
        table.Rows.SelectMany(row => row.Cells).Any(cell => cell.ShadingFillColorHex != null || cell.LeftBorder != null ||
            cell.RightBorder != null || cell.TopBorder != null || cell.BottomBorder != null);

    private static int CountHeaderFooterBlocks(WordSectionSnapshot section) => new[] {
        section.DefaultHeader, section.DefaultFooter, section.FirstHeader, section.FirstFooter, section.EvenHeader, section.EvenFooter
    }.Where(item => item != null).Sum(item => item!.Elements.Count);

    private static void ApplyWordPageLayout(WordSectionSnapshot source, OdtPageLayout target) {
        if (source.PageWidthPoints.HasValue) target.Width = OdfLength.Points(source.PageWidthPoints.Value);
        if (source.PageHeightPoints.HasValue) target.Height = OdfLength.Points(source.PageHeightPoints.Value);
        if (source.MarginTopPoints.HasValue) target.MarginTop = OdfLength.Points(source.MarginTopPoints.Value);
        if (source.MarginBottomPoints.HasValue) target.MarginBottom = OdfLength.Points(source.MarginBottomPoints.Value);
        if (source.MarginLeftPoints.HasValue) target.MarginLeft = OdfLength.Points(source.MarginLeftPoints.Value);
        if (source.MarginRightPoints.HasValue) target.MarginRight = OdfLength.Points(source.MarginRightPoints.Value);
    }

    private static void ApplyOdtPageLayout(OdtPageLayout source, WordSection target) {
        target.PageSettings.Width = checked((uint)Math.Round(source.Width.ToPoints() * 20D));
        target.PageSettings.Height = checked((uint)Math.Round(source.Height.ToPoints() * 20D));
        target.Margins.Top = checked((int)Math.Round(source.MarginTop.ToPoints() * 20D));
        target.Margins.Bottom = checked((int)Math.Round(source.MarginBottom.ToPoints() * 20D));
        target.Margins.Left = checked((uint)Math.Round(source.MarginLeft.ToPoints() * 20D));
        target.Margins.Right = checked((uint)Math.Round(source.MarginRight.ToPoints() * 20D));
    }

    private static void AddUnmappedWordFindings(WordFeatureReport features, OdfConversionReport report,
        int images, int hyperlinks, int bookmarks) {
        var structural = new HashSet<string>(StringComparer.Ordinal) { "Paragraphs", "Tables", "Sections", "Footnotes" };
        foreach (WordFeatureFinding finding in features.Features.Where(item => item.Count > 0 && !structural.Contains(item.Name))) {
            int handled = finding.Name == "Images" ? images : finding.Name == "External hyperlinks" ? hyperlinks :
                finding.Name == "Bookmarks" ? bookmarks : 0;
            int remaining = Math.Max(0, finding.Count - handled);
            if (remaining > 0) report.Add("source-" + Slug(finding.Name), OdfConversionMappingStatus.Unsupported, remaining, finding.Note);
        }
    }

    private static void AddUnmappedOdfFindings(OdfFeatureReport features, OdfConversionReport report, int hyperlinks) {
        int remainingHyperlinks = hyperlinks;
        foreach (OdfFeatureFinding finding in features.Findings) {
            int handled = finding.Name == "external-links" ? Math.Min(remainingHyperlinks, finding.Count) : 0;
            remainingHyperlinks -= handled;
            int remaining = Math.Max(0, finding.Count - handled);
            if (remaining > 0) report.Add("source-" + finding.Name, OdfConversionMappingStatus.Unsupported, remaining,
                "The source feature is not represented by the DOCX conversion surface.");
        }
    }

    private static string Slug(string value) => new string(value.ToLowerInvariant().Select(character =>
        char.IsLetterOrDigit(character) ? character : '-').ToArray()).Trim('-');

    private static WordDocument Normalize(WordDocument document) {
        var stream = new MemoryStream();
        document.Save(stream);
        document.Dispose();
        stream.Position = 0;
        return WordDocument.Load(stream, autoSave: false);
    }
}
