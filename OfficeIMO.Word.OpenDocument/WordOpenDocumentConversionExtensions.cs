using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.OpenDocument;
using OfficeIMO.Word;

namespace OfficeIMO.Word.OpenDocument;

/// <summary>Explicit conversions between OfficeIMO Word and native OpenDocument text models.</summary>
public static class WordOpenDocumentConversionExtensions {
    /// <summary>Converts a Word document to an in-memory ODT document.</summary>
    public static OdtDocument ToOpenDocument(this WordDocument source,
        WordOpenDocumentConversionOptions? options = null) => source.ToOpenDocumentResult(options).Value;

    /// <summary>Converts a Word document to an in-memory ODT document and reports every lossy mapping.</summary>
    public static OdfConversionResult<OdtDocument> ToOpenDocumentResult(this WordDocument source,
        WordOpenDocumentConversionOptions? options = null) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        WordOpenDocumentConversionOptions effective = options ?? new WordOpenDocumentConversionOptions();
        WordDocumentSnapshot snapshot = source.CreateInspectionSnapshot();
        OdtDocument target = OdtDocument.Create();
        var report = new OdfConversionReport("DOCX", "ODT");

        int paragraphs = 0, headings = 0, lists = 0, tables = 0, hyperlinks = 0, images = 0, unsupportedImages = 0, bookmarks = 0;
        int unsupportedFootnotes = 0, nestedListLevels = 0;
        IReadOnlyList<WordParagraphSnapshot> sourceParagraphs = EnumerateParagraphs(snapshot).ToList();
        IReadOnlyList<WordParagraphSnapshot> convertedHeaderFooterParagraphs = effective.IncludeHeadersAndFooters && snapshot.Sections.Count > 0
            ? EnumerateDefaultHeaderFooterParagraphs(snapshot.Sections[0]).ToList()
            : Array.Empty<WordParagraphSnapshot>();
        IEnumerable<WordParagraphSnapshot> convertedParagraphs = sourceParagraphs.Concat(convertedHeaderFooterParagraphs);
        int paragraphFormatting = convertedParagraphs.Count(HasUnsupportedParagraphFormatting);
        int runFormatting = convertedParagraphs.SelectMany(paragraph => paragraph.Runs).Count(HasUnsupportedRunFormatting);
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
                        CopyParagraph(paragraph, listParagraph, effective, ref hyperlinks, ref images, ref unsupportedImages, ref bookmarks, ref unsupportedFootnotes);
                        if (paragraph.ListLevel > 0) nestedListLevels++;
                        paragraphs++;
                        continue;
                    }

                    currentList = null;
                    currentOrdered = null;
                    int headingLevel = GetHeadingLevel(paragraph);
                    OdtParagraph converted = headingLevel > 0 ? target.AddHeading(string.Empty, headingLevel) : target.AddParagraph();
                    CopyParagraph(paragraph, converted, effective, ref hyperlinks, ref images, ref unsupportedImages, ref bookmarks, ref unsupportedFootnotes);
                    if (headingLevel > 0) headings++; else paragraphs++;
                } else if (block is WordTableSnapshot table) {
                    currentList = null;
                    currentOrdered = null;
                    ConvertTable(table, target, effective, ref hyperlinks, ref images, ref unsupportedImages,
                        ref bookmarks, ref unsupportedFootnotes);
                    tables++;
                }
            }
        }

        int headerFooterBlocks = snapshot.Sections.Sum(CountHeaderFooterBlocks);
        if (effective.IncludeHeadersAndFooters && snapshot.Sections.Count > 0) {
            WordSectionSnapshot first = snapshot.Sections[0];
            CopyHeaderFooter(first.DefaultHeader, target.PageLayout.Header, effective, ref hyperlinks, ref images,
                ref unsupportedImages, ref bookmarks, ref unsupportedFootnotes);
            CopyHeaderFooter(first.DefaultFooter, target.PageLayout.Footer, effective, ref hyperlinks, ref images,
                ref unsupportedImages, ref bookmarks, ref unsupportedFootnotes);
            int firstDefaultTables = (first.DefaultHeader?.Tables.Count ?? 0) + (first.DefaultFooter?.Tables.Count ?? 0);
            if (firstDefaultTables > 0) report.Add("header-footer-tables", OdfConversionMappingStatus.Skipped, firstDefaultTables,
                "Tables in the first section's default header and footer are not represented by the current ODT header/footer surface.");
            int laterDefaultBlocks = snapshot.Sections.Skip(1).Sum(section =>
                (section.DefaultHeader?.Elements.Count ?? 0) + (section.DefaultFooter?.Elements.Count ?? 0));
            if (laterDefaultBlocks > 0) report.Add("section-headers-footers", OdfConversionMappingStatus.Skipped, laterDefaultBlocks,
                "Default header and footer content from later Word sections is omitted because ODT conversion emits one page layout.");
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
        if (unsupportedImages > 0) report.Add("images", OdfConversionMappingStatus.Unsupported, unsupportedImages,
            "Word image parts using formats unsupported by OpenDocument were skipped.");
        AddCount(report, "bookmarks", bookmarks);
        if (snapshot.Sections.Count > 0) report.Add("page-layout", OdfConversionMappingStatus.Converted, 1);
        if (snapshot.Sections.Count > 1) report.Add("sections", OdfConversionMappingStatus.Approximated, snapshot.Sections.Count,
            "Section content is retained in order, but section-specific layout is collapsed to one ODT page layout.");
        if (paragraphFormatting > 0) report.Add("paragraph-formatting", OdfConversionMappingStatus.Approximated, paragraphFormatting,
            "Patterned shading, line spacing, borders, tab stops, bidirectional layout, and pagination controls outside the shared subset are flattened or omitted.");
        if (runFormatting > 0) report.Add("run-formatting", OdfConversionMappingStatus.Approximated, runFormatting,
            "Patterned or overlapping shading, capitalization, vertical alignment, double strike, non-single underline variants, and other Word-only run details are simplified or omitted.");
        if (tableFormatting > 0) report.Add("table-formatting", OdfConversionMappingStatus.Approximated, tableFormatting,
            "Table text and merges are retained; widths, borders, shading, styles, and repeated-header behavior are not fully mapped.");
        if (imageLayout > 0) report.Add("image-layout", OdfConversionMappingStatus.Approximated, imageLayout,
            "Image descriptions, titles, and advanced wrapping are not represented by the current ODT adapter.");
        if (unsupportedFootnotes > 0) report.Add("footnotes", OdfConversionMappingStatus.Unsupported, unsupportedFootnotes,
            "Footnote references are omitted from the current ODT adapter.");
        if (nestedListLevels > 0) report.Add("list-levels", OdfConversionMappingStatus.Approximated, nestedListLevels,
            "Nested Word list items are retained as top-level ODT list items because hierarchical list emission is not yet supported.");
        AddUnmappedWordFindings(source.InspectFeatures(), report, images, hyperlinks, bookmarks);
        return new OdfConversionResult<OdtDocument>(target, report).ApplyPolicy(effective.LossPolicy);
    }

    /// <summary>Converts an ODT document to an in-memory Word document.</summary>
    public static WordDocument ToWordDocument(this OdtDocument source,
        WordOpenDocumentConversionOptions? options = null) => source.ToWordDocumentResult(options).Value;

    /// <summary>Converts an ODT document to an in-memory Word document and reports every lossy mapping.</summary>
    public static OdfConversionResult<WordDocument> ToWordDocumentResult(this OdtDocument source,
        WordOpenDocumentConversionOptions? options = null) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        WordOpenDocumentConversionOptions effective = options ?? new WordOpenDocumentConversionOptions();
        WordDocument target = WordDocument.Create();
        var report = new OdfConversionReport("ODT", "DOCX");
        int paragraphs = 0, headings = 0, lists = 0, tables = 0, hyperlinks = 0, externalHyperlinks = 0, images = 0, bookmarks = 0;
        int approximatedRuns = 0, approximatedBookmarkRanges = 0, unsupportedMeasurements = 0;
        int sourceImages = source.ContentBlocks.Where(block => block.Paragraph != null).Sum(block => block.Paragraph!.Images.Count) +
            source.ContentBlocks.Where(block => block.Table != null).Sum(block => block.Table!.Rows
                .Sum(row => row.Cells.Sum(cell => cell.Paragraphs.Sum(paragraph => paragraph.Images.Count)))) +
            source.PageLayout.Header.Paragraphs.Sum(paragraph => paragraph.Images.Count) +
            source.PageLayout.Footer.Paragraphs.Sum(paragraph => paragraph.Images.Count);
        WordList? currentList = null;
        bool? currentOrdered = null;

        foreach (OdtContentBlock block in source.ContentBlocks) {
            if (block.Table != null) {
                currentList = null;
                currentOrdered = null;
                ConvertTable(block.Table, target, effective, ref hyperlinks, ref externalHyperlinks, ref images,
                    ref bookmarks, ref approximatedRuns, ref approximatedBookmarkRanges, ref unsupportedMeasurements);
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

            CopyParagraph(paragraph, converted, effective, ref hyperlinks, ref externalHyperlinks, ref images, ref bookmarks,
                ref approximatedRuns, ref approximatedBookmarkRanges, ref unsupportedMeasurements);
        }

        int unsupportedPageMeasurements = ApplyOdtPageLayout(source.PageLayout, target.Sections[0]);
        unsupportedMeasurements += unsupportedPageMeasurements;
        report.Add("page-layout", unsupportedPageMeasurements == 0
            ? OdfConversionMappingStatus.Converted
            : OdfConversionMappingStatus.Approximated, 1,
            unsupportedPageMeasurements == 0 ? null : "Relative page measurements were omitted while absolute layout values were retained.");

        if (effective.IncludeHeadersAndFooters &&
            (source.PageLayout.Header.Paragraphs.Count > 0 || source.PageLayout.Footer.Paragraphs.Count > 0)) {
            target.AddHeadersAndFooters();
            foreach (OdtParagraph paragraph in source.PageLayout.Header.Paragraphs) {
                WordParagraph converted = target.Header!.Default!.AddParagraph();
                CopyParagraph(paragraph, converted, effective, ref hyperlinks, ref externalHyperlinks, ref images, ref bookmarks,
                    ref approximatedRuns, ref approximatedBookmarkRanges, ref unsupportedMeasurements);
            }
            foreach (OdtParagraph paragraph in source.PageLayout.Footer.Paragraphs) {
                WordParagraph converted = target.Footer!.Default!.AddParagraph();
                CopyParagraph(paragraph, converted, effective, ref hyperlinks, ref externalHyperlinks, ref images, ref bookmarks,
                    ref approximatedRuns, ref approximatedBookmarkRanges, ref unsupportedMeasurements);
            }
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
        AddCount(report, "bookmarks", bookmarks);
        if (approximatedRuns > 0) report.Add("inline-formatting", OdfConversionMappingStatus.Approximated, approximatedRuns,
            "Inline elements outside the typed ODT text, span, hyperlink, image, and bookmark syntax were flattened to text.");
        if (approximatedBookmarkRanges > 0) report.Add("bookmark-ranges", OdfConversionMappingStatus.Approximated,
            approximatedBookmarkRanges, "ODT bookmark ranges were retained as collapsed Word bookmark targets at their start position.");
        if (sourceImages > images) report.Add("images", OdfConversionMappingStatus.Skipped, sourceImages - images,
            "Images were omitted because IncludeImages is disabled or their source bytes were unavailable.");
        if (unsupportedMeasurements > 0) report.Add("relative-measurements", OdfConversionMappingStatus.Unsupported,
            unsupportedMeasurements,
            "Relative or unsupported ODF lengths could not be projected to fixed Word point measurements and were omitted.");
        AddUnmappedOdfFindings(source.InspectFeatures(), report, externalHyperlinks, bookmarks, pageLayouts: 1);
        target = Normalize(target);
        return new OdfConversionResult<WordDocument>(target, report).ApplyPolicy(effective.LossPolicy);
    }

    private static void CopyParagraph(WordParagraphSnapshot source, OdtParagraph target,
        WordOpenDocumentConversionOptions options, ref int hyperlinks, ref int images, ref int unsupportedImages,
        ref int bookmarks, ref int unsupportedFootnotes) {
        bool wrote = false;
        foreach (WordRunSnapshot run in source.Runs) {
            if (!string.IsNullOrEmpty(run.Text)) {
                if (run.IsHyperlink && (!string.IsNullOrWhiteSpace(run.HyperlinkUri) || !string.IsNullOrWhiteSpace(run.HyperlinkAnchor))) {
                    OdtHyperlink link = target.AddHyperlink(run.Text, run.HyperlinkUri ?? "#" + run.HyperlinkAnchor);
                    ApplyWordRunFormatting(run, link);
                    hyperlinks++;
                } else {
                    OdtSpan span = target.AddSpan(run.Text);
                    ApplyWordRunFormatting(run, span);
                }
                wrote = true;
            }
            if (options.IncludeImages && run.InlineImage?.Bytes is { Length: > 0 } bytes) {
                WordInlineImageSnapshot image = run.InlineImage;
                try {
                    target.AddImage(bytes, image.FileName ?? "image.png",
                        OdfLength.Points(image.Width ?? 72D), OdfLength.Points(image.Height ?? 72D),
                        image.IsInline ? OdtImageAnchor.Inline : OdtImageAnchor.Paragraph);
                    images++;
                    wrote = true;
                } catch (NotSupportedException) {
                    unsupportedImages++;
                }
            }
            if (run.Footnote != null) unsupportedFootnotes++;
        }
        if (!wrote && source.Text.Length > 0) target.Text = source.Text;
        target.PageBreakBefore = source.PageBreakBefore;
        ApplyWordParagraphFormatting(source, target);
        if (!string.IsNullOrWhiteSpace(source.BookmarkName)) { target.AddBookmark(source.BookmarkName!); bookmarks++; }
    }

    private static void CopyParagraph(OdtParagraph source, WordParagraph target,
        WordOpenDocumentConversionOptions options, ref int hyperlinks, ref int externalHyperlinks, ref int images, ref int bookmarks,
        ref int approximatedRuns, ref int approximatedBookmarkRanges, ref int unsupportedMeasurements) {
        foreach (OdtInlineNode node in source.InlineNodes) {
            switch (node.Kind) {
                case OdtInlineNodeKind.Text:
                    unsupportedMeasurements += ApplyOdtParagraphTextFormatting(source, target.AddText(node.Text));
                    break;
                case OdtInlineNodeKind.Span:
                    unsupportedMeasurements += ApplyOdtSpanFormatting(node.Span!, source, target.AddText(node.Text));
                    break;
                case OdtInlineNodeKind.Hyperlink:
                    OdtHyperlink link = node.Hyperlink!;
                    WordParagraph? hyperlinkRun = null;
                    if (OdfUriReference.TryDecodeFragment(link.Href, out string fragment)) {
                        hyperlinkRun = target.AddHyperLink(link.Text, fragment, addStyle: true);
                    } else if (!link.Href.StartsWith("#", StringComparison.Ordinal)
                        && Uri.TryCreate(link.Href, UriKind.RelativeOrAbsolute, out Uri? uri)) {
                        hyperlinkRun = target.AddHyperLink(link.Text, uri, addStyle: true);
                    }
                    if (hyperlinkRun != null) {
                        unsupportedMeasurements += ApplyOdtHyperlinkFormatting(link, source, hyperlinkRun);
                        hyperlinks++;
                        if (IsExternalOdfHref(link.Href)) externalHyperlinks++;
                    } else {
                        unsupportedMeasurements += ApplyOdtParagraphTextFormatting(source, target.AddText(link.Text));
                        approximatedRuns++;
                    }
                    break;
                case OdtInlineNodeKind.Image:
                    if (!options.IncludeImages) break;
                    try {
                        OdtImage image = node.Image!;
                        using var stream = new MemoryStream(image.GetImageBytes(), writable: false);
                        target.AddImage(stream, Path.GetFileName(image.Path), image.Width.ToPoints(), image.Height.ToPoints());
                        images++;
                    } catch (Exception exception) when (exception is NotSupportedException || exception is InvalidDataException ||
                        exception is ArgumentException) {
                        // The loss report compares sourceImages with images and records the skipped media.
                    }
                    break;
                case OdtInlineNodeKind.Bookmark:
                    if (!string.IsNullOrWhiteSpace(node.Name)) {
                        target.AddBookmark(node.Name!);
                        bookmarks++;
                    }
                    break;
                case OdtInlineNodeKind.BookmarkStart:
                    if (!string.IsNullOrWhiteSpace(node.Name)) {
                        target.AddBookmark(node.Name!);
                        bookmarks++;
                        approximatedBookmarkRanges++;
                    }
                    break;
                case OdtInlineNodeKind.BookmarkEnd:
                    break;
                case OdtInlineNodeKind.Other:
                    if (node.Text.Length > 0) unsupportedMeasurements += ApplyOdtParagraphTextFormatting(source, target.AddText(node.Text));
                    approximatedRuns++;
                    break;
            }
        }

        target.PageBreakBefore = source.PageBreakBefore;
        unsupportedMeasurements += ApplyOdtParagraphFormatting(source, target);
    }

    private static void ApplyWordRunFormatting(WordRunSnapshot source, OdtSpan target) {
        target.Bold = source.Bold ? true : (bool?)null;
        target.Italic = source.Italic ? true : (bool?)null;
        target.Underline = source.Underline ? true : (bool?)null;
        target.StrikeThrough = source.Strike ? true : (bool?)null;
        if (source.FontSizePoints.HasValue) target.FontSize = OdfLength.Points(source.FontSizePoints.Value);
        if (!string.IsNullOrWhiteSpace(source.FontFamily)) target.FontFamily = source.FontFamily;
        if (OdfColor.TryParse(source.ColorHex, out OdfColor color)) target.Color = color;
        if (OdfColor.TryParse(source.RunShadingFillColorHex, out OdfColor shading)) target.BackgroundColor = shading;
        else if (TryMapWordHighlight(source.HighlightColor, out OdfColor highlight)) target.BackgroundColor = highlight;
    }

    private static void ApplyWordRunFormatting(WordRunSnapshot source, OdtHyperlink target) {
        target.Bold = source.Bold ? true : (bool?)null;
        target.Italic = source.Italic ? true : (bool?)null;
        target.Underline = source.Underline ? true : (bool?)null;
        target.StrikeThrough = source.Strike ? true : (bool?)null;
        if (source.FontSizePoints.HasValue) target.FontSize = OdfLength.Points(source.FontSizePoints.Value);
        if (!string.IsNullOrWhiteSpace(source.FontFamily)) target.FontFamily = source.FontFamily;
        if (OdfColor.TryParse(source.ColorHex, out OdfColor color)) target.Color = color;
        if (OdfColor.TryParse(source.RunShadingFillColorHex, out OdfColor shading)) target.BackgroundColor = shading;
        else if (TryMapWordHighlight(source.HighlightColor, out OdfColor highlight)) target.BackgroundColor = highlight;
    }

    private static void ApplyWordParagraphFormatting(WordParagraphSnapshot source, OdtParagraph target) {
        if (source.Alignment != null && TryMapWordAlignment(source.Alignment, out OdtParagraphAlignment alignment)) {
            target.Alignment = alignment;
        }
        if (source.IndentStartPoints.HasValue) target.IndentStart = OdfLength.Points(source.IndentStartPoints.Value);
        if (source.IndentEndPoints.HasValue) target.IndentEnd = OdfLength.Points(source.IndentEndPoints.Value);
        if (source.IndentFirstLinePoints.HasValue) target.FirstLineIndent = OdfLength.Points(source.IndentFirstLinePoints.Value);
        if (source.SpaceAbovePoints.HasValue) target.SpaceAbove = OdfLength.Points(source.SpaceAbovePoints.Value);
        if (source.SpaceBelowPoints.HasValue) target.SpaceBelow = OdfLength.Points(source.SpaceBelowPoints.Value);
        if (OdfColor.TryParse(source.ShadingFillColorHex, out OdfColor background)) target.BackgroundColor = background;
    }

    private static int ApplyOdtParagraphFormatting(OdtParagraph source, WordParagraph target) {
        int unsupported = 0;
        switch (source.Alignment) {
            case OdtParagraphAlignment.Start: target.ParagraphAlignment = WordParagraphAlignment.Start; break;
            case OdtParagraphAlignment.Center: target.ParagraphAlignment = WordParagraphAlignment.Center; break;
            case OdtParagraphAlignment.End: target.ParagraphAlignment = WordParagraphAlignment.End; break;
            case OdtParagraphAlignment.Justify: target.ParagraphAlignment = WordParagraphAlignment.Both; break;
        }
        if (source.IndentStart.HasValue) {
            if (source.IndentStart.Value.TryToPoints(out double points)) target.IndentationBeforePoints = points; else unsupported++;
        }
        if (source.IndentEnd.HasValue) {
            if (source.IndentEnd.Value.TryToPoints(out double points)) target.IndentationAfterPoints = points; else unsupported++;
        }
        if (source.FirstLineIndent.HasValue) {
            if (source.FirstLineIndent.Value.TryToPoints(out double points)) target.IndentationFirstLinePoints = points; else unsupported++;
        }
        if (source.SpaceAbove.HasValue) {
            if (source.SpaceAbove.Value.TryToPoints(out double points)) target.LineSpacingBeforePoints = points; else unsupported++;
        }
        if (source.SpaceBelow.HasValue) {
            if (source.SpaceBelow.Value.TryToPoints(out double points)) target.LineSpacingAfterPoints = points; else unsupported++;
        }
        if (source.BackgroundColor.HasValue) target.ShadingFillColorHex = source.BackgroundColor.Value.ToString();
        return unsupported;
    }

    private static int ApplyOdtHyperlinkFormatting(OdtHyperlink source, OdtParagraph paragraph, WordParagraph target) {
        target.Bold = source.Bold ?? paragraph.Bold ?? false;
        target.Italic = source.Italic ?? paragraph.Italic ?? false;
        target.Underline = (source.Underline ?? paragraph.Underline) == true ? WordUnderlineStyle.Single : (WordUnderlineStyle?)null;
        target.Strike = (source.StrikeThrough ?? paragraph.StrikeThrough) == true;
        OdfLength? fontSize = source.FontSize ?? paragraph.FontSize;
        int unsupported = ApplyOdtFontSize(fontSize, target);
        string? fontFamily = source.FontFamily ?? paragraph.FontFamily;
        if (!string.IsNullOrWhiteSpace(fontFamily)) target.FontFamily = fontFamily;
        OdfColor? color = source.Color ?? paragraph.Color;
        if (color.HasValue) target.ColorHex = color.Value.ToString();
        ApplyOdfTextBackground(source.BackgroundColor ?? paragraph.TextBackgroundColor, target);
        return unsupported;
    }

    private static int ApplyOdtParagraphTextFormatting(OdtParagraph source, WordParagraph target) {
        target.Bold = source.Bold == true;
        target.Italic = source.Italic == true;
        target.Underline = source.Underline == true ? WordUnderlineStyle.Single : (WordUnderlineStyle?)null;
        target.Strike = source.StrikeThrough == true;
        int unsupported = ApplyOdtFontSize(source.FontSize, target);
        if (!string.IsNullOrWhiteSpace(source.FontFamily)) target.FontFamily = source.FontFamily;
        if (source.Color.HasValue) target.ColorHex = source.Color.Value.ToString();
        ApplyOdfTextBackground(source.TextBackgroundColor, target);
        return unsupported;
    }

    private static int ApplyOdtSpanFormatting(OdtSpan source, OdtParagraph paragraph, WordParagraph target) {
        target.Bold = source.Bold ?? paragraph.Bold ?? false;
        target.Italic = source.Italic ?? paragraph.Italic ?? false;
        target.Underline = (source.Underline ?? paragraph.Underline) == true ? WordUnderlineStyle.Single : (WordUnderlineStyle?)null;
        target.Strike = (source.StrikeThrough ?? paragraph.StrikeThrough) == true;
        OdfLength? fontSize = source.FontSize ?? paragraph.FontSize;
        int unsupported = ApplyOdtFontSize(fontSize, target);
        string? fontFamily = source.FontFamily ?? paragraph.FontFamily;
        if (!string.IsNullOrWhiteSpace(fontFamily)) target.FontFamily = fontFamily;
        OdfColor? color = source.Color ?? paragraph.Color;
        if (color.HasValue) target.ColorHex = color.Value.ToString();
        ApplyOdfTextBackground(source.BackgroundColor ?? paragraph.TextBackgroundColor, target);
        return unsupported;
    }

    private static int ApplyOdtFontSize(OdfLength? fontSize, WordParagraph target) {
        if (!fontSize.HasValue) return 0;
        if (!fontSize.Value.TryToPoints(out double points)) return 1;
        double halfPoints = points * 2D;
        double roundedHalfPoints = Math.Round(halfPoints, MidpointRounding.AwayFromZero);
        if (Math.Abs(halfPoints - roundedHalfPoints) > 0.000000001D) return 1;
        target.FontSizePoints = roundedHalfPoints / 2D;
        return 0;
    }

    private static void ApplyOdfTextBackground(OdfColor? source, WordParagraph target) {
        if (!source.HasValue) return;
        if (TryMapOdfHighlight(source.Value, out WordHighlightColor highlight)) target.Highlight = highlight;
        else target.RunShadingFillColorHex = source.Value.ToString();
    }

    private static bool TryMapWordAlignment(string value, out OdtParagraphAlignment alignment) {
        switch (value.ToLowerInvariant()) {
            case "left":
            case "start": alignment = OdtParagraphAlignment.Start; return true;
            case "center": alignment = OdtParagraphAlignment.Center; return true;
            case "right":
            case "end": alignment = OdtParagraphAlignment.End; return true;
            case "both": alignment = OdtParagraphAlignment.Justify; return true;
            default: alignment = default; return false;
        }
    }

    private static bool TryMapWordHighlight(string? value, out OdfColor color) {
        string? hex;
        switch (value?.ToLowerInvariant()) {
            case "black": hex = "#000000"; break;
            case "blue": hex = "#0000FF"; break;
            case "cyan": hex = "#00FFFF"; break;
            case "green": hex = "#00FF00"; break;
            case "magenta": hex = "#FF00FF"; break;
            case "red": hex = "#FF0000"; break;
            case "yellow": hex = "#FFFF00"; break;
            case "white": hex = "#FFFFFF"; break;
            case "darkblue": hex = "#000080"; break;
            case "darkcyan": hex = "#008080"; break;
            case "darkgreen": hex = "#008000"; break;
            case "darkmagenta": hex = "#800080"; break;
            case "darkred": hex = "#800000"; break;
            case "darkyellow": hex = "#808000"; break;
            case "darkgray": hex = "#808080"; break;
            case "lightgray": hex = "#C0C0C0"; break;
            default: color = default; return false;
        }
        color = OdfColor.Parse(hex);
        return true;
    }

    private static bool TryMapOdfHighlight(OdfColor value, out WordHighlightColor highlight) {
        switch (value.ToString().ToUpperInvariant()) {
            case "#000000": highlight = WordHighlightColor.Black; return true;
            case "#0000FF": highlight = WordHighlightColor.Blue; return true;
            case "#00FFFF": highlight = WordHighlightColor.Cyan; return true;
            case "#00FF00": highlight = WordHighlightColor.Green; return true;
            case "#FF00FF": highlight = WordHighlightColor.Magenta; return true;
            case "#FF0000": highlight = WordHighlightColor.Red; return true;
            case "#FFFF00": highlight = WordHighlightColor.Yellow; return true;
            case "#FFFFFF": highlight = WordHighlightColor.White; return true;
            case "#000080": highlight = WordHighlightColor.DarkBlue; return true;
            case "#008080": highlight = WordHighlightColor.DarkCyan; return true;
            case "#008000": highlight = WordHighlightColor.DarkGreen; return true;
            case "#800080": highlight = WordHighlightColor.DarkMagenta; return true;
            case "#800000": highlight = WordHighlightColor.DarkRed; return true;
            case "#808000": highlight = WordHighlightColor.DarkYellow; return true;
            case "#808080": highlight = WordHighlightColor.DarkGray; return true;
            case "#C0C0C0": highlight = WordHighlightColor.LightGray; return true;
            default: highlight = default; return false;
        }
    }

    private static void ConvertTable(WordTableSnapshot source, OdtDocument targetDocument,
        WordOpenDocumentConversionOptions options, ref int hyperlinks, ref int images, ref int unsupportedImages,
        ref int bookmarks, ref int unsupportedFootnotes) {
        int rows = Math.Max(1, source.RowCount);
        int columns = Math.Max(1, source.ColumnCount);
        OdtTable target = targetDocument.AddTable(rows, columns, source.Title);
        var covered = new bool[rows, columns];
        foreach (WordTableRowSnapshot row in source.Rows) {
            foreach (WordTableCellSnapshot cell in row.Cells) {
                int column = cell.ColumnIndex;
                if (row.RowIndex < 0 || row.RowIndex >= rows || column < 0 || column >= columns || covered[row.RowIndex, column]) continue;
                OdtTableCell targetCell = target.Cell(row.RowIndex, column);
                for (int paragraphIndex = 0; paragraphIndex < cell.Paragraphs.Count; paragraphIndex++) {
                    OdtParagraph targetParagraph = paragraphIndex == 0
                        ? targetCell.Paragraphs[0]
                        : targetCell.AddParagraph();
                    CopyParagraph(cell.Paragraphs[paragraphIndex], targetParagraph, options, ref hyperlinks, ref images,
                        ref unsupportedImages, ref bookmarks, ref unsupportedFootnotes);
                }
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

    private static void ConvertTable(OdtTable source, WordDocument targetDocument,
        WordOpenDocumentConversionOptions options, ref int hyperlinks, ref int externalHyperlinks, ref int images,
        ref int bookmarks, ref int approximatedRuns, ref int approximatedBookmarkRanges, ref int unsupportedMeasurements) {
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
                for (int paragraphIndex = 0; paragraphIndex < cell.Paragraphs.Count; paragraphIndex++) {
                    WordParagraph targetParagraph = targetCell.AddParagraph(removeExistingParagraphs: paragraphIndex == 0);
                    CopyParagraph(cell.Paragraphs[paragraphIndex], targetParagraph, options, ref hyperlinks,
                        ref externalHyperlinks, ref images, ref bookmarks, ref approximatedRuns,
                        ref approximatedBookmarkRanges, ref unsupportedMeasurements);
                }
                if (cell.RowSpan > 1 || cell.ColumnSpan > 1) merges.Add((row, column, cell.RowSpan, cell.ColumnSpan));
            }
        }
        foreach (var merge in merges) {
            int rowSpan = Math.Min(merge.RowSpan, rows - merge.Row);
            int columnSpan = Math.Min(merge.ColumnSpan, columns - merge.Column);
            target.MergeCells(merge.Row, merge.Column, rowSpan, columnSpan);
        }
    }

    private static void CopyHeaderFooter(WordHeaderFooterSnapshot? source, OdtHeaderFooter target,
        WordOpenDocumentConversionOptions options, ref int hyperlinks, ref int images, ref int unsupportedImages,
        ref int bookmarks, ref int unsupportedFootnotes) {
        if (source == null) return;
        foreach (WordParagraphSnapshot paragraph in source.Paragraphs) {
            CopyParagraph(paragraph, target.AddParagraph(), options, ref hyperlinks, ref images, ref unsupportedImages,
                ref bookmarks, ref unsupportedFootnotes);
        }
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

    private static IEnumerable<WordParagraphSnapshot> EnumerateDefaultHeaderFooterParagraphs(WordSectionSnapshot section) =>
        new[] { section.DefaultHeader, section.DefaultFooter }
            .Where(item => item != null)
            .SelectMany(item => item!.Paragraphs);

    private static bool HasUnsupportedParagraphFormatting(WordParagraphSnapshot paragraph) =>
        (paragraph.Alignment != null && !TryMapWordAlignment(paragraph.Alignment, out _)) ||
        paragraph.LineSpacingValue.HasValue || paragraph.LineSpacingRule != null ||
        (paragraph.ShadingPattern.HasValue && paragraph.ShadingPattern.Value != WordShadingPattern.Nil &&
            paragraph.ShadingPattern.Value != WordShadingPattern.Clear) ||
        paragraph.LeftBorder != null || paragraph.RightBorder != null || paragraph.TopBorder != null || paragraph.BottomBorder != null ||
        paragraph.IsRightToLeft || paragraph.KeepWithNext || paragraph.KeepLinesTogether || paragraph.AvoidWidowAndOrphan || paragraph.TabStops.Count > 0;

    private static bool HasUnsupportedRunFormatting(WordRunSnapshot run) =>
        !string.IsNullOrWhiteSpace(run.VerticalTextAlignment) || !string.IsNullOrWhiteSpace(run.CapsStyle) ||
        run.DoubleStrike || (run.UnderlineStyle.HasValue && run.UnderlineStyle.Value != WordUnderlineStyle.None &&
            run.UnderlineStyle.Value != WordUnderlineStyle.Single) ||
        (run.RunShadingPattern.HasValue && run.RunShadingPattern.Value != WordShadingPattern.Nil &&
            run.RunShadingPattern.Value != WordShadingPattern.Clear) ||
        (!string.IsNullOrWhiteSpace(run.RunShadingFillColorHex) && !string.IsNullOrWhiteSpace(run.HighlightColor));

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

    private static int ApplyOdtPageLayout(OdtPageLayout source, WordSection target) {
        int unsupported = 0;
        if (source.Width.TryToPoints(out double width)) target.PageSettings.Width = checked((uint)Math.Round(width * 20D)); else unsupported++;
        if (source.Height.TryToPoints(out double height)) target.PageSettings.Height = checked((uint)Math.Round(height * 20D)); else unsupported++;
        if (source.MarginTop.TryToPoints(out double top)) target.Margins.Top = checked((int)Math.Round(top * 20D)); else unsupported++;
        if (source.MarginBottom.TryToPoints(out double bottom)) target.Margins.Bottom = checked((int)Math.Round(bottom * 20D)); else unsupported++;
        if (source.MarginLeft.TryToPoints(out double left)) target.Margins.Left = checked((uint)Math.Round(left * 20D)); else unsupported++;
        if (source.MarginRight.TryToPoints(out double right)) target.Margins.Right = checked((uint)Math.Round(right * 20D)); else unsupported++;
        return unsupported;
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

    private static void AddUnmappedOdfFindings(OdfFeatureReport features, OdfConversionReport report,
        int hyperlinks, int bookmarks, int pageLayouts) {
        foreach (OdfFeatureDiagnostic diagnostic in features.Diagnostics) {
            report.Add("source-inspection", OdfConversionMappingStatus.Unsupported, 1,
                diagnostic.Code + " in " + diagnostic.PartPath + ": " + diagnostic.Message);
        }
        int remainingHyperlinks = hyperlinks, remainingBookmarks = bookmarks, remainingPageLayouts = pageLayouts;
        foreach (OdfFeatureFinding finding in features.Findings) {
            int handled = 0;
            if (finding.Name == "external-links") {
                handled = Math.Min(remainingHyperlinks, finding.Count);
                remainingHyperlinks -= handled;
            } else if (finding.Name == "text-bookmarks") {
                handled = Math.Min(remainingBookmarks, finding.Count);
                remainingBookmarks -= handled;
            } else if (finding.Name == "master-pages") {
                handled = Math.Min(remainingPageLayouts, finding.Count);
                remainingPageLayouts -= handled;
            }
            int remaining = Math.Max(0, finding.Count - handled);
            if (remaining > 0) report.Add("source-" + finding.Name, OdfConversionMappingStatus.Unsupported, remaining,
                "The source feature is not represented by the DOCX conversion surface.");
        }
    }

    private static string Slug(string value) => new string(value.ToLowerInvariant().Select(character =>
        char.IsLetterOrDigit(character) ? character : '-').ToArray()).Trim('-');

    private static bool IsExternalOdfHref(string href) =>
        !string.IsNullOrWhiteSpace(href) && !href.StartsWith("#", StringComparison.Ordinal)
        && (href.StartsWith("//", StringComparison.Ordinal) || Uri.TryCreate(href, UriKind.Absolute, out _));

    private static WordDocument Normalize(WordDocument document) {
        byte[] bytes;
        try {
            using var stream = new MemoryStream();
            document.Save(stream);
            bytes = stream.ToArray();
        } finally {
            document.Dispose();
        }

        using var detachedSource = new MemoryStream(bytes, writable: false);
        return WordDocument.Load(detachedSource);
    }
}
