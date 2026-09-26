using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.OpenDocument;
using OfficeIMO.Word;
using System.Globalization;

namespace OfficeIMO.Word.OpenDocument;

/// <summary>Explicit conversions between OfficeIMO Word and native OpenDocument text models.</summary>
public static partial class WordOpenDocumentConversionExtensions {
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
        int nestedListLevels = 0;
        var notes = new NoteMappingStats(source);
        var imageValidationBudget = new OdfImageValidationBudget();
        IReadOnlyList<WordParagraphSnapshot> sourceParagraphs = EnumerateParagraphs(snapshot).ToList();
        IReadOnlyList<WordParagraphSnapshot> convertedHeaderFooterParagraphs = effective.IncludeHeadersAndFooters && snapshot.Sections.Count > 0
            ? EnumerateMappedHeaderFooterParagraphs(snapshot.Sections[0]).ToList()
            : Array.Empty<WordParagraphSnapshot>();
        IEnumerable<WordParagraphSnapshot> convertedParagraphs = sourceParagraphs.Concat(convertedHeaderFooterParagraphs);
        int paragraphFormatting = convertedParagraphs.Count(HasUnsupportedParagraphFormatting);
        int runFormatting = convertedParagraphs.SelectMany(paragraph => paragraph.Runs).Count(HasUnsupportedRunFormatting);
        int fieldResultFormatting = convertedParagraphs.SelectMany(paragraph => paragraph.InlineFields)
            .Count(field => field.HasFormattedResult && field.ResultText.Length > 0);
        int tableFormatting = snapshot.Sections.SelectMany(section => section.Elements).OfType<WordTableSnapshot>().Count(HasUnsupportedTableFormatting);
        int imageLayout = convertedParagraphs.SelectMany(paragraph => paragraph.Runs)
            .SelectMany(run => run.PositionedImages).Count(positioned =>
                !string.IsNullOrWhiteSpace(positioned.Image.Description) || !string.IsNullOrWhiteSpace(positioned.Image.Title) ||
                (!positioned.Image.IsInline && !string.IsNullOrWhiteSpace(positioned.Image.WrapText)));
        if (snapshot.Sections.Count > 0) ApplyWordPageLayout(snapshot.Sections[0], target.PageLayout);
        for (int sectionIndex = 0; sectionIndex < snapshot.Sections.Count; sectionIndex++) {
            WordSectionSnapshot section = snapshot.Sections[sectionIndex];
            notes.CurrentSectionIndex = sectionIndex;
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
                        CopyParagraph(paragraph, listParagraph, effective, imageValidationBudget, ref hyperlinks, ref images, ref unsupportedImages, ref bookmarks, notes);
                        if (paragraph.ListLevel > 0) nestedListLevels++;
                        paragraphs++;
                        continue;
                    }

                    currentList = null;
                    currentOrdered = null;
                    int headingLevel = GetHeadingLevel(paragraph);
                    OdtParagraph converted = headingLevel > 0 ? target.AddHeading(string.Empty, headingLevel) : target.AddParagraph();
                    CopyParagraph(paragraph, converted, effective, imageValidationBudget, ref hyperlinks, ref images, ref unsupportedImages, ref bookmarks, notes);
                    if (headingLevel > 0) headings++; else paragraphs++;
                } else if (block is WordTableSnapshot table) {
                    currentList = null;
                    currentOrdered = null;
                    ConvertTable(table, target, effective, imageValidationBudget, ref hyperlinks, ref images, ref unsupportedImages,
                        ref bookmarks, notes);
                    tables++;
                }
            }
        }

        int headerFooterBlocks = snapshot.Sections.Sum(CountHeaderFooterBlocks);
        if (effective.IncludeHeadersAndFooters && snapshot.Sections.Count > 0) {
            notes.CurrentSectionIndex = 0;
            WordSectionSnapshot first = snapshot.Sections[0];
            CopyHeaderFooter(first.DefaultHeader, target.PageLayout.Header, effective, imageValidationBudget, ref hyperlinks, ref images,
                ref unsupportedImages, ref bookmarks, notes);
            CopyHeaderFooter(first.DefaultFooter, target.PageLayout.Footer, effective, imageValidationBudget, ref hyperlinks, ref images,
                ref unsupportedImages, ref bookmarks, notes);
            if (first.DifferentFirstPage) {
                CopyHeaderFooter(first.FirstHeader, target.PageLayout.EnsureFirstHeader(), effective, imageValidationBudget,
                    ref hyperlinks, ref images, ref unsupportedImages, ref bookmarks, notes);
                CopyHeaderFooter(first.FirstFooter, target.PageLayout.EnsureFirstFooter(), effective, imageValidationBudget,
                    ref hyperlinks, ref images, ref unsupportedImages, ref bookmarks, notes);
            }
            if (first.DocumentOddEvenSettingEnabled) {
                CopyHeaderFooter(first.EvenHeader, target.PageLayout.EnsureLeftHeader(), effective, imageValidationBudget,
                    ref hyperlinks, ref images, ref unsupportedImages, ref bookmarks, notes);
                CopyHeaderFooter(first.EvenFooter, target.PageLayout.EnsureLeftFooter(), effective, imageValidationBudget,
                    ref hyperlinks, ref images, ref unsupportedImages, ref bookmarks, notes);
            }
            int firstTables = EnumerateMappedHeaderFooters(first).Sum(part => part?.Tables.Count ?? 0);
            if (firstTables > 0) report.Add("header-footer-tables", OdfConversionMappingStatus.Skipped, firstTables,
                "Header and footer tables are not represented by the current ODT header/footer surface.");
            int laterDefaultBlocks = snapshot.Sections.Skip(1).Sum(section =>
                (section.HasExplicitDefaultHeader ? Math.Max(1, section.DefaultHeader?.Elements.Count ?? 0) : 0) +
                (section.HasExplicitDefaultFooter ? Math.Max(1, section.DefaultFooter?.Elements.Count ?? 0) : 0));
            if (laterDefaultBlocks > 0) report.Add("section-headers-footers", OdfConversionMappingStatus.Skipped, laterDefaultBlocks,
                "Default header and footer content from later Word sections is omitted because ODT conversion emits one page layout.");
            int laterAlternate = snapshot.Sections.Skip(1).Sum(section =>
                (section.HasExplicitFirstHeader ? 1 : 0) + (section.HasExplicitFirstFooter ? 1 : 0) +
                (section.HasExplicitEvenHeader ? 1 : 0) + (section.HasExplicitEvenFooter ? 1 : 0));
            int inactiveAlternate = (first.DifferentFirstPage ? 0 : (first.HasExplicitFirstHeader ? 1 : 0) + (first.HasExplicitFirstFooter ? 1 : 0)) +
                (first.DocumentOddEvenSettingEnabled ? 0 : (first.HasExplicitEvenHeader ? 1 : 0) + (first.HasExplicitEvenFooter ? 1 : 0));
            if (laterAlternate + inactiveAlternate > 0) report.Add("alternate-headers-footers", OdfConversionMappingStatus.Unsupported,
                laterAlternate + inactiveAlternate, "Alternate header and footer parts outside the active first-section mapping are omitted.");
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
        int mappedFields = CountOdtFields(target);
        AddCount(report, "fields", mappedFields);
        int unmappedFields = Math.Max(0, source.InspectFields().Count - mappedFields);
        if (unmappedFields > 0) report.Add("fields", OdfConversionMappingStatus.Unsupported,
            unmappedFields, "Unsupported Word fields are flattened to visible cached text where available.");
        if (fieldResultFormatting > 0) report.Add("field-result-formatting", OdfConversionMappingStatus.Unsupported,
            fieldResultFormatting, "Direct formatting on cached Word field results is not retained when those results become ODT text.");
        if (snapshot.Sections.Count > 0) report.Add("page-layout", OdfConversionMappingStatus.Converted, 1);
        if (snapshot.Sections.Count > 1) report.Add("sections", OdfConversionMappingStatus.Approximated, snapshot.Sections.Count,
            "Section content is retained in order, but section-specific layout is collapsed to one ODT page layout.");
        if (paragraphFormatting > 0) report.Add("paragraph-formatting", OdfConversionMappingStatus.Approximated, paragraphFormatting,
            "Patterned shading, line spacing, borders, tab stops, bidirectional layout, and pagination controls outside the shared subset are flattened or omitted.");
        if (runFormatting > 0) report.Add("run-formatting", OdfConversionMappingStatus.Approximated, runFormatting,
            "Words-only or heavy underline variants and overlapping highlight/shading details are simplified because ODF has no exact equivalent.");
        if (tableFormatting > 0) report.Add("table-formatting", OdfConversionMappingStatus.Approximated, tableFormatting,
            "Table text and merges are retained; widths, borders, shading, styles, and repeated-header behavior are not fully mapped.");
        if (imageLayout > 0) report.Add("image-layout", OdfConversionMappingStatus.Approximated, imageLayout,
            "Image descriptions, titles, and advanced wrapping are not represented by the current ODT adapter.");
        int columnBreaks = convertedParagraphs.SelectMany(paragraph => paragraph.Runs)
            .Sum(run => run.NonTextBreaks?.Values.Count(kind => kind == WordBreakType.Column) ?? 0);
        if (columnBreaks > 0) report.Add("column-breaks", OdfConversionMappingStatus.Approximated, columnBreaks,
            "Word column breaks are retained as line breaks because the ODT paragraph projection does not preserve column flow.");
        CountWordNoteSettingsLoss(source, notes);
        AddNoteMappings(report, notes);
        if (nestedListLevels > 0) report.Add("list-levels", OdfConversionMappingStatus.Approximated, nestedListLevels,
            "Nested Word list items are retained as top-level ODT list items because hierarchical list emission is not yet supported.");
        AddUnmappedWordFindings(source.InspectFeatures(), report, images, hyperlinks, bookmarks, notes);
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
        int approximatedHeadingLevels = 0;
        int approximatedRuns = 0, approximatedBookmarkRanges = 0, unsupportedMeasurements = 0;
        int mappedFields = 0, unsupportedFields = 0;
        var handledUnsupportedFieldElements = new HashSet<System.Xml.Linq.XElement>();
        int approximatedFontFamilyLists = 0, unsupportedFontFamilies = 0;
        var notes = new NoteMappingStats {
            HasOdtDefaultNoteBodyFormatting = HasOdtDefaultNoteBodyFormatting(source),
            HasOdtDefaultNoteReferenceFormatting = HasOdtDefaultNoteReferenceFormatting(source)
        };
        CultureInfo textCaseCulture = OdfTextCultureResolver.Resolve(source.Metadata.Language);
        int approximatedTextDecorations = CountNonSolidTextDecorations(source);
        int unsupportedWritingModes = CountUnsupportedWritingModes(source);
        OdtPageLayout sourcePageLayout = source.PageLayout;
        OdtHeaderFooter[] sourceHeaderFooters = EnumerateOdtHeaderFooters(sourcePageLayout).ToArray();
        int sourceImages = source.ContentBlocks.Where(block => block.Paragraph != null).Sum(block => block.Paragraph!.Images.Count) +
            source.ContentBlocks.Where(block => block.Table != null).Sum(block => block.Table!.Rows
                .Sum(row => row.Cells.Sum(cell => cell.Paragraphs.Sum(paragraph => paragraph.Images.Count)))) +
            sourceHeaderFooters.Where(part => part.IsDisplayed)
                .Sum(part => part.Paragraphs.Sum(paragraph => paragraph.Images.Count));
        WordList? currentList = null;
        bool? currentOrdered = null;

        foreach (OdtContentBlock block in source.ContentBlocks) {
            if (block.Table != null) {
                currentList = null;
                currentOrdered = null;
                ConvertTable(block.Table, target, effective, textCaseCulture, ref hyperlinks, ref externalHyperlinks, ref images,
                    ref bookmarks, ref approximatedRuns, ref approximatedBookmarkRanges, ref unsupportedMeasurements,
                    ref approximatedFontFamilyLists, ref unsupportedFontFamilies, ref mappedFields, ref unsupportedFields,
                    handledUnsupportedFieldElements, notes);
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
                    if (paragraph.HeadingLevel > 9) approximatedHeadingLevels++;
                    converted.Style = HeadingStyle(paragraph.HeadingLevel ?? 1);
                    headings++;
                } else {
                    paragraphs++;
                }
            }

            CopyParagraph(paragraph, converted, effective, textCaseCulture, ref hyperlinks, ref externalHyperlinks, ref images, ref bookmarks,
                ref approximatedRuns, ref approximatedBookmarkRanges, ref unsupportedMeasurements,
                ref approximatedFontFamilyLists, ref unsupportedFontFamilies, ref mappedFields, ref unsupportedFields,
                handledUnsupportedFieldElements, notes);
        }

        int unsupportedPageMeasurements = ApplyOdtPageLayout(sourcePageLayout, target.Sections[0]);
        unsupportedMeasurements += unsupportedPageMeasurements;
        report.Add("page-layout", unsupportedPageMeasurements == 0
            ? OdfConversionMappingStatus.Converted
            : OdfConversionMappingStatus.Approximated, 1,
            unsupportedPageMeasurements == 0 ? null : "Relative page measurements were omitted while absolute layout values were retained.");

        int headerFooterParagraphs = sourceHeaderFooters.Sum(part => part.Paragraphs.Count);
        int displayedHeaderFooterParagraphs = sourceHeaderFooters.Where(part => part.IsDisplayed).Sum(part => part.Paragraphs.Count);
        int unsupportedHeaderFooterBlocks = sourceHeaderFooters.Where(part => part.IsDisplayed).Sum(part => part.NonParagraphBlockCount);
        if (effective.IncludeHeadersAndFooters) approximatedHeadingLevels += sourceHeaderFooters.Where(part => part.IsDisplayed)
            .Sum(part => part.Paragraphs.Count(paragraph => paragraph.HeadingLevel > 9));
        int hiddenHeaderFooterBlocks = sourceHeaderFooters.Where(part => !part.IsDisplayed)
            .Sum(part => part.Paragraphs.Count + part.NonParagraphBlockCount);
        bool hasAlternateHeaderFooter = sourcePageLayout.FirstHeader != null || sourcePageLayout.FirstFooter != null ||
            sourcePageLayout.LeftHeader != null || sourcePageLayout.LeftFooter != null;
        if (effective.IncludeHeadersAndFooters && (headerFooterParagraphs > 0 || unsupportedHeaderFooterBlocks > 0 || hiddenHeaderFooterBlocks > 0 || hasAlternateHeaderFooter)) {
            target.AddHeadersAndFooters();
            WordSection firstSection = target.Sections[0];
            if (sourcePageLayout.FirstHeader != null || sourcePageLayout.FirstFooter != null) {
                firstSection.DifferentFirstPage = true;
            }
            if (sourcePageLayout.LeftHeader != null || sourcePageLayout.LeftFooter != null) {
                firstSection.DifferentOddAndEvenPages = true;
            }
            foreach ((OdtHeaderFooter? story, WordHeaderFooterType kind, bool isHeader) in
                EnumerateOdtHeaderFooterVariants(sourcePageLayout)) {
                OdtHeaderFooter? effectiveStory = story ?? ResolveOdtHeaderFooterFallback(sourcePageLayout, kind, isHeader);
                if (effectiveStory == null) continue;
                WordHeaderFooter destination = isHeader
                    ? firstSection.GetOrCreateHeader(kind)
                    : firstSection.GetOrCreateFooter(kind);
                if (!effectiveStory.IsDisplayed) continue;
                if (story == null) {
                    CopyOdtHeaderFooterFallback(effectiveStory, destination, effective, textCaseCulture,
                        handledUnsupportedFieldElements);
                } else {
                    CopyOdtHeaderFooter(effectiveStory, destination, effective, textCaseCulture, ref hyperlinks, ref externalHyperlinks,
                        ref images, ref bookmarks, ref approximatedRuns, ref approximatedBookmarkRanges, ref unsupportedMeasurements,
                        ref approximatedFontFamilyLists, ref unsupportedFontFamilies, ref mappedFields, ref unsupportedFields,
                        handledUnsupportedFieldElements, notes);
                }
            }
            AddCount(report, "headers-footers", displayedHeaderFooterParagraphs);
            if (hiddenHeaderFooterBlocks > 0) report.Add("hidden-header-footer-content", OdfConversionMappingStatus.Skipped,
                hiddenHeaderFooterBlocks, "Header and footer content marked style:display='false' is omitted from the Word story.");
            if (unsupportedHeaderFooterBlocks > 0) report.Add("header-footer-blocks", OdfConversionMappingStatus.Unsupported,
                unsupportedHeaderFooterBlocks, "Header and footer blocks other than paragraphs and headings are omitted.");
        } else if (!effective.IncludeHeadersAndFooters && (headerFooterParagraphs > 0 || unsupportedHeaderFooterBlocks > 0 || hiddenHeaderFooterBlocks > 0 || hasAlternateHeaderFooter)) {
            report.Add("headers-footers", OdfConversionMappingStatus.Skipped,
                Math.Max(1, displayedHeaderFooterParagraphs + unsupportedHeaderFooterBlocks + hiddenHeaderFooterBlocks),
                "Header and footer content was omitted because IncludeHeadersAndFooters is disabled.");
        }

        AddCount(report, "paragraphs", paragraphs);
        AddCount(report, "headings", headings);
        if (approximatedHeadingLevels > 0) report.Add("heading-levels", OdfConversionMappingStatus.Approximated,
            approximatedHeadingLevels, "ODF outline level 10 was mapped to Word Heading9.");
        AddCount(report, "lists", lists);
        AddCount(report, "tables", tables);
        AddCount(report, "hyperlinks", hyperlinks);
        AddCount(report, "images", images);
        AddCount(report, "bookmarks", bookmarks);
        AddCount(report, "fields", mappedFields);
        if (unsupportedFields > 0) report.Add("fields", OdfConversionMappingStatus.Unsupported, unsupportedFields,
            "ODT fields with format, adjustment, fixed-value, or other unsupported properties retain only displayed text.");
        CountOdtNoteConfigurationLoss(source, notes);
        AddNoteMappings(report, notes);
        if (approximatedRuns > 0) report.Add("inline-formatting", OdfConversionMappingStatus.Approximated, approximatedRuns,
            "Inline elements outside the typed ODT text, span, hyperlink, image, and bookmark syntax were flattened to text.");
        if (approximatedTextDecorations > 0) report.Add("text-decorations", OdfConversionMappingStatus.Approximated,
            approximatedTextDecorations, "Patterned ODF line-through and non-wave patterned double underline variants are simplified to Word's nearest native decoration.");
        if (approximatedFontFamilyLists > 0) report.Add("font-family-fallbacks", OdfConversionMappingStatus.Approximated,
            approximatedFontFamilyLists, "Word run properties retain the first ODF font family but cannot retain the authored fallback list.");
        if (unsupportedFontFamilies > 0) report.Add("font-families", OdfConversionMappingStatus.Unsupported,
            unsupportedFontFamilies,
            "Malformed ODF font-family syntax was omitted instead of being emitted as an invalid Word typeface name.");
        if (unsupportedWritingModes > 0) report.Add("writing-mode", OdfConversionMappingStatus.Unsupported,
            unsupportedWritingModes, "Vertical and page-relative ODF writing modes are not represented by the current Word paragraph surface.");
        if (approximatedBookmarkRanges > 0) report.Add("bookmark-ranges", OdfConversionMappingStatus.Approximated,
            approximatedBookmarkRanges, "ODT bookmark ranges were retained as collapsed Word bookmark targets at their start position.");
        if (sourceImages > images) report.Add("images", OdfConversionMappingStatus.Skipped, sourceImages - images,
            "Images were omitted because IncludeImages is disabled or their source bytes were unavailable.");
        if (unsupportedMeasurements > 0) report.Add("relative-measurements", OdfConversionMappingStatus.Unsupported,
            unsupportedMeasurements,
            "Relative or unsupported ODF lengths could not be projected to fixed Word point measurements and were omitted.");
        AddUnmappedOdfFindings(source, source.InspectFeatures(), report, externalHyperlinks, bookmarks, pageLayouts: 1,
            handledUnsupportedFieldElements, notes);
        target = Normalize(target);
        return new OdfConversionResult<WordDocument>(target, report).ApplyPolicy(effective.LossPolicy);
    }

    private static void CopyParagraph(OdtParagraph source, WordParagraph target,
        WordOpenDocumentConversionOptions options, CultureInfo textCaseCulture,
        ref int hyperlinks, ref int externalHyperlinks, ref int images, ref int bookmarks,
        ref int approximatedRuns, ref int approximatedBookmarkRanges, ref int unsupportedMeasurements,
        ref int approximatedFontFamilyLists, ref int unsupportedFontFamilies,
        ref int mappedFields, ref int unsupportedFields,
        HashSet<System.Xml.Linq.XElement> handledUnsupportedFieldElements, NoteMappingStats notes,
        bool allowNotes = true) {
        OdtInlineLeaf[] leaves = FlattenOdtInlineNodes(source.InlineNodes).ToArray();
        OdfTextTransform?[] transforms = leaves.Select(leaf =>
            leaf.Span?.TextTransform ?? leaf.StyleLink?.TextTransform ?? source.TextTransform).ToArray();
        string[] displayTexts = leaves.Select(leaf => leaf.Node.Text).ToArray();
        for (int start = 0; start < transforms.Length;) {
            OdfTextTransform? transform = transforms[start];
            int end = start + 1;
            while (end < transforms.Length && transforms[end] == transform) end++;
            OfficeTextCase? textCase = transform switch {
                OdfTextTransform.Capitalize => OfficeTextCase.Capitalize,
                OdfTextTransform.Lowercase => OfficeTextCase.Lowercase,
                _ => null
            };
            if (textCase.HasValue) {
                IReadOnlyList<string> transformed = OfficeTextCaseTransformer.ApplySegments(
                    displayTexts.Skip(start).Take(end - start).ToArray(),
                    textCase.Value,
                    textCaseCulture);
                for (int index = start; index < end; index++) {
                    displayTexts[index] = transformed[index - start];
                }
            }
            start = end;
        }
        var convertedLinks = new HashSet<OdtHyperlink>();
        for (int nodeIndex = 0; nodeIndex < leaves.Length; nodeIndex++) {
            OdtInlineLeaf leaf = leaves[nodeIndex];
            OdtInlineNode node = leaf.Node;
            string displayText = displayTexts[nodeIndex];
            switch (node.Kind) {
                case OdtInlineNodeKind.Text:
                case OdtInlineNodeKind.Span:
                case OdtInlineNodeKind.Hyperlink:
                case OdtInlineNodeKind.Other:
                    WordParagraph? output = null;
                    OdtHyperlink? link = leaf.TargetLink;
                    if (link != null) {
                        if (OdfUriReference.TryDecodeFragment(link.Href, out string fragment)) {
                            output = target.AddHyperLink(displayText, fragment, addStyle: true);
                        } else if (!link.Href.StartsWith("#", StringComparison.Ordinal)
                            && Uri.TryCreate(link.Href, UriKind.RelativeOrAbsolute, out Uri? uri)) {
                            output = target.AddHyperLink(displayText, uri, addStyle: true);
                        }
                        if (output != null && convertedLinks.Add(link)) {
                            hyperlinks++;
                            if (IsExternalOdfHref(link.Href)) externalHyperlinks++;
                        }
                        if (output == null) approximatedRuns++;
                    }
                    output ??= target.AddText(displayText);
                    if (leaf.Span != null) {
                        unsupportedMeasurements += ApplyOdtSpanFormatting(leaf.Span, source, output,
                            ref approximatedFontFamilyLists, ref unsupportedFontFamilies);
                    } else if (leaf.StyleLink != null) {
                        unsupportedMeasurements += ApplyOdtHyperlinkFormatting(leaf.StyleLink, source, output,
                            ref approximatedFontFamilyLists, ref unsupportedFontFamilies);
                    } else {
                        unsupportedMeasurements += ApplyOdtParagraphTextFormatting(source, output,
                            ref approximatedFontFamilyLists, ref unsupportedFontFamilies);
                    }
                    if (node.Kind == OdtInlineNodeKind.Other) approximatedRuns++;
                    break;
                case OdtInlineNodeKind.Image:
                    if (leaf.TargetLink != null) approximatedRuns++;
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
                case OdtInlineNodeKind.Field:
                    OdtField field = node.Field!;
                    if (leaf.TargetLink == null && TryMapOdtField(field, out WordFieldType fieldType)) {
                        target.AddField(fieldType);
                        WordField wordField = target.Field!;
                        wordField.Text = displayText;
                        if (field.IsFixed) {
                            wordField.LockField = true;
                            wordField.UpdateField = false;
                        }
                        if (target._simpleField?.GetFirstChild<Run>() is Run resultRun) {
                            var result = new WordParagraph(target._document, target._paragraph, resultRun);
                            if (leaf.Span != null) {
                                unsupportedMeasurements += ApplyOdtSpanFormatting(leaf.Span, source, result,
                                    ref approximatedFontFamilyLists, ref unsupportedFontFamilies);
                            } else if (leaf.StyleLink != null) {
                                unsupportedMeasurements += ApplyOdtHyperlinkFormatting(leaf.StyleLink, source, result,
                                    ref approximatedFontFamilyLists, ref unsupportedFontFamilies);
                            } else {
                                unsupportedMeasurements += ApplyOdtParagraphTextFormatting(source, result,
                                    ref approximatedFontFamilyLists, ref unsupportedFontFamilies);
                            }
                        }
                        mappedFields++;
                        if (leaf.Span != null || leaf.StyleLink != null) approximatedRuns++;
                        if (transforms[nodeIndex] is OdfTextTransform.Lowercase or OdfTextTransform.Capitalize)
                            approximatedRuns++;
                    } else {
                        WordParagraph? result = null;
                        OdtHyperlink? fieldLink = leaf.TargetLink;
                        if (fieldLink != null) {
                            if (OdfUriReference.TryDecodeFragment(fieldLink.Href, out string fragment)) {
                                result = target.AddHyperLink(displayText, fragment, addStyle: true);
                            } else if (!fieldLink.Href.StartsWith("#", StringComparison.Ordinal)
                                && Uri.TryCreate(fieldLink.Href, UriKind.RelativeOrAbsolute, out Uri? uri)) {
                                result = target.AddHyperLink(displayText, uri, addStyle: true);
                            }
                            if (result != null && convertedLinks.Add(fieldLink)) {
                                hyperlinks++;
                                if (IsExternalOdfHref(fieldLink.Href)) externalHyperlinks++;
                            }
                            if (result == null) approximatedRuns++;
                        }
                        result ??= target.AddText(displayText);
                        if (leaf.Span != null) {
                            unsupportedMeasurements += ApplyOdtSpanFormatting(leaf.Span, source, result,
                                ref approximatedFontFamilyLists, ref unsupportedFontFamilies);
                        } else if (leaf.StyleLink != null) {
                            unsupportedMeasurements += ApplyOdtHyperlinkFormatting(leaf.StyleLink, source, result,
                                ref approximatedFontFamilyLists, ref unsupportedFontFamilies);
                        } else {
                            unsupportedMeasurements += ApplyOdtParagraphTextFormatting(source, result,
                                ref approximatedFontFamilyLists, ref unsupportedFontFamilies);
                        }
                        unsupportedFields++;
                        handledUnsupportedFieldElements.Add(field.Element);
                    }
                    break;
                case OdtInlineNodeKind.Note:
                    if (allowNotes) {
                        if (leaf.Span?.StyleName != null || leaf.StyleLink != null || leaf.TargetLink != null ||
                            HasOdtParagraphNoteReferenceFormatting(source) ||
                            notes.HasOdtDefaultNoteReferenceFormatting)
                            notes.ApproximatedReferenceFormatting++;
                        CopyOdtNote(node.Note!, target, notes);
                    }
                    else CountUnsupportedHeaderFooterNote(notes);
                    break;
            }
        }

        target.PageBreakBefore = source.PageBreakBefore;
        unsupportedMeasurements += ApplyOdtParagraphFormatting(source, target);
    }

    private readonly record struct OdtInlineLeaf(OdtInlineNode Node, OdtSpan? Span,
        OdtHyperlink? StyleLink, OdtHyperlink? TargetLink);

    private static IEnumerable<OdtInlineLeaf> FlattenOdtInlineNodes(
        IReadOnlyList<OdtInlineNode> nodes, OdtSpan? span = null,
        OdtHyperlink? styleLink = null, OdtHyperlink? targetLink = null) {
        foreach (OdtInlineNode node in nodes) {
            if (node.Kind == OdtInlineNodeKind.Span) {
                OdtSpan? styledSpan = node.Span?.StyleName != null ? node.Span : span;
                if (node.Children.Count == 0) yield return new OdtInlineLeaf(node, styledSpan, null, targetLink);
                else foreach (OdtInlineLeaf child in FlattenOdtInlineNodes(node.Children, styledSpan, null, targetLink))
                    yield return child;
            } else if (node.Kind == OdtInlineNodeKind.Hyperlink) {
                if (node.Children.Count == 0) yield return new OdtInlineLeaf(node, null, node.Hyperlink, node.Hyperlink);
                else foreach (OdtInlineLeaf child in FlattenOdtInlineNodes(node.Children, null, node.Hyperlink, node.Hyperlink))
                    yield return child;
            } else {
                yield return new OdtInlineLeaf(node, span, styleLink, targetLink);
            }
        }
    }

    private static void ApplyWordRunFormatting(WordRunSnapshot source, OdtSpan target) {
        target.Bold = source.Bold ? true : (bool?)null;
        target.Italic = source.Italic ? true : (bool?)null;
        target.Underline = source.Underline ? true : (bool?)null;
        target.StrikeThrough = source.Strike ? true : (bool?)null;
        ApplyWordDecoration(source, target);
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
        ApplyWordDecoration(source, target);
        if (source.FontSizePoints.HasValue) target.FontSize = OdfLength.Points(source.FontSizePoints.Value);
        if (!string.IsNullOrWhiteSpace(source.FontFamily)) target.FontFamily = source.FontFamily;
        if (OdfColor.TryParse(source.ColorHex, out OdfColor color)) target.Color = color;
        if (OdfColor.TryParse(source.RunShadingFillColorHex, out OdfColor shading)) target.BackgroundColor = shading;
        else if (TryMapWordHighlight(source.HighlightColor, out OdfColor highlight)) target.BackgroundColor = highlight;
    }

    private static bool TryMapWordAlignment(string value, out OdtParagraphAlignment alignment) {
        switch (value.ToLowerInvariant()) {
            case "left": alignment = OdtParagraphAlignment.Left; return true;
            case "start": alignment = OdtParagraphAlignment.Start; return true;
            case "center": alignment = OdtParagraphAlignment.Center; return true;
            case "right": alignment = OdtParagraphAlignment.Right; return true;
            case "end": alignment = OdtParagraphAlignment.End; return true;
            case "both": alignment = OdtParagraphAlignment.Justify; return true;
            default: alignment = default; return false;
        }
    }

    private static void ApplyWordDecoration(WordRunSnapshot source, OdtSpan target) {
        (target.UnderlineStyle, target.UnderlineType) = MapWordUnderline(source.UnderlineStyle);
        target.LineThroughStyle = source.Strike || source.DoubleStrike ? OdfTextDecorationStyle.Solid : OdfTextDecorationStyle.None;
        target.LineThroughType = source.DoubleStrike ? OdfTextDecorationType.Double : source.Strike ? OdfTextDecorationType.Single : OdfTextDecorationType.None;
        target.TextPosition = MapWordTextPosition(source.VerticalTextAlignment);
        target.TextTransform = string.Equals(source.CapsStyle, nameof(WordCapsStyle.Caps), StringComparison.OrdinalIgnoreCase)
            ? OdfTextTransform.Uppercase
            : OdfTextTransform.None;
        target.SmallCaps = string.Equals(source.CapsStyle, nameof(WordCapsStyle.SmallCaps), StringComparison.OrdinalIgnoreCase) ? true : (bool?)null;
    }

    private static void ApplyWordDecoration(WordRunSnapshot source, OdtHyperlink target) {
        (target.UnderlineStyle, target.UnderlineType) = MapWordUnderline(source.UnderlineStyle);
        target.LineThroughStyle = source.Strike || source.DoubleStrike ? OdfTextDecorationStyle.Solid : OdfTextDecorationStyle.None;
        target.LineThroughType = source.DoubleStrike ? OdfTextDecorationType.Double : source.Strike ? OdfTextDecorationType.Single : OdfTextDecorationType.None;
        target.TextPosition = MapWordTextPosition(source.VerticalTextAlignment);
        target.TextTransform = string.Equals(source.CapsStyle, nameof(WordCapsStyle.Caps), StringComparison.OrdinalIgnoreCase)
            ? OdfTextTransform.Uppercase
            : OdfTextTransform.None;
        target.SmallCaps = string.Equals(source.CapsStyle, nameof(WordCapsStyle.SmallCaps), StringComparison.OrdinalIgnoreCase) ? true : (bool?)null;
    }

    private static (OdfTextDecorationStyle? Style, OdfTextDecorationType? Type) MapWordUnderline(WordUnderlineStyle? style) => style switch {
        null or WordUnderlineStyle.None => (OdfTextDecorationStyle.None, OdfTextDecorationType.None),
        WordUnderlineStyle.Double => (OdfTextDecorationStyle.Solid, OdfTextDecorationType.Double),
        WordUnderlineStyle.WavyDouble => (OdfTextDecorationStyle.Wave, OdfTextDecorationType.Double),
        WordUnderlineStyle.Dotted or WordUnderlineStyle.DottedHeavy => (OdfTextDecorationStyle.Dotted, OdfTextDecorationType.Single),
        WordUnderlineStyle.Dash or WordUnderlineStyle.DashedHeavy => (OdfTextDecorationStyle.Dash, OdfTextDecorationType.Single),
        WordUnderlineStyle.DashLong or WordUnderlineStyle.DashLongHeavy => (OdfTextDecorationStyle.LongDash, OdfTextDecorationType.Single),
        WordUnderlineStyle.DotDash or WordUnderlineStyle.DashDotHeavy => (OdfTextDecorationStyle.DotDash, OdfTextDecorationType.Single),
        WordUnderlineStyle.DotDotDash or WordUnderlineStyle.DashDotDotHeavy => (OdfTextDecorationStyle.DotDotDash, OdfTextDecorationType.Single),
        WordUnderlineStyle.Wave or WordUnderlineStyle.WavyHeavy => (OdfTextDecorationStyle.Wave, OdfTextDecorationType.Single),
        _ => (OdfTextDecorationStyle.Solid, OdfTextDecorationType.Single)
    };

    private static OdfTextPosition? MapWordTextPosition(string? value) {
        if (string.Equals(value, nameof(WordVerticalTextPosition.Superscript), StringComparison.OrdinalIgnoreCase)) return OdfTextPosition.Superscript;
        if (string.Equals(value, nameof(WordVerticalTextPosition.Subscript), StringComparison.OrdinalIgnoreCase)) return OdfTextPosition.Subscript;
        if (string.Equals(value, nameof(WordVerticalTextPosition.Baseline), StringComparison.OrdinalIgnoreCase)) return OdfTextPosition.Normal;
        return null;
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
        WordOpenDocumentConversionOptions options, OdfImageValidationBudget imageValidationBudget,
        ref int hyperlinks, ref int images, ref int unsupportedImages,
        ref int bookmarks, NoteMappingStats notes) {
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
                    CopyParagraph(cell.Paragraphs[paragraphIndex], targetParagraph, options, imageValidationBudget, ref hyperlinks, ref images,
                        ref unsupportedImages, ref bookmarks, notes);
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
        WordOpenDocumentConversionOptions options, CultureInfo textCaseCulture,
        ref int hyperlinks, ref int externalHyperlinks, ref int images,
        ref int bookmarks, ref int approximatedRuns, ref int approximatedBookmarkRanges, ref int unsupportedMeasurements,
        ref int approximatedFontFamilyLists, ref int unsupportedFontFamilies,
        ref int mappedFields, ref int unsupportedFields,
        HashSet<System.Xml.Linq.XElement> handledUnsupportedFieldElements, NoteMappingStats notes) {
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
                    CopyParagraph(cell.Paragraphs[paragraphIndex], targetParagraph, options, textCaseCulture, ref hyperlinks,
                        ref externalHyperlinks, ref images, ref bookmarks, ref approximatedRuns,
                        ref approximatedBookmarkRanges, ref unsupportedMeasurements,
                        ref approximatedFontFamilyLists, ref unsupportedFontFamilies,
                        ref mappedFields, ref unsupportedFields, handledUnsupportedFieldElements, notes);
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
        WordOpenDocumentConversionOptions options, OdfImageValidationBudget imageValidationBudget,
        ref int hyperlinks, ref int images, ref int unsupportedImages,
        ref int bookmarks, NoteMappingStats notes) {
        if (source == null) return;
        foreach (WordParagraphSnapshot paragraph in source.Paragraphs) {
            int headingLevel = GetHeadingLevel(paragraph);
            OdtParagraph converted = headingLevel > 0 ? target.AddHeading(string.Empty, headingLevel) : target.AddParagraph();
            CopyParagraph(paragraph, converted, options, imageValidationBudget, ref hyperlinks, ref images, ref unsupportedImages,
                ref bookmarks, notes);
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

    private static IEnumerable<WordHeaderFooterSnapshot?> EnumerateMappedHeaderFooters(WordSectionSnapshot section) =>
        new[] { section.DefaultHeader, section.DefaultFooter }
            .Concat(section.DifferentFirstPage
                ? new[] { section.FirstHeader, section.FirstFooter }
                : Array.Empty<WordHeaderFooterSnapshot?>())
            .Concat(section.DocumentOddEvenSettingEnabled
                ? new[] { section.EvenHeader, section.EvenFooter }
                : Array.Empty<WordHeaderFooterSnapshot?>());

    private static IEnumerable<WordParagraphSnapshot> EnumerateMappedHeaderFooterParagraphs(WordSectionSnapshot section) =>
        EnumerateMappedHeaderFooters(section)
            .Where(item => item != null)
            .SelectMany(item => item!.Paragraphs);

    private static bool HasUnsupportedParagraphFormatting(WordParagraphSnapshot paragraph) =>
        (paragraph.Alignment != null && !TryMapWordAlignment(paragraph.Alignment, out _)) ||
        HasUnsupportedWordLineHeight(paragraph) ||
        (paragraph.ShadingPattern.HasValue && paragraph.ShadingPattern.Value != WordShadingPattern.Nil &&
            paragraph.ShadingPattern.Value != WordShadingPattern.Clear) ||
        paragraph.LeftBorder != null || paragraph.RightBorder != null || paragraph.TopBorder != null || paragraph.BottomBorder != null ||
        paragraph.KeepWithNext || paragraph.KeepLinesTogether || paragraph.AvoidWidowAndOrphan || paragraph.TabStops.Count > 0;

    private static bool HasUnsupportedWordLineHeight(WordParagraphSnapshot paragraph) {
        if (!paragraph.LineSpacingValue.HasValue && paragraph.LineSpacingRule == null) return false;
        return !TryMapWordLineHeight(paragraph, out _);
    }

    private static bool HasUnsupportedRunFormatting(WordRunSnapshot run) =>
        run.UnderlineStyle is WordUnderlineStyle.Words or WordUnderlineStyle.Thick or
            WordUnderlineStyle.DottedHeavy or WordUnderlineStyle.DashedHeavy or
            WordUnderlineStyle.DashLongHeavy or WordUnderlineStyle.DashDotHeavy or
            WordUnderlineStyle.DashDotDotHeavy or WordUnderlineStyle.WavyHeavy ||
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
        int images, int hyperlinks, int bookmarks, NoteMappingStats notes) {
        var structural = new HashSet<string>(StringComparer.Ordinal) { "Paragraphs", "Tables", "Sections", "Fields" };
        foreach (WordFeatureFinding finding in features.Features.Where(item => item.Count > 0 && !structural.Contains(item.Name))) {
            int handled = finding.Name == "Images" ? images : finding.Name == "External hyperlinks" ? hyperlinks :
                finding.Name == "Bookmarks" ? bookmarks : finding.Name == "Footnotes" ?
                    notes.SeenWordFootnotes + notes.UnreferencedFootnoteDefinitions :
                finding.Name == "Endnotes" ? notes.SeenWordEndnotes + notes.UnreferencedEndnoteDefinitions : 0;
            int remaining = Math.Max(0, finding.Count - handled);
            if (remaining > 0) report.Add("source-" + Slug(finding.Name), OdfConversionMappingStatus.Unsupported, remaining, finding.Note);
        }
    }

    private static void AddUnmappedOdfFindings(OdtDocument source, OdfFeatureReport features, OdfConversionReport report,
        int hyperlinks, int bookmarks, int pageLayouts,
        HashSet<System.Xml.Linq.XElement> handledUnsupportedFieldElements, NoteMappingStats notes) {
        foreach (OdfFeatureDiagnostic diagnostic in features.Diagnostics) {
            report.Add("source-inspection", OdfConversionMappingStatus.Unsupported, 1,
                diagnostic.Code + " in " + diagnostic.PartPath + ": " + diagnostic.Message);
        }
        int remainingHyperlinks = hyperlinks, remainingBookmarks = bookmarks,
            remainingPageLayouts = pageLayouts, remainingNotes = notes.SeenOdtNotes;
        foreach (OdfFeatureFinding finding in features.Findings) {
            if (finding.Name == "text-fields" && finding.Support == OdfFeatureSupport.Editable) continue;
            int handled = 0;
            if (finding.Name == "text-fields" && finding.Support == OdfFeatureSupport.Inspected &&
                finding.PartPath is string partPath && source.Package.ContainsEntry(partPath)) {
                System.Xml.Linq.XDocument part = source.Package.GetXml(partPath);
                handled = handledUnsupportedFieldElements.Count(element =>
                    ReferenceEquals(element.Document, part) &&
                    !(OdtField.IsBasicElement(element) &&
                        OdfFeatureInspector.IsEditableOdtField(source.Package.Kind, partPath, element)));
            } else if (finding.Name == "external-links") {
                handled = Math.Min(remainingHyperlinks, finding.Count);
                remainingHyperlinks -= handled;
            } else if (finding.Name == "text-bookmarks") {
                handled = Math.Min(remainingBookmarks, finding.Count);
                remainingBookmarks -= handled;
            } else if (finding.Name == "master-pages") {
                handled = Math.Min(remainingPageLayouts, finding.Count);
                remainingPageLayouts -= handled;
            } else if (finding.Name == "text-notes") {
                handled = Math.Min(remainingNotes, finding.Count);
                remainingNotes -= handled;
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
