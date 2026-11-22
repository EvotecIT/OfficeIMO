using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    private static bool RequiresNativeRunningStory(WordHeaderFooter? story) => story != null &&
        (story.ChildElements.Any(element => element is W.Table || element.Descendants<W.Table>().Any()) ||
        EnumerateNativeRunningFootnoteReferences(story).Any() ||
        NativeRunningStoryUsesSectionPageCount(story) ||
        story.ChildElements.Any(element => element.Descendants<DocumentFormat.OpenXml.Drawing.GraphicData>()
            .Any(data => data.Uri?.Value == "http://schemas.openxmlformats.org/drawingml/2006/diagram")));

    private static bool NativeRunningStoryUsesSectionPageCount(WordHeaderFooter? story) => story != null &&
        story.ChildElements.SelectMany(element => element is W.Paragraph paragraph
            ? new[] { paragraph }.Concat(element.Descendants<W.Paragraph>()) : element.Descendants<W.Paragraph>())
            .Any(paragraph => GetNativeHeaderFooterVisibleTextRuns(new WordParagraph(story.Document, paragraph))
                .Any(item => item.IsField && item.Text == "{sectionpages}"));

    private static IEnumerable<WordHeaderFooter> EnumerateNativeActiveRunningStories(WordSection section) =>
        new WordHeaderFooter?[] { section.ResolveDefaultHeader(), section.ResolveDefaultFooter(),
            section.PageSetup.DifferentFirstPage ? section.ResolveFirstHeader() : null,
            section.PageSetup.DifferentFirstPage ? section.ResolveFirstFooter() : null,
            section.DocumentOddEvenSettingEnabled ? section.ResolveEvenHeader() : null,
            section.DocumentOddEvenSettingEnabled ? section.ResolveEvenFooter() : null }
        .Where(story => story != null).Cast<WordHeaderFooter>();

    private static bool UsesNativeSectionPageCounts(WordSection section) =>
        EnumerateNativeActiveRunningStories(section).Any(NativeRunningStoryUsesSectionPageCount);

    private static bool UsesNativeRunningHeader(WordSection section, WordPdfExportOptions? options) => UsesNativeRunningStories(section, options, false,
        section.ResolveDefaultHeader(), section.PageSetup.DifferentFirstPage ? section.ResolveFirstHeader() : null,
        section.DocumentOddEvenSettingEnabled ? section.ResolveEvenHeader() : null);

    private static bool UsesNativeRunningFooter(WordSection section, WordPdfExportOptions? options) => UsesNativeRunningStories(section, options, true,
        section.ResolveDefaultFooter(), section.PageSetup.DifferentFirstPage ? section.ResolveFirstFooter() : null,
        section.DocumentOddEvenSettingEnabled ? section.ResolveEvenFooter() : null);

    private static bool UsesNativeRunningStories(WordSection section, WordPdfExportOptions? options, bool isFooter, params WordHeaderFooter?[] variants) {
        // Required table/field stories retain rich flow for every active sibling.
        // The paragraph renderer qualifies non-intersecting fixed page anchors;
        // other mixed/nested boxes retain bounded flow with approximation reports.
        if (variants.Any(RequiresNativeRunningStory)) return true;
        bool hasNumberedParagraph = false;
        bool hasParagraph = false;
        bool ordinaryStoriesSupported = !HasNativeRunningStoryOverrides(options?.PdfOptions, isFooter);
        PdfCore.PdfPageMargins margins = GetNativeMargins(section, options, (0D, 0D));
        double width = GetNativePageSize(section, options).Width.ToPoints() - margins.Left.ToPoints() - margins.Right.ToPoints();
        foreach (WordHeaderFooter? story in variants) {
            if (story == null) continue;
            hasParagraph |= !story.ChildElements.Any();
            double minimumHeight = 0D;
            double joinedLineHeight = double.PositiveInfinity;
            ordinaryStoriesSupported &= story.ChildElements.All(element => element is W.Paragraph ||
                element is W.SdtBlock block && IsNativeParagraphOnlyRunningBlock(block));
            // Without a required table/field story, retain the established simple
            // projection unless every text box can keep its own anchored frame.
            // It preserves its distinct presentation override contract.
            if (story.ChildElements.Any(element => element.Descendants<W.TextBoxContent>().Any()) &&
                !CanPlaceNativeRunningStoryTextBoxes(story, section, options)) return false;
            foreach (W.Paragraph source in story.ChildElements.SelectMany(element => element is W.Paragraph paragraph
                ? new[] { paragraph }.Concat(element.Descendants<W.Paragraph>()) : element.Descendants<W.Paragraph>())) {
                // Text-box paragraphs flow inside their own frames.
                if (source.Ancestors<W.TextBoxContent>().Any()) continue;
                var paragraph = new WordParagraph(story.Document, source);
                hasParagraph = true;
                // Inline pictures retain their measured line height; supported
                // qualified anchors retain paragraph/page placement without
                // reserving their paint frame as flowing header/footer content.
                // Anchored shape groups keep their own placement in the story's drawing layers.
                bool placedGroups = CanPlaceNativeRunningShapeGroups(GetNativePageSize(section, options), paragraph, options);
                ordinaryStoriesSupported &= source.Descendants<W.Drawing>()
                    .All(drawing => IsNativePlacedRunningTextBoxDrawing(drawing) || IsNativeRunningPictureDrawing(drawing, section, options, margins) ||
                        placedGroups && IsNativeShapeGroupDrawing(drawing)) &&
                    source.Descendants<W.Picture>().All(picture => IsNativePlacedRunningTextBoxFallback(picture) || IsNativeRunningVmlPicture(picture) ||
                        placedGroups && IsNativeShapeGroupFallback(picture)) &&
                    !source.Descendants<W.EmbeddedObject>().Any() &&
                    source.Descendants<AlternateContent>().All(alternate => IsNativePlacedRunningTextBoxAlternate(alternate) ||
                        placedGroups && IsNativeShapeGroupAlternate(alternate)) && !HasNativePageBreakBefore(paragraph) &&
                    !source.Descendants<W.Break>().Any(item => item.Type?.Value == W.BreakValues.Page ||
                        item.Type?.Value == W.BreakValues.Column);
                WordDocumentTraversal.ListInfo? info = WordDocumentTraversal.GetListInfo(paragraph);
                hasNumberedParagraph |= info?.MarkerVisible == true;
                var style = CreateNativeParagraphStyle(paragraph);
                // Admission needs a lower bound, not an independent layout
                // estimate. Adjacent spacing collapses, and hidden marks may
                // combine many source paragraphs into one line box. Ignore
                // spacing and use the smallest line box in each such chain.
                joinedLineHeight = Math.Min(joinedLineHeight,
                    ((style.FontSize?.ToPoints()) ?? NativeHeaderFooterFontSize) * (style.LineHeight ?? 1.2D));
                if (ShouldRenderNativeEmptyParagraphLineBox(paragraph)) {
                    minimumHeight += joinedLineHeight;
                    joinedLineHeight = double.PositiveInfinity;
                }
                if (info.HasValue) ApplyNativeInlineListIndent(paragraph, style);
                // Keep the established bounded simple-story export for indents
                // that cannot leave both paragraph-flow text frames positive.
                // Admission switches every paragraph in every active variant.
                double textWidth = width - style.LeftIndent.ToPoints() - style.RightIndent.ToPoints();
                double firstLineWidth = textWidth - style.FirstLineIndent.ToPoints();
                if (textWidth <= 0D || firstLineWidth <= 0D ||
                    double.IsNaN(textWidth) || double.IsInfinity(textWidth) ||
                    double.IsNaN(firstLineWidth) || double.IsInfinity(firstLineWidth)) return false;
            }
            // Retain the bounded simple-story export for obviously oversized
            // ordinary stories. Its expansion clamp preserves a usable body;
            // bounded running flow cannot create continuation pages for them.
            double distance = (isFooter ? checked((uint)(section.PageSetup.Margins.FooterDistance.GetValueOrDefault().ToTwips())) : checked((uint)(section.PageSetup.Margins.HeaderDistance.GetValueOrDefault().ToTwips()))) / 20D;
            ordinaryStoriesSupported &= minimumHeight + distance < GetNativePageSize(section, options).Height.ToPoints();
        }
        return hasNumberedParagraph || (hasParagraph && ordinaryStoriesSupported);
    }

    /// <summary>Every text box of the story is anchored with a frame the placed route reproduces.</summary>
    private static bool CanPlaceNativeRunningStoryTextBoxes(WordHeaderFooter story, WordSection section, WordPdfExportOptions? options) {
        PdfCore.PdfPageSize pageSize = GetNativePageSize(section, options);
        foreach (W.Paragraph source in story.ChildElements.SelectMany(element => element is W.Paragraph paragraph
            ? new[] { paragraph }.Concat(element.Descendants<W.Paragraph>()) : element.Descendants<W.Paragraph>())) {
            options?.CancellationToken.ThrowIfCancellationRequested();
            bool insideBox = source.Ancestors<W.TextBoxContent>().Any();
            if (!source.Descendants<W.TextBoxContent>().Any()) continue;
            // Nested boxes and boxes anchored in table cells keep the bounded flow.
            if (insideBox || source.Ancestors<W.TableCell>().Any()) return false;
            var paragraph = new WordParagraph(story.Document, source);
            if (!TryCollectNativePlacedRunningTextBoxes(pageSize, paragraph, GetNativeRuns(paragraph), section, options, out _, out _))
                return false;
        }
        return true;
    }

    /// <summary>
    /// Every DrawingML shape group of a header or footer paragraph is anchored and keeps its fixed placement.
    /// Groups that would flow or cannot be drawn keep the story on its established simple projection.
    /// </summary>
    private static bool CanPlaceNativeRunningShapeGroups(PdfCore.PdfPageSize pageSize, WordParagraph paragraph, WordPdfExportOptions? options) {
        if (!paragraph._paragraph.Descendants<W.Drawing>().Any(IsNativeShapeGroupDrawing)) return false;
        List<WordParagraph> runs = GetNativeRuns(paragraph);
        int groups = runs.SelectMany(run => run.EnumerateEffectiveRunContent().OfType<W.Drawing>()).Count(IsNativeShapeGroupDrawing);
        var style = new PdfCore.PdfParagraphStyle();
        int imageCount = 0;
        bool flowed = false;
        // A read-only projection: keep its warnings out of the real operation.
        ProjectNativeParagraphShapeGroups(pageSize, paragraph, runs, options?.CloneForConversion(), style,
            _ => flowed = true, ref imageCount, runningStory: true);
        return !flowed && groups > 0 && style.AnchoredCanvas?.Items.Count == groups;
    }

    /// <summary>The simple header and footer projection has no shape group route, so it reports each story's groups as omitted.</summary>
    private static void RecordNativeRunningShapeGroupDiagnostics(WordSection section, WordPdfExportOptions? options,
        bool mapsHeader, bool mapsFooter) {
        if (options == null) return;
        if (!mapsHeader) {
            Record(section.ResolveDefaultHeader(), "default header");
            if (section.PageSetup.DifferentFirstPage) Record(section.ResolveFirstHeader(), "first header");
            if (section.DocumentOddEvenSettingEnabled) Record(section.ResolveEvenHeader(), "even header");
        }
        if (!mapsFooter) {
            Record(section.ResolveDefaultFooter(), "default footer");
            if (section.PageSetup.DifferentFirstPage) Record(section.ResolveFirstFooter(), "first footer");
            if (section.DocumentOddEvenSettingEnabled) Record(section.ResolveEvenFooter(), "even footer");
        }

        void Record(WordHeaderFooter? story, string source) {
            OpenXmlElement? root = (OpenXmlElement?)story?._header ?? story?._footer;
            if (root == null) return;
            // A DrawingML group's VML fallback is the same group, so only fallback-free VML groups count separately.
            int groups = root.Descendants<W.Drawing>().Count(IsNativeShapeGroupDrawing) +
                root.Descendants<V.Group>().Count(group => !group.Ancestors<V.Group>().Any() && !group.Ancestors<AlternateContentFallback>().Any());
            if (groups == 0) return;
            AddNativeExportWarning(options, "native-header-footer-shape-group-unsupported", source,
                groups == 1 ? "A shape group in this header or footer is not drawn. Groups placed relative to the page or paragraph are drawn; margin-relative and VML-only groups are not yet."
                    : $"{groups} shape groups in this header or footer are not drawn. Groups placed relative to the page or paragraph are drawn; margin-relative and VML-only groups are not yet.");
        }
    }

    private static bool IsNativeShapeGroupDrawing(W.Drawing drawing) =>
        drawing.Descendants<DocumentFormat.OpenXml.Drawing.GraphicData>().Any(data => data.ChildElements
            .Any(child => child.NamespaceUri == NativeWordGroupNamespace && child.LocalName == "wgp"));

    // Word stores a VML copy of a DrawingML group as the compatibility fallback.
    private static bool IsNativeShapeGroupFallback(W.Picture picture) =>
        picture.Ancestors<AlternateContentFallback>().FirstOrDefault()?.Parent is AlternateContent alternate && IsNativeShapeGroupAlternate(alternate);

    private static bool IsNativeShapeGroupAlternate(AlternateContent alternate) =>
        alternate.Elements<AlternateContentChoice>().Any() &&
        alternate.Elements<AlternateContentChoice>().All(choice => choice.Descendants<W.Drawing>().Any() &&
            choice.Descendants<W.Drawing>().All(IsNativeShapeGroupDrawing));

    private static bool IsNativePlacedRunningTextBoxDrawing(W.Drawing drawing) =>
        drawing.Descendants<W.TextBoxContent>().Any() || drawing.Ancestors<W.TextBoxContent>().Any();

    // Word stores a VML copy of each text box as the compatibility fallback.
    private static bool IsNativePlacedRunningTextBoxFallback(W.Picture picture) =>
        picture.Ancestors<AlternateContentFallback>().Any() && picture.Descendants<W.TextBoxContent>().Any() ||
        picture.Ancestors<W.TextBoxContent>().Any();

    private static bool IsNativePlacedRunningTextBoxAlternate(AlternateContent alternate) =>
        alternate.Descendants<W.TextBoxContent>().Any() &&
        alternate.Elements<AlternateContentChoice>().All(choice => choice.Descendants<W.Drawing>().All(IsNativePlacedRunningTextBoxDrawing));

    // Word's page-number galleries wrap the footer paragraph in a block content
    // control. Its paragraphs flow like direct story paragraphs.
    private static bool IsNativeParagraphOnlyRunningBlock(W.SdtBlock block) =>
        block.SdtContentBlock is { } content && content.ChildElements.Any() &&
        content.ChildElements.All(child => child is W.Paragraph ||
            child is W.SdtBlock nested && IsNativeParagraphOnlyRunningBlock(nested));

    private static bool IsNativeRunningPictureDrawing(W.Drawing drawing, WordSection section,
        WordPdfExportOptions? options, PdfCore.PdfPageMargins margins) {
        var graphicData = drawing.Inline?.Graphic?.GraphicData ??
            drawing.Anchor?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Graphic>()?.GraphicData;
        if (graphicData?.Uri?.Value != "http://schemas.openxmlformats.org/drawingml/2006/picture" ||
            graphicData.Elements<DocumentFormat.OpenXml.Drawing.Pictures.Picture>().Count() != 1 ||
            drawing.Ancestors<W.SdtRun>().Any(IsNativePictureControl)) return false;
        if (drawing.Inline != null) return drawing.Anchor == null;
        if (!WordDrawingLayoutReader.TryRead(drawing, out WordDrawingLayoutInfo layout)) return false;
        if (layout.Wrap == WordDrawingWrapKind.None)
            return TryResolveNativeAnchoredImagePlacement(layout, section, GetNativePageSize(section, options), margins,
                layout.WidthPoints, layout.HeightPoints, out _);
        return TryResolveNativeAnchoredImageSource(drawing, section, GetNativePageSize(section, options), margins,
            layout.WidthPoints, layout.HeightPoints, out _, out _, out _, out _);
    }

    /// <summary>
    /// Keeps caller-configured simple-story positioning and typography on the
    /// path that implements those overrides. Registered body font resources do
    /// not prevent source paragraphs from using running flow.
    /// </summary>
    private static bool HasNativeRunningStoryOverrides(PdfCore.PdfCreateOptions? options, bool isFooter) => options != null &&
        (isFooter
            ? options.PageSetup.Footer.Offset.ToPoints() != 18D || options.PageSetup.Footer.HasExplicitStandardFont || options.PageSetup.Footer.FontFamily != null ||
                options.PageSetup.Footer.FontSize.ToPoints() != 9D || options.PageSetup.Footer.Color.HasValue || options.PageSetup.Footer.Alignment != global::OfficeIMO.OfficeAlignment.Center
            : options.PageSetup.Header.Offset.ToPoints() != 18D || options.PageSetup.Header.HasExplicitStandardFont || options.PageSetup.Header.FontFamily != null ||
                options.PageSetup.Header.FontSize.ToPoints() != 9D || options.PageSetup.Header.Color.HasValue || options.PageSetup.Header.Alignment != global::OfficeIMO.OfficeAlignment.Center);

    private static void ConfigureNativeRunningHeaderFooter(PdfCore.PdfPageBuilder page, WordSection section,
        WordPdfExportOptions? options, NativeFontMap fontMap,
        IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers, NativeDocumentDefaults defaults) {
        if (UsesNativeRunningHeader(section, options)) {
            double distance = checked((uint)(section.PageSetup.Margins.HeaderDistance.GetValueOrDefault().ToTwips())) / 20D;
            page.Header(header => {
                if (section.ResolveDefaultHeader() != null)
                    header.Content(CreateNativeRunningStory(section.ResolveDefaultHeader(), section, options, fontMap, listMarkers, defaults, false), global::OfficeIMO.OfficeLength.Points(distance));
                if (section.PageSetup.DifferentFirstPage && section.ResolveFirstHeader() != null)
                    header.FirstPageContent(CreateNativeRunningStory(section.ResolveFirstHeader(), section, options, fontMap, listMarkers, defaults, false), global::OfficeIMO.OfficeLength.Points(distance));
                if (section.DocumentOddEvenSettingEnabled && section.ResolveEvenHeader() != null)
                    header.EvenPagesContent(CreateNativeRunningStory(section.ResolveEvenHeader(), section, options, fontMap, listMarkers, defaults, false), global::OfficeIMO.OfficeLength.Points(distance));
            });
            ReserveNativeEmptyRunningContentDistances(page.Options, isHeader: true, section.ResolveDefaultHeader(),
                section.PageSetup.DifferentFirstPage ? section.ResolveFirstHeader() : null,
                section.DocumentOddEvenSettingEnabled ? section.ResolveEvenHeader() : null);
            if (!new[] { section.ResolveDefaultHeader(), section.PageSetup.DifferentFirstPage ? section.ResolveFirstHeader() : null,
                    section.DocumentOddEvenSettingEnabled ? section.ResolveEvenHeader() : null }.Any(RequiresNativeRunningStory))
                page.Options.AllowRunningContentTextFallback(isHeader: true);
        } else ConfigureNativeImplicitRunningStories(page, section, options, fontMap, listMarkers, defaults, isFooter: false);
        if (UsesNativeRunningFooter(section, options)) {
            double distance = checked((uint)(section.PageSetup.Margins.FooterDistance.GetValueOrDefault().ToTwips())) / 20D;
            page.Footer(footer => {
                if (section.ResolveDefaultFooter() != null)
                    footer.Content(CreateNativeRunningStory(section.ResolveDefaultFooter(), section, options, fontMap, listMarkers, defaults, true), global::OfficeIMO.OfficeLength.Points(distance));
                if (section.PageSetup.DifferentFirstPage && section.ResolveFirstFooter() != null)
                    footer.FirstPageContent(CreateNativeRunningStory(section.ResolveFirstFooter(), section, options, fontMap, listMarkers, defaults, true), global::OfficeIMO.OfficeLength.Points(distance));
                if (section.DocumentOddEvenSettingEnabled && section.ResolveEvenFooter() != null)
                    footer.EvenPagesContent(CreateNativeRunningStory(section.ResolveEvenFooter(), section, options, fontMap, listMarkers, defaults, true), global::OfficeIMO.OfficeLength.Points(distance));
            });
            ReserveNativeEmptyRunningContentDistances(page.Options, isHeader: false, section.ResolveDefaultFooter(),
                section.PageSetup.DifferentFirstPage ? section.ResolveFirstFooter() : null,
                section.DocumentOddEvenSettingEnabled ? section.ResolveEvenFooter() : null);
            RetainNativeTextFreeFooterProjection(page.Options, section.ResolveDefaultFooter(), listMarkers, fontMap,
                PdfCore.PdfRunningContentVariant.Default);
            if (section.PageSetup.DifferentFirstPage)
                RetainNativeTextFreeFooterProjection(page.Options, section.ResolveFirstFooter(), listMarkers, fontMap,
                    PdfCore.PdfRunningContentVariant.FirstPage);
            if (section.DocumentOddEvenSettingEnabled)
                RetainNativeTextFreeFooterProjection(page.Options, section.ResolveEvenFooter(), listMarkers, fontMap,
                    PdfCore.PdfRunningContentVariant.EvenPages);
            if (!new[] { section.ResolveDefaultFooter(), section.PageSetup.DifferentFirstPage ? section.ResolveFirstFooter() : null,
                    section.DocumentOddEvenSettingEnabled ? section.ResolveEvenFooter() : null }.Any(RequiresNativeRunningStory))
                page.Options.AllowRunningContentTextFallback(isHeader: false);
        } else ConfigureNativeImplicitRunningStories(page, section, options, fontMap, listMarkers, defaults, isFooter: true);
    }

    private static Func<PdfCore.PdfRunningContentContext, Action<PdfCore.PdfContentBuilder>> CreateNativeRunningStory(
        WordHeaderFooter? story, WordSection section, WordPdfExportOptions? options, NativeFontMap fontMap,
        IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers, NativeDocumentDefaults defaults, bool isFooter,
        bool includeGeneratedPageNumber = true) {
        defaults = defaults with { RunningStoryHasDisallowedTextBoxPeers = HasNativeRunningStoryPeerConstraint(story, defaults) };
        IReadOnlyList<WordElement> elements = story == null ? Array.Empty<WordElement>() : CollapseNativeParagraphElements(story.Elements);
        RecordNativeRunningStoryJoinDiagnostics(elements, listMarkers, defaults, fontMap, options);
        // Running callbacks are evaluated while the PDF is written, after the
        // operation report is captured. Discover approximation diagnostics now.
        if (story != null)
            foreach (W.Paragraph source in story.ChildElements.SelectMany(element => element is W.Paragraph paragraph
                ? new[] { paragraph }.Concat(element.Descendants<W.Paragraph>()) : element.Descendants<W.Paragraph>())) {
                if (source.Ancestors<W.TextBoxContent>().Any()) continue;
                var paragraph = new WordParagraph(story.Document, source);
                List<WordParagraph> runs = GetNativeRuns(paragraph);
                // Box frames are qualified from geometry and paint, not from the
                // paragraph's text, so field results cannot change the decision.
                if (!TryCollectNativePlacedRunningTextBoxes(GetNativePageSize(section, options), paragraph, runs, section, options, out _, out _,
                        defaults with { RunningContentSection = section }))
                    RecordNativeMixedTextBoxDiagnostic(paragraph, runs, options, "header/footer paragraph");
            }
        // A text-free source footer owns blank geometry; its generated number
        // keeps the established text projection and physical margin. Missing
        // variants have no native story and must not acquire an exclusion.
        bool addPageNumber = includeGeneratedPageNumber && isFooter && options?.IncludePageNumbers == true &&
            GetNativeHeaderFooterText(story, listMarkers, fontMap) is { HasPageTokens: false };
        WordHeaderFooterResolutionSnapshot? resolutionSnapshot = WordHeaderFooterResolutionSnapshot.Current(section._document);
        WordDocument sourceDocument = section._document;
        WordPdfTrackedChangeView trackedChangeView = CurrentNativeTrackedChangeView;
        return context => content => {
            // Running stories materialize while the PDF is written. Their source
            // document must still be open; report that contract plainly.
            if (sourceDocument.IsDisposed)
                throw new ObjectDisposedException(nameof(WordDocument),
                    "The source Word document was disposed before its PDF was written. Save or write the PDF before disposing the document.");
            using IDisposable? resolutionScope = resolutionSnapshot?.Use();
            using IDisposable trackedChangeScope = UseNativeTrackedChangeView(trackedChangeView);
            options?.CancellationToken.ThrowIfCancellationRequested();
            // Header and footer stories lay out on one page, so positioned frames can float there.
            var flow = new NativeSpacingCollapseFlow(new NativePdfColumnFlow(content,
                new PdfCore.PdfPageSize(context.PageWidth, context.PageHeight)) { SupportsPositionedTables = true }, preserveInitialSpacingBefore: true);
            NativeDocumentDefaults runningDefaults = defaults with { RunningContentContext = context, RunningContentSection = section };
            if (story != null && !story.ChildElements.Any())
                RenderNativeImplicitRunningParagraph(flow, story, runningDefaults, fontMap);
            NativeNoteNumbering notes = defaults.NoteNumbering?.CreateRunningReferenceSnapshot(section)
                ?? new NativeNoteNumbering(section._document, options);
            notes.BeginSection(section);
            var destinations = new Dictionary<W.Paragraph, string>();
            int rejectedJoinUntil = -1;
            for (int index = 0; index < elements.Count; index++) {
                options?.CancellationToken.ThrowIfCancellationRequested();
                if (TryRenderNativeParagraphFrame(flow, elements, ref index,
                    paragraph => listMarkers.TryGetValue(paragraph, out var frameMarker) ? frameMarker : null,
                    notes, options, context.ContentWidth.ToPoints(), runningDefaults, fontMap)) continue;
                if (TryRenderNativeJoinedParagraphs(flow, elements, ref index, ref rejectedJoinUntil,
                    paragraph => listMarkers.TryGetValue(paragraph, out var joinedMarker) ? joinedMarker : null,
                    notes, options, runningDefaults, fontMap)) continue;
                RenderNativeElement(flow, elements[index], section,
                    paragraph => listMarkers.TryGetValue(paragraph, out var marker) ? marker : null,
                    Array.Empty<int>(), notes, options, Array.Empty<NativeTableOfContentsEntry>(), destinations,
                    context.ContentWidth.ToPoints(), runningDefaults, fontMap,
                    index + 1 < elements.Count ? elements[index + 1] : null);
            }
            if (addPageNumber)
                content.AddParagraph(paragraph => paragraph.AddText(FormatNativeRunningPageTokens(
                    GetNativePageNumberFormat(options), context, null)), global::OfficeIMO.OfficeAlignment.Right);
        };
    }

    private static void RecordNativeRunningStoryJoinDiagnostics(IReadOnlyList<WordElement> elements,
        IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers, NativeDocumentDefaults defaults,
        NativeFontMap fontMap, WordPdfExportOptions? options) {
        if (options == null) return;
        for (int index = 0; index < elements.Count; index++) {
            if (elements[index] is WordTable table) {
                RecordNativeRunningTableJoinDiagnostics(table, listMarkers, defaults, fontMap, options);
                continue;
            }
            if (!TryGetNativeParagraphJoinChain(elements, index, out var paragraphs, out int lastIndex)) continue;
            if (!CanJoinNativeParagraphChain(paragraphs,
                paragraph => listMarkers.TryGetValue(paragraph, out var marker) ? marker : null, defaults, fontMap))
                AddNativeExportWarning(options, "native-hidden-paragraph-join-unsupported", "header/footer",
                    "Hidden paragraph marks beside lists, headings, decorated paragraphs, flow breaks or objects retain separate PDF paragraph boundaries.");
            index = lastIndex;
        }
    }

    private static void RecordNativeRunningTableJoinDiagnostics(WordTable table,
        IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers, NativeDocumentDefaults defaults,
        NativeFontMap fontMap, WordPdfExportOptions options, int depth = 0) {
        EnsureNativeTableDepth(depth);
        TableLayout layout = TableLayoutCache.GetLayout(table);
        NativeTableStyleDefaults tableStyle = GetNativeTableStyleDefaults(table, defaults,
            options.PdfOptions?.HasExplicitDefaultTableStyle == true);
        int columnCount = GetNativeTableColumnCount(layout);
        int headerRows = GetNativeTableVisualHeaderRowCount(table, layout.Rows.Count,
            GetNativeTableRepeatedHeaderRowCount(table, layout.Rows.Count));
        int footerStart = table.ConditionalFormattingLastRow == true && layout.Rows.Count > headerRows
            ? layout.Rows.Count - 1 : layout.Rows.Count;
        for (int rowIndex = 0; rowIndex < layout.Rows.Count; rowIndex++) {
            options.CancellationToken.ThrowIfCancellationRequested();
            int column = GetNativeTableRowStartColumn(layout, rowIndex);
            foreach (WordTableCell cell in layout.Rows[rowIndex]) {
                if (IsNativeHorizontalMergeContinuation(cell)) continue;
                int span = GetNativeCellColumnSpan(cell);
                if (!IsNativeVerticalMergeContinuation(cell)) {
                    NativeTableStyleDefaults cellStyle = GetNativeTableCellStyleDefaults(table, tableStyle,
                        rowIndex, column, span, columnCount, headerRows, footerStart);
                    List<WordElement> elements = EnumerateNativeTableCellElements(cell).ToList();
                    int rejectedCellJoinUntil = -1;
                    for (int index = 0; index < elements.Count; index++) {
                        if (elements[index] is WordTable nested) {
                            RecordNativeRunningTableJoinDiagnostics(nested, listMarkers, defaults, fontMap, options, depth + 1);
                        } else if (elements[index] is WordParagraph) {
                            // Use the renderer's eligibility, including effective
                            // conditional styles and caller-configured typography.
                            index = GetNativeJoinedCellParagraphEnd(elements, index, ref rejectedCellJoinUntil,
                                paragraph => listMarkers.TryGetValue(paragraph, out var marker) ? marker : null,
                                defaults, cellStyle, fontMap, options);
                        }
                    }
                }
                column += span;
            }
        }
    }

    private static List<WordParagraph> GetNativeRunningContentRuns(WordParagraph paragraph, PdfCore.PdfRunningContentContext? context) {
        if (context == null || paragraph._paragraph == null) return GetNativeRuns(paragraph);
        var serialized = new List<(W.Run Run, string Text, bool IsField, OpenXmlElement? SourceChild, PdfCore.PdfPageNumberStyle? FieldStyle)>();
        if (!TryBuildNativeHeaderFooterParagraphText(paragraph, out _, out _, serialized)) return GetNativeRuns(paragraph);
        var runs = new List<WordParagraph>(serialized.Count);
        foreach (var item in serialized) {
            var source = new WordParagraph(paragraph._document,
                item.Run.Ancestors<W.Paragraph>().FirstOrDefault() ?? paragraph._paragraph, item.Run) {
                _hyperlink = item.Run.Ancestors<W.Hyperlink>().FirstOrDefault()
            };
            if (item.IsField) {
                string text = FormatNativeRunningPageTokens(item.Text, context, item.FieldStyle);
                runs.Add(CreateNativeRunContentView(source, new OpenXmlElement[] {
                    new W.Text(text) { Space = SpaceProcessingModeValues.Preserve }
                }));
            } else if (item.SourceChild != null) {
                runs.Add(CreateNativeRunContentView(source, new[] { item.SourceChild }));
            }
        }
        return runs;
    }

    private static string FormatNativeRunningPageTokens(string text, PdfCore.PdfRunningContentContext context, PdfCore.PdfPageNumberStyle? style) =>
        text.Replace("{page}", context.FormatPageNumber(context.PageNumber, style))
            .Replace("{pages}", context.FormatPageNumber(context.TotalPages, style))
            .Replace("{sectionpages}", context.FormatPageNumber(context.SectionPages, style))
            .Replace("{documentpages}", context.FormatPageNumber(context.DocumentPages, style));
}
