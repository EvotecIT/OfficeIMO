using System;
using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.Word {
    internal static partial class WordDocumentImageRenderer {
        private const double TwipsPerPoint = 20D;
        private const double DefaultPageWidthPoints = 595.3D;
        private const double DefaultPageHeightPoints = 841.9D;
        private const double DefaultMarginPoints = 72D;
        private const double ParagraphGapPoints = 6D;
        private const double DefaultCellMarginPoints = 5.4D;
        private const double MinimumTableRowHeightPoints = 22D;

        internal static OfficeImageExportResult Render(
            WordDocument document,
            OfficeImageExportFormat format,
            WordImageExportOptions options,
            CancellationToken cancellationToken = default) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            if (options == null) {
                throw new ArgumentNullException(nameof(options));
            }

            cancellationToken.ThrowIfCancellationRequested();
            WordDocumentVisualSnapshot snapshot = CreateSnapshot(document, options, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            return RenderSnapshot(snapshot, format, options, cancellationToken);
        }

        private static OfficeImageExportResult RenderSnapshot(WordDocumentVisualSnapshot snapshot,
            OfficeImageExportFormat format, WordImageExportOptions options, CancellationToken cancellationToken = default) {
            OfficeDrawing drawing = snapshot.Drawing;

            if (format == OfficeImageExportFormat.Svg) {
                List<OfficeImageExportDiagnostic> diagnostics = new List<OfficeImageExportDiagnostic>(snapshot.Diagnostics);
                var fallbackCodec = new OfficeRasterImageFallbackCodec(options.ImageCodec, diagnostics, "Word document");
                byte[] svg = OfficeDrawingSvgExporter.ToSvgBytes(drawing, options.Scale, OfficeSvgSizeUnit.Pixel, fallbackCodec);
                return options.EnsureAccepted(new OfficeImageExportResult(format, ScaledWidth(drawing, options), ScaledHeight(drawing, options), svg, "Page " + (options.PageIndex + 1), "Word document", diagnostics));
            }

            if (format.IsRaster()) {
                List<OfficeImageExportDiagnostic> diagnostics = new List<OfficeImageExportDiagnostic>(snapshot.Diagnostics);
                const string source = "Word document";
                OfficeRasterExportPlan plan = OfficeRasterExportPlanner.Resolve(
                    drawing.Width,
                    drawing.Height,
                    format,
                    options,
                    source);
                if (plan.Diagnostic != null) diagnostics.Add(plan.Diagnostic);
                var fallbackCodec = new OfficeRasterImageFallbackCodec(options.ImageCodec, diagnostics, source);
                OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing, new OfficeDrawingRasterRenderOptions {
                    Scale = plan.Limit.Scale,
                    Background = options.BackgroundColor,
                    ImageCodec = fallbackCodec,
                    TextShapingProvider = options.TextShapingProvider,
                    TextShapingLanguage = options.TextShapingLanguage,
                    DiagnosticSink = diagnostics,
                    DiagnosticSource = source,
                    CancellationToken = cancellationToken
                });
                byte[] bytes = OfficeRasterImageEncoder.Encode(
                    image,
                    format,
                    plan.CreateEncodingOptions());
                cancellationToken.ThrowIfCancellationRequested();
                return options.EnsureAccepted(new OfficeImageExportResult(format, image.Width, image.Height, bytes, "Page " + (options.PageIndex + 1), source, diagnostics));
            }

            throw new ArgumentOutOfRangeException(nameof(format), format, "Unsupported image export format.");
        }

        internal static WordDocumentVisualSnapshot CreateSnapshot(
            WordDocument document,
            WordImageExportOptions options,
            CancellationToken cancellationToken = default) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            if (options == null) {
                throw new ArgumentNullException(nameof(options));
            }

            cancellationToken.ThrowIfCancellationRequested();
            return CreateSnapshot(
                document,
                options,
                EstimateSectionPageCounts(
                    document,
                    cancellationToken,
                    options.CancellationCheckpoint),
                cancellationToken);
        }

        private static WordDocumentVisualSnapshot CreateSnapshot(WordDocument document,
            WordImageExportOptions options,
            IReadOnlyList<int> sectionPageCounts,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            List<OfficeImageExportDiagnostic> diagnostics = new List<OfficeImageExportDiagnostic>();
            List<WordDocumentVisualFragment> fragments = new List<WordDocumentVisualFragment>();
            OfficeDrawing drawing = CreateDrawing(
                document,
                options,
                diagnostics,
                fragments,
                sectionPageCounts,
                cancellationToken);
            drawing.AppendFontDiagnostics(diagnostics, "Word document");
            return new WordDocumentVisualSnapshot(
                drawing,
                options.PageIndex,
                diagnostics.AsReadOnly(),
                fragments.AsReadOnly());
        }

        private static OfficeDrawing CreateDrawing(WordDocument document, WordImageExportOptions options,
            List<OfficeImageExportDiagnostic> diagnostics,
            List<WordDocumentVisualFragment> fragments,
            IReadOnlyList<int> sectionPageCounts,
            CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            IReadOnlyList<int> sectionPageNumberStarts = ResolveSectionPageNumberStarts(document, sectionPageCounts);
            WordImagePageContext pageContext = ResolvePageContext(document, options.PageIndex, sectionPageCounts);
            (double width, double height) = GetPageSizePoints(pageContext.Section);
            OfficeDrawing drawing = new OfficeDrawing(width, height).ApplyImageExportOptions(options);
            AddBackgroundRectangle(drawing, options.BackgroundColor);

            if (options.IncludeDocumentContent) {
                IReadOnlyDictionary<OpenXmlElement, WordImageSourceBlock> sourceBlocks =
                    BuildImageSourceBlocks(document, cancellationToken);
                int totalPageCount = Math.Max(1, sectionPageCounts.Sum());
                int sectionPageCount = pageContext.SectionIndex < sectionPageCounts.Count
                    ? sectionPageCounts[pageContext.SectionIndex]
                    : 1;
                int sectionPageNumberStart = pageContext.SectionIndex < sectionPageNumberStarts.Count
                    ? sectionPageNumberStarts[pageContext.SectionIndex]
                    : options.PageIndex + 1;
                (int pageNumberValue, string pageNumberText) = pageContext.Section != null
                    ? ResolveSectionPageNumber(pageContext.Section, sectionPageNumberStart, pageContext.SectionPageIndex)
                    : (options.PageIndex + 1, (options.PageIndex + 1).ToString(CultureInfo.InvariantCulture));
                WordHeaderFooterPageFrame? headerFooterFrame = pageContext.Section != null
                    ? AddSupportedHeaderFooterContent(pageContext.Section, drawing, diagnostics, options.PageIndex, pageContext.SectionIndex, sectionPageNumberStart, pageContext.SectionPageIndex, totalPageCount, sectionPageCount)
                    : null;

                AddSupportedBodyContent(
                    document,
                    drawing,
                    diagnostics,
                    pageContext.SectionIndex,
                    pageContext.SectionPageIndex,
                    pageContext.Section,
                    resolveDynamicPageFields: true,
                    totalPageCount: totalPageCount,
                    sectionNumber: pageContext.SectionIndex + 1,
                    sectionPageCount: sectionPageCount,
                    pageNumberValue: pageNumberValue,
                    pageNumberText: pageNumberText,
                    contentTop: headerFooterFrame?.BodyTop,
                    contentBottom: headerFooterFrame?.BodyBottom,
                    sectionPageCounts: sectionPageCounts,
                    bodyFrameProvider: pageContext.Section != null
                        ? CreateBodyFrameProvider(pageContext.Section, drawing, pageContext.SectionIndex, sectionPageNumberStart, totalPageCount, sectionPageCount, pageContext.SectionPageIndex, headerFooterFrame)
                        : null,
                    sourceBlocks: sourceBlocks,
                    fragments: fragments,
                    cancellationToken: cancellationToken,
                    cancellationCheckpoint: options.CancellationCheckpoint);
            }

            return drawing;
        }

        private static IReadOnlyList<int> ResolveSectionPageNumberStarts(WordDocument document, IReadOnlyList<int> sectionPageCounts) {
            var starts = new List<int>(document.Sections.Count);
            int nextImplicitStart = 1;
            for (int i = 0; i < document.Sections.Count; i++) {
                PageNumberType? pageNumberType = document.Sections[i]._sectionProperties.GetFirstChild<PageNumberType>();
                int start = pageNumberType?.Start?.Value ?? nextImplicitStart;
                start = Math.Max(1, start);
                starts.Add(start);

                int sectionPages = i < sectionPageCounts.Count ? Math.Max(0, sectionPageCounts[i]) : 1;
                nextImplicitStart = Math.Max(1, start + sectionPages);
            }

            return starts;
        }

        private static WordImageFlowContext AddSupportedBodyContent(
            WordDocument document,
            OfficeDrawing drawing,
            List<OfficeImageExportDiagnostic> diagnostics,
            int sectionIndex,
            int sectionPageIndex,
            WordSection? pageSection,
            bool resolveDynamicPageFields = false,
            int totalPageCount = 1,
            int sectionNumber = 1,
            int sectionPageCount = 1,
            int pageNumberValue = 0,
            string? pageNumberText = null,
            double? contentTop = null,
            double? contentBottom = null,
            IReadOnlyList<int>? sectionPageCounts = null,
            Func<int, WordImageBodyFrame>? bodyFrameProvider = null,
            IReadOnlyDictionary<OpenXmlElement, WordImageSourceBlock>? sourceBlocks = null,
            List<WordDocumentVisualFragment>? fragments = null,
            CancellationToken cancellationToken = default,
            Action<WordImageCancellationCheckpoint>? cancellationCheckpoint = null) {
            cancellationToken.ThrowIfCancellationRequested();
            int targetPageIndex = Math.Max(0, sectionPageIndex);
            WordImageFlowContext context = CreateFlowContext(
                pageSection,
                drawing,
                targetPageIndex,
                resolveDynamicPageFields,
                totalPageCount,
                sectionNumber,
                sectionPageCount,
                pageNumberValue,
                    pageNumberText,
                    contentTop,
                    contentBottom,
                    bodyFrameProvider,
                    sourceBlocks,
                    fragments,
                    cancellationToken,
                    cancellationCheckpoint);
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers = DocumentTraversal.BuildListMarkers(document);
            IReadOnlyList<WordSectionBodyElement> bodyEntries = GetSectionBodyElementEntries(
                document,
                sectionIndex,
                cancellationToken);
            IReadOnlyList<OpenXmlElement> bodyChildren = bodyEntries.Select(entry => entry.Element).ToList();
            for (int index = 0; index < bodyChildren.Count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                WordSectionBodyElement entry = bodyEntries[index];
                OpenXmlElement element = entry.Element;
                if (entry.SectionIndex != sectionIndex) {
                    int mergedSectionPageCount = ResolveSectionPageCountForFieldContext(sectionPageCounts, entry.SectionIndex, sectionPageCount);
                    context.UpdateSectionContext(entry.SectionIndex + 1, mergedSectionPageCount);
                }

                TryAdvanceForKeepWithNext(document, bodyChildren, index, context, listMarkers);
                if (context.PastTargetPage || context.StoppedForPagination) {
                    break;
                }

                bool added = AddBodyElementContent(document, element, context, diagnostics, listMarkers);

                if (context.PastTargetPage || context.StoppedForPagination) {
                    break;
                }

                if (context.IsTargetPage && !added && element is Paragraph) {
                    context.ClearParagraphSpacingState();
                    context.Y += ParagraphGapPoints;
                }
            }

            if (context.PageIndex < targetPageIndex) {
                AddDiagnostic(diagnostics, "unsupported-word-page-index", "Rendered a blank page because the requested Word page is beyond the currently estimated Word document content.");
            }

            return context;
        }

        private static int ResolveSectionPageCountForFieldContext(IReadOnlyList<int>? sectionPageCounts, int sectionIndex, int fallbackSectionPageCount) {
            if (sectionPageCounts == null || sectionIndex < 0 || sectionIndex >= sectionPageCounts.Count) {
                return Math.Max(1, fallbackSectionPageCount);
            }

            int sectionPageCount = sectionPageCounts[sectionIndex];
            if (sectionPageCount > 0) {
                return sectionPageCount;
            }

            for (int i = sectionIndex - 1; i >= 0; i--) {
                if (sectionPageCounts[i] > 0) {
                    return sectionPageCounts[i];
                }
            }

            return Math.Max(1, fallbackSectionPageCount);
        }

        private static int EstimateDocumentPageCount(WordDocument document, WordSection? pageSection) {
            (double width, double height) = GetPageSizePoints(pageSection);
            var drawing = new OfficeDrawing(width, height);
            var diagnostics = new List<OfficeImageExportDiagnostic>();
            WordImageFlowContext context = AddSupportedBodyContent(document, drawing, diagnostics, 0, int.MaxValue, pageSection);
            return Math.Max(1, context.PageIndex + 1);
        }

        private static WordImageFlowContext CreateFlowContext(
            WordSection? section,
            OfficeDrawing drawing,
            int targetPageIndex = 0,
            bool resolveDynamicPageFields = false,
            int totalPageCount = 1,
            int sectionNumber = 1,
            int sectionPageCount = 1,
            int pageNumberValue = 0,
            string? pageNumberText = null,
            double? contentTop = null,
            double? contentBottom = null,
            Func<int, WordImageBodyFrame>? bodyFrameProvider = null,
            IReadOnlyDictionary<OpenXmlElement, WordImageSourceBlock>? sourceBlocks = null,
            List<WordDocumentVisualFragment>? fragments = null,
            CancellationToken cancellationToken = default,
            Action<WordImageCancellationCheckpoint>? cancellationCheckpoint = null) {
            WordMargins? margins = section?.Margins;
            double left = ToPoints(margins?.Left?.Value, DefaultMarginPoints);
            double right = ToPoints(margins?.Right?.Value, DefaultMarginPoints);
            double top = Math.Min(drawing.Height, Math.Max(0D, contentTop ?? ToPoints(margins?.Top, DefaultMarginPoints)));
            double bottom = ToPoints(margins?.Bottom, DefaultMarginPoints);
            double contentWidth = Math.Max(1D, drawing.Width - left - right);
            double resolvedContentBottom = contentBottom.HasValue
                ? Math.Min(drawing.Height, Math.Max(top, contentBottom.Value))
                : Math.Max(top, drawing.Height - bottom);
            IReadOnlyList<WordImageColumnFrame> columns = CreateColumnFrames(section, left, contentWidth);
            return new WordImageFlowContext(
                drawing,
                left,
                top,
                contentWidth,
                resolvedContentBottom,
                columns,
                targetPageIndex: targetPageIndex,
                allowPageAdvanceForOverflow: true,
                resolveDynamicPageFields: resolveDynamicPageFields,
                totalPageCount: totalPageCount,
                sectionNumber: sectionNumber,
                sectionPageCount: sectionPageCount,
                pageNumberValue: pageNumberValue,
                pageNumberText: pageNumberText,
                bodyFrameProvider: bodyFrameProvider,
                sourceBlocks: sourceBlocks,
                fragments: fragments,
                cancellationToken: cancellationToken,
                cancellationCheckpoint: cancellationCheckpoint);
        }

        private static WordImageFlowContext CreateFlowContext(
            OfficeDrawing drawing,
            double left,
            double y,
            double contentWidth,
            double contentBottom,
            string overflowDiagnosticCode,
            string overflowDiagnosticMessage,
            int targetPageIndex = 0,
            int initialPageIndex = 0,
            bool resolveDynamicPageFields = false,
            int totalPageCount = 1,
            int sectionNumber = 1,
            int sectionPageCount = 1,
            int pageNumberValue = 0,
            string? pageNumberText = null,
            CancellationToken cancellationToken = default) =>
            new WordImageFlowContext(
                drawing,
                left,
                y,
                contentWidth,
                contentBottom,
                Array.Empty<WordImageColumnFrame>(),
                overflowDiagnosticCode,
                overflowDiagnosticMessage,
                targetPageIndex,
                allowPageAdvanceForOverflow: false,
                initialPageIndex,
                resolveDynamicPageFields,
                totalPageCount,
                sectionNumber,
                sectionPageCount,
                pageNumberValue,
                pageNumberText,
                cancellationToken: cancellationToken);

        private static bool AddParagraphContent(
            WordDocument document,
            Paragraph paragraph,
            WordImageFlowContext context,
            List<OfficeImageExportDiagnostic> diagnostics,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers) {
            context.ThrowIfCancellationRequested();
            bool added = false;
            var colorScheme = GetDocumentColorScheme(document);
            WordImageListMarker? listMarker = CreateListMarker(document, paragraph, listMarkers);
            bool markerRendered = false;
            var textRuns = new List<WordParagraph>();

            bool FlushTextRuns() {
                context.ThrowIfCancellationRequested();
                if (textRuns.Count == 0) {
                    textRuns.Clear();
                    return false;
                }

                WordImageListMarker? currentMarker = markerRendered ? null : listMarker;
                bool runAdded = textRuns.Count == 1 && !HasRunHighlight(textRuns[0])
                    ? AddTextRun(textRuns[0], context, diagnostics, currentMarker, colorScheme)
                    : AddRichTextRuns(textRuns, context, diagnostics, currentMarker, colorScheme);
                if (runAdded && currentMarker.HasValue) {
                    markerRendered = true;
                }

                textRuns.Clear();
                return runAdded;
            }

            foreach (WordParagraph run in WordSection.ConvertParagraphToWordParagraphs(
                         document,
                         paragraph,
                         splitPaginationMarkers: true,
                         context.CancellationToken)) {
                context.ThrowIfCancellationRequested();
                if (run.IsPageBreak) {
                    added |= FlushTextRuns();
                    context.AdvancePage();
                    if (context.PastTargetPage) {
                        return added;
                    }

                    continue;
                }

                if (run.IsColumnBreak) {
                    added |= FlushTextRuns();
                    context.AdvanceColumnOrPage();
                    if (context.PastTargetPage) {
                        return added;
                    }

                    continue;
                }

                WordImage? image = run.Image;
                if (image != null) {
                    added |= FlushTextRuns();
                    context.ClearParagraphSpacingState();
                    added |= AddImage(image, context, diagnostics);
                    context.ClearParagraphSpacingState();
                    continue;
                }

                WordSmartArt? smartArt = run.SmartArt;
                if (smartArt != null) {
                    added |= FlushTextRuns();
                    context.ClearParagraphSpacingState();
                    added |= AddSmartArt(smartArt, context, diagnostics);
                    context.ClearParagraphSpacingState();
                    continue;
                }

                WordShape? shape = run.Shape;
                if (shape != null) {
                    added |= FlushTextRuns();
                    context.ClearParagraphSpacingState();
                    added |= AddShape(shape, context, diagnostics);
                    context.ClearParagraphSpacingState();
                    continue;
                }

                WordTextBox? textBox = run.TextBox;
                if (textBox != null) {
                    added |= FlushTextRuns();
                    context.ClearParagraphSpacingState();
                    added |= AddTextBox(textBox, context, diagnostics, colorScheme);
                    context.ClearParagraphSpacingState();
                    continue;
                }

                string text = ResolveImageExportText(run, context);
                if (string.IsNullOrEmpty(text)) {
                    continue;
                }

                textRuns.Add(run);
            }

            added |= FlushTextRuns();
            return added;
        }

        private static bool AddRichTextRuns(IReadOnlyList<WordParagraph> paragraphs, WordImageFlowContext context, List<OfficeImageExportDiagnostic> diagnostics, WordImageListMarker? listMarker, DocumentFormat.OpenXml.Drawing.ColorScheme? colorScheme) {
            List<OfficeRichTextRun> richRuns = CreateRichTextRuns(paragraphs, colorScheme, context);
            if (richRuns.Count == 0) {
                return false;
            }

            double maxFontSize = 10D;
            for (int i = 0; i < richRuns.Count; i++) {
                maxFontSize = Math.Max(maxFontSize, richRuns[i].FontSize);
            }

            double lineHeight = Math.Max(maxFontSize * 1.25D, 12D);
            WordImageTextLayout textLayout = ResolveTextLayout(context, listMarker, paragraphs[0]);
            double height = EstimateRichTextHeight(
                richRuns,
                maxFontSize,
                textLayout.ContentWidth,
                lineHeight,
                textLayout.ParagraphIndent,
                context.CancellationToken);
            int estimatedLineCount = Math.Max(1, (int)Math.Ceiling(height / lineHeight));
            WordParagraphSpacing spacing = ResolveParagraphSpacing(paragraphs[0], maxFontSize, lineHeight, context, out WordParagraphSpacingState spacingState);
            bool keepLinesTogether = ResolveKeepLinesTogether(paragraphs[0]);
            bool avoidWidowAndOrphan = ResolveAvoidWidowAndOrphan(paragraphs[0]);
            if (spacing.Before + height > context.ContentHeight) {
                return AddPaginatedRichTextRun(paragraphs[0], richRuns, maxFontSize, lineHeight, spacing, spacingState, listMarker, colorScheme, avoidWidowAndOrphan, context, diagnostics);
            }

            if (!keepLinesTogether && ShouldPaginateParagraphOverflow(context, spacing.Before + height, estimatedLineCount)) {
                return AddPaginatedRichTextRun(paragraphs[0], richRuns, maxFontSize, lineHeight, spacing, spacingState, listMarker, colorScheme, avoidWidowAndOrphan, context, diagnostics);
            }

            if (!EnsureVerticalSpace(context, spacing.Before + height, diagnostics)) {
                return false;
            }

            context.Y += spacing.Before;
            if (context.IsTargetPage) {
                AddParagraphFrame(paragraphs[0], context, textLayout, height, colorScheme);
                if (listMarker.HasValue) {
                    WordImageListMarker marker = listMarker.Value;
                    context.Drawing.AddText(
                        marker.Marker,
                        textLayout.MarkerLeft,
                        context.Y,
                        textLayout.MarkerWidth,
                        lineHeight,
                        marker.Font,
                        marker.Color,
                        marker.Alignment,
                        lineHeight,
                        wrapText: false);
                }

                context.Drawing.AddRichText(
                    richRuns,
                    textLayout.TextLeft,
                    context.Y,
                    textLayout.TextWidth,
                    height,
                    MapTextAlignment(paragraphs[0].ParagraphAlignment),
                    lineHeight,
                    wrapText: true,
                    padding: textLayout.Padding,
                    paragraphIndent: textLayout.ParagraphIndent);
            }

            context.Y += height + spacing.After;
            context.SetParagraphSpacingState(spacingState);
            return true;
        }

        private static bool AddTextRun(WordParagraph paragraph, WordImageFlowContext context, List<OfficeImageExportDiagnostic> diagnostics, WordImageListMarker? listMarker, DocumentFormat.OpenXml.Drawing.ColorScheme? colorScheme) {
            OfficeFontInfo font = CreateFont(paragraph);
            double lineHeight = Math.Max(font.Size * 1.25D, 12D);
            WordImageTextLayout textLayout = ResolveTextLayout(context, listMarker, paragraph);
            string text = ResolveImageExportText(paragraph, context);
            List<string> lines = WrapTextIntoMeasuredLines(
                text,
                font,
                textLayout.LayoutWidth,
                context.CancellationToken,
                context.CancellationCheckpoint);
            double height = Math.Max(lineHeight, lines.Count * lineHeight);
            WordParagraphSpacing spacing = ResolveParagraphSpacing(paragraph, font.Size, lineHeight, context, out WordParagraphSpacingState spacingState);
            bool keepLinesTogether = ResolveKeepLinesTogether(paragraph);
            bool avoidWidowAndOrphan = ResolveAvoidWidowAndOrphan(paragraph);
            if (spacing.Before + height <= context.ContentHeight) {
                if (!keepLinesTogether && ShouldPaginateParagraphOverflow(context, spacing.Before + height, lines.Count)) {
                    return AddPaginatedTextRun(paragraph, lines, font, lineHeight, spacing, spacingState, listMarker, colorScheme, avoidWidowAndOrphan, context, diagnostics);
                }

                return AddTextRunBlock(paragraph, font, lineHeight, height, spacing, spacingState, textLayout, listMarker, colorScheme, context, diagnostics);
            }

            return AddPaginatedTextRun(paragraph, lines, font, lineHeight, spacing, spacingState, listMarker, colorScheme, avoidWidowAndOrphan, context, diagnostics);
        }

        private static List<OfficeRichTextRun> CreateRichTextRuns(IReadOnlyList<WordParagraph> paragraphs, DocumentFormat.OpenXml.Drawing.ColorScheme? colorScheme, WordImageFlowContext? context = null) {
            var richRuns = new List<OfficeRichTextRun>(paragraphs.Count);
            for (int i = 0; i < paragraphs.Count; i++) {
                context?.ThrowIfCancellationRequested();
                WordParagraph paragraph = paragraphs[i];
                string text = ResolveImageExportText(paragraph, context);
                if (string.IsNullOrEmpty(text)) {
                    continue;
                }

                richRuns.Add(CreateRichTextRun(paragraph, colorScheme, text));
            }

            return richRuns;
        }

        private static OfficeRichTextRun CreateRichTextRun(WordParagraph paragraph, DocumentFormat.OpenXml.Drawing.ColorScheme? colorScheme, string? text = null) {
            double fontSize = paragraph.FontSize ?? 11;
            return new OfficeRichTextRun(
                text ?? paragraph.Text,
                fontSize,
                ResolveParagraphTextColor(paragraph, colorScheme),
                paragraph.Bold,
                paragraph.Italic,
                paragraph.Underline.HasValue && paragraph.Underline.Value != UnderlineValues.None,
                paragraph.FontFamily ?? "Calibri",
                paragraph.Strike || paragraph.DoubleStrike,
                ResolveRunHighlightColor(ResolveRunHighlight(paragraph)));
        }

        private static OfficeFontInfo CreateFont(WordParagraph paragraph) {
            OfficeFontStyle style = OfficeFontStyle.Regular;
            if (paragraph.Bold) {
                style |= OfficeFontStyle.Bold;
            }

            if (paragraph.Italic) {
                style |= OfficeFontStyle.Italic;
            }

            if (paragraph.Underline.HasValue && paragraph.Underline.Value != UnderlineValues.None) {
                style |= OfficeFontStyle.Underline;
            }

            if (paragraph.Strike || paragraph.DoubleStrike) {
                style |= OfficeFontStyle.Strikethrough;
            }

            return new OfficeFontInfo(paragraph.FontFamily ?? "Calibri", paragraph.FontSize ?? 11, style);
        }

        private static OfficeTextAlignment MapTextAlignment(JustificationValues? alignment) {
            if (alignment == JustificationValues.Center) {
                return OfficeTextAlignment.Center;
            }

            if (alignment == JustificationValues.Right || alignment == JustificationValues.End) {
                return OfficeTextAlignment.Right;
            }

            if (alignment == JustificationValues.Both ||
                alignment == JustificationValues.Distribute ||
                alignment == JustificationValues.ThaiDistribute ||
                alignment == JustificationValues.HighKashida ||
                alignment == JustificationValues.MediumKashida ||
                alignment == JustificationValues.LowKashida) {
                return OfficeTextAlignment.Justify;
            }

            return OfficeTextAlignment.Left;
        }

        private static double EstimateTextHeight(
            string text,
            double fontSize,
            double contentWidth,
            double lineHeight,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            double averageCharacterWidth = Math.Max(1D, fontSize * 0.52D);
            int charactersPerLine = Math.Max(1, (int)Math.Floor(contentWidth / averageCharacterWidth));
            string normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
            cancellationToken.ThrowIfCancellationRequested();
            string[] explicitLines = normalized.Split('\n');
            int lineCount = 0;
            foreach (string line in explicitLines) {
                cancellationToken.ThrowIfCancellationRequested();
                lineCount += Math.Max(1, (int)Math.Ceiling(line.Length / (double)charactersPerLine));
            }

            return Math.Max(lineHeight, lineCount * lineHeight);
        }

        private static double EstimateRichTextHeight(
            IReadOnlyList<OfficeRichTextRun> runs,
            double maxFontSize,
            double contentWidth,
            double lineHeight,
            OfficeTextParagraphIndent paragraphIndent,
            CancellationToken cancellationToken) {
            double lineHeightFactor = Math.Max(1D, lineHeight / Math.Max(1D, maxFontSize));
            OfficeRichTextBlockLayout layout = OfficeTextLayoutEngine.LayoutRichTextBlock(
                runs,
                contentWidth,
                double.MaxValue,
                lineHeightFactor,
                CreateRichTextMeasure(cancellationToken),
                wrap: true,
                shrinkToFit: false,
                minimumFontSize: Math.Min(6D, maxFontSize),
                overflowBehavior: OfficeTextOverflowBehavior.Clip,
                paragraphIndent: paragraphIndent,
                cancellationToken: cancellationToken);
            return Math.Max(lineHeight, layout.Height);
        }

        private static bool ShouldPaginateParagraphOverflow(WordImageFlowContext context, double height, int lineCount) =>
            lineCount > 1 &&
            context.CanAdvancePageForOverflow &&
            context.Y > context.Top &&
            context.Y + height > context.ContentBottom;

        private static Func<string?, double, string?, double> CreateRichTextMeasure(
            CancellationToken cancellationToken = default) {
            OfficeTextMeasurer measurer = OfficeTextMeasurer.Create();
            return (value, size, family) => {
                    cancellationToken.ThrowIfCancellationRequested();
                    OfficeTextMeasurementStyle measuredStyle = measurer.CreateStyle(new OfficeFontInfo(family, size));
                    return measurer.MeasureWidth(value, measuredStyle);
                };
        }

        private static void AddBackgroundRectangle(OfficeDrawing drawing, OfficeColor fillColor) {
            OfficeShape shape = OfficeShape.Rectangle(drawing.Width, drawing.Height);
            shape.FillColor = fillColor;
            shape.StrokeWidth = 0D;
            drawing.AddShape(shape, 0D, 0D);
        }

        private static (double Width, double Height) GetPageSizePoints(WordDocument document) {
            return GetPageSizePoints(document.Sections.FirstOrDefault());
        }

        private static (double Width, double Height) GetPageSizePoints(WordSection? section) {
            WordPageSizes? pageSettings = section?.PageSettings;
            double width = ToPoints(pageSettings?.Width?.Value, DefaultPageWidthPoints);
            double height = ToPoints(pageSettings?.Height?.Value, DefaultPageHeightPoints);
            return (Math.Max(1D, width), Math.Max(1D, height));
        }

        private static WordImagePageContext ResolvePageContext(WordDocument document, int pageIndex, IReadOnlyList<int> sectionPageCounts) {
            if (document.Sections.Count == 0) {
                return new WordImagePageContext(null, 0, 0);
            }

            int targetPageIndex = Math.Max(0, pageIndex);
            int firstPageInSection = 0;
            for (int sectionIndex = 0; sectionIndex < document.Sections.Count; sectionIndex++) {
                int sectionPages = sectionIndex < sectionPageCounts.Count
                    ? sectionPageCounts[sectionIndex]
                    : 1;
                if (sectionPages <= 0) {
                    continue;
                }

                if (targetPageIndex < firstPageInSection + sectionPages) {
                    return new WordImagePageContext(document.Sections[sectionIndex], sectionIndex, targetPageIndex - firstPageInSection);
                }

                firstPageInSection += sectionPages;
            }

            int lastSectionIndex = FindLastRenderableSectionIndex(document, sectionPageCounts);
            return new WordImagePageContext(document.Sections[lastSectionIndex], lastSectionIndex, Math.Max(0, targetPageIndex - firstPageInSection));
        }

        private static int FindLastRenderableSectionIndex(WordDocument document, IReadOnlyList<int> sectionPageCounts) {
            for (int sectionIndex = Math.Min(document.Sections.Count, sectionPageCounts.Count) - 1; sectionIndex >= 0; sectionIndex--) {
                if (sectionPageCounts[sectionIndex] > 0) {
                    return sectionIndex;
                }
            }

            return document.Sections.Count - 1;
        }

        private static bool StartsNewPage(SectionProperties? sectionProperties) {
            if (sectionProperties == null) {
                return false;
            }

            SectionMarkValues? value = ResolveSectionMark(sectionProperties);
            return value == null ||
                   value == SectionMarkValues.NextPage ||
                   value == SectionMarkValues.OddPage ||
                   value == SectionMarkValues.EvenPage;
        }

        private static bool IsNextColumnSectionBreak(SectionProperties? sectionProperties) =>
            ResolveSectionMark(sectionProperties) == SectionMarkValues.NextColumn;

        private static int CountSectionBreakPageAdvance(int currentPageIndex, SectionProperties? sectionProperties) {
            SectionMarkValues? value = ResolveSectionMark(sectionProperties);
            if (value == SectionMarkValues.NextColumn) {
                return 1;
            }

            if (value != SectionMarkValues.OddPage && value != SectionMarkValues.EvenPage) {
                return StartsNewPage(sectionProperties) ? 1 : 0;
            }

            int nextPageIndex = currentPageIndex + 1;
            bool nextPageIsEven = !IsOddWordPageIndex(nextPageIndex);
            if ((value == SectionMarkValues.OddPage && !nextPageIsEven) ||
                (value == SectionMarkValues.EvenPage && nextPageIsEven)) {
                return 1;
            }

            return 2;
        }

        private static SectionMarkValues? ResolveSectionMark(SectionProperties? sectionProperties) =>
            sectionProperties?.GetFirstChild<SectionType>()?.Val?.Value;

        private static bool IsOddWordPageIndex(int pageIndex) =>
            pageIndex % 2 == 0;

        private static double ToPoints(uint? twips, double fallbackPoints) =>
            twips.HasValue ? twips.Value / TwipsPerPoint : fallbackPoints;

        private static double ToPoints(int? twips, double fallbackPoints) =>
            twips.HasValue ? twips.Value / TwipsPerPoint : fallbackPoints;

        private static double ToPoints(short? twips, double fallbackPoints) =>
            twips.HasValue ? twips.Value / TwipsPerPoint : fallbackPoints;

        private static bool EnsureVerticalSpace(WordImageFlowContext context, double height, List<OfficeImageExportDiagnostic> diagnostics) {
            if (context.Y + height <= context.ContentBottom) {
                return true;
            }

            if (context.CanAdvancePageForOverflow) {
                context.AdvanceColumnOrPage();
                if (context.Y + height <= context.ContentBottom) {
                    return !context.PastTargetPage;
                }
            }

            if (context.PastTargetPage) {
                return false;
            }

            if (!context.StoppedForPagination) {
                AddDiagnostic(diagnostics, context.OverflowDiagnosticCode, context.OverflowDiagnosticMessage);
                context.StoppedForPagination = true;
            }

            return false;
        }

        private static int ScaledWidth(OfficeDrawing drawing, WordImageExportOptions options) =>
            Math.Max(1, (int)Math.Ceiling(drawing.Width * options.Scale));

        private static int ScaledHeight(OfficeDrawing drawing, WordImageExportOptions options) =>
            Math.Max(1, (int)Math.Ceiling(drawing.Height * options.Scale));

        private static int UnscaledWidth(OfficeDrawing drawing) =>
            Math.Max(1, (int)Math.Ceiling(drawing.Width));

        private static int UnscaledHeight(OfficeDrawing drawing) =>
            Math.Max(1, (int)Math.Ceiling(drawing.Height));

        private static void AddDiagnostic(List<OfficeImageExportDiagnostic> diagnostics, string code, string message, string? source = null) {
            diagnostics.Add(WordImageExportDiagnosticClassifier.Create(
                OfficeImageExportDiagnosticSeverity.Warning,
                code,
                message,
                string.IsNullOrWhiteSpace(source) ? "Word document" : source));
        }

        private readonly struct WordImagePageContext {
            internal WordImagePageContext(WordSection? section, int sectionIndex, int sectionPageIndex) {
                Section = section;
                SectionIndex = Math.Max(0, sectionIndex);
                SectionPageIndex = Math.Max(0, sectionPageIndex);
            }

            internal WordSection? Section { get; }

            internal int SectionIndex { get; }

            internal int SectionPageIndex { get; }
        }

        private sealed class WordImageFlowContext {
            internal WordImageFlowContext(
                OfficeDrawing drawing,
                double left,
                double top,
                double contentWidth,
                double contentBottom,
                IReadOnlyList<WordImageColumnFrame> columns,
                string overflowDiagnosticCode = "unsupported-word-pagination",
                string overflowDiagnosticMessage = "Stopped rendering Word content because it does not fit within the current dependency-free page layout frame.",
                int targetPageIndex = 0,
                bool allowPageAdvanceForOverflow = false,
                int initialPageIndex = 0,
                bool resolveDynamicPageFields = false,
                int totalPageCount = 1,
                int sectionNumber = 1,
                int sectionPageCount = 1,
                int pageNumberValue = 0,
                string? pageNumberText = null,
                Func<int, WordImageBodyFrame>? bodyFrameProvider = null,
                IReadOnlyDictionary<OpenXmlElement, WordImageSourceBlock>? sourceBlocks = null,
                List<WordDocumentVisualFragment>? fragments = null,
                CancellationToken cancellationToken = default,
                Action<WordImageCancellationCheckpoint>? cancellationCheckpoint = null) {
                Drawing = drawing;
                Left = left;
                Top = top;
                Y = top;
                ContentWidth = contentWidth;
                ContentBottom = contentBottom;
                Columns = columns;
                OverflowDiagnosticCode = overflowDiagnosticCode;
                OverflowDiagnosticMessage = overflowDiagnosticMessage;
                PageIndex = Math.Max(0, initialPageIndex);
                TargetPageIndex = Math.Max(0, targetPageIndex);
                AllowPageAdvanceForOverflow = allowPageAdvanceForOverflow;
                ResolveDynamicPageFields = resolveDynamicPageFields;
                TotalPageCount = Math.Max(1, totalPageCount);
                SectionNumber = Math.Max(1, sectionNumber);
                SectionPageCount = Math.Max(1, sectionPageCount);
                PageNumberValue = pageNumberValue > 0 ? pageNumberValue : PageIndex + 1;
                PageNumberText = pageNumberText ?? PageNumberValue.ToString(CultureInfo.InvariantCulture);
                BodyFrameProvider = bodyFrameProvider;
                SourceBlocks = sourceBlocks;
                Fragments = fragments;
                CancellationToken = cancellationToken;
                CancellationCheckpoint = cancellationCheckpoint;
                ApplyBodyFrame(PageIndex);
                ApplyColumn(0);
            }

            internal OfficeDrawing Drawing { get; }

            internal double Left { get; private set; }

            internal double Top { get; private set; }

            internal double Y { get; set; }

            internal double ContentWidth { get; private set; }

            internal double ContentBottom { get; private set; }

            internal double ContentHeight => Math.Max(0D, ContentBottom - Top);

            internal bool StoppedForPagination { get; set; }

            internal bool IsAtPageFrameStart => ColumnIndex == 0 && Math.Abs(Y - Top) < 0.001D;

            internal int PageIndex { get; private set; }

            internal int TargetPageIndex { get; }

            internal bool IsTargetPage => PageIndex == TargetPageIndex;

            internal bool PastTargetPage => PageIndex > TargetPageIndex;

            internal bool CanAdvancePageForOverflow => AllowPageAdvanceForOverflow && PageIndex <= TargetPageIndex;

            internal bool ResolveDynamicPageFields { get; }

            internal int TotalPageCount { get; }

            internal int SectionNumber { get; private set; }

            internal int SectionPageCount { get; private set; }

            internal int PageNumberValue { get; }

            internal string PageNumberText { get; }

            internal CancellationToken CancellationToken { get; }

            internal Action<WordImageCancellationCheckpoint>? CancellationCheckpoint { get; }

            internal void ThrowIfCancellationRequested() =>
                CancellationToken.ThrowIfCancellationRequested();

            internal void UpdateSectionContext(int sectionNumber, int sectionPageCount) {
                SectionNumber = Math.Max(1, sectionNumber);
                SectionPageCount = Math.Max(1, sectionPageCount);
            }

            private bool AllowPageAdvanceForOverflow { get; }

            private IReadOnlyList<WordImageColumnFrame> Columns { get; }

            private int ColumnIndex { get; set; }

            private Func<int, WordImageBodyFrame>? BodyFrameProvider { get; }

            private IReadOnlyDictionary<OpenXmlElement, WordImageSourceBlock>? SourceBlocks { get; }

            private List<WordDocumentVisualFragment>? Fragments { get; }

            internal string OverflowDiagnosticCode { get; }

            internal string OverflowDiagnosticMessage { get; }

            internal WordParagraphSpacingState? PreviousParagraphSpacingState { get; private set; }

            private List<WordTextExclusion>? TextExclusions { get; set; }

            internal bool TryGetSourceBlock(OpenXmlElement element, out WordImageSourceBlock sourceBlock) {
                if (SourceBlocks != null && SourceBlocks.TryGetValue(element, out sourceBlock)) {
                    return true;
                }

                sourceBlock = default;
                return false;
            }

            internal void AddFragment(WordDocumentVisualFragment fragment) {
                Fragments?.Add(fragment);
            }

            internal void AdvancePage() {
                PageIndex++;
                ApplyBodyFrame(PageIndex);
                ApplyColumn(0);
                StoppedForPagination = false;
                ClearParagraphSpacingState();
                TextExclusions = null;
            }

            internal void AdvancePages(int count) {
                for (int i = 0; i < count; i++) {
                    AdvancePage();
                }
            }

            internal void AdvanceColumnOrPage() {
                if (Columns.Count > 1 && ColumnIndex < Columns.Count - 1) {
                    ApplyColumn(ColumnIndex + 1);
                    StoppedForPagination = false;
                    ClearParagraphSpacingState();
                    TextExclusions = null;
                    return;
                }

                AdvancePage();
            }

            private void ApplyColumn(int columnIndex) {
                ColumnIndex = Columns.Count > 1 ? Math.Min(Math.Max(0, columnIndex), Columns.Count - 1) : 0;
                if (Columns.Count > 1) {
                    WordImageColumnFrame column = Columns[ColumnIndex];
                    Left = column.Left;
                    ContentWidth = column.Width;
                }

                Y = Top;
            }

            private void ApplyBodyFrame(int pageIndex) {
                if (BodyFrameProvider == null) {
                    return;
                }

                WordImageBodyFrame frame = BodyFrameProvider(Math.Max(0, pageIndex));
                Top = Math.Min(Drawing.Height, Math.Max(0D, frame.Top));
                ContentBottom = Math.Min(Drawing.Height, Math.Max(Top, frame.Bottom));
            }

            internal void AddTextExclusion(double left, double top, double right, double bottom) =>
                AddTextExclusion(left, top, right, bottom, WordTextWrapSide.Largest);

            internal void AddTextExclusion(double left, double top, double right, double bottom, WordTextWrapSide wrapSide) {
                if (right <= left || bottom <= top) {
                    return;
                }

                TextExclusions ??= new List<WordTextExclusion>();
                TextExclusions.Add(new WordTextExclusion(left, top, right, bottom, wrapSide));
            }

            internal void AddTextExclusion(IReadOnlyList<OfficePoint> polygon, WordTextWrapSide wrapSide) {
                if (polygon == null || polygon.Count < 3) {
                    return;
                }

                double left = double.MaxValue;
                double top = double.MaxValue;
                double right = double.MinValue;
                double bottom = double.MinValue;
                for (int index = 0; index < polygon.Count; index++) {
                    OfficePoint point = polygon[index];
                    left = Math.Min(left, point.X);
                    top = Math.Min(top, point.Y);
                    right = Math.Max(right, point.X);
                    bottom = Math.Max(bottom, point.Y);
                }

                if (right <= left || bottom <= top) {
                    return;
                }

                TextExclusions ??= new List<WordTextExclusion>();
                TextExclusions.Add(new WordTextExclusion(left, top, right, bottom, wrapSide, polygon));
            }

            internal WordTextFlowFrame ResolveTextFlowFrame() {
                double contentRight = Left + ContentWidth;
                if (TextExclusions == null || TextExclusions.Count == 0) {
                    return new WordTextFlowFrame(Left, ContentWidth);
                }

                for (int index = TextExclusions.Count - 1; index >= 0; index--) {
                    if (Y >= TextExclusions[index].Bottom) {
                        TextExclusions.RemoveAt(index);
                    }
                }

                if (TextExclusions.Count == 0) {
                    return new WordTextFlowFrame(Left, ContentWidth);
                }

                var activeExclusions = new List<WordTextExclusion>();
                for (int index = 0; index < TextExclusions.Count; index++) {
                    WordTextExclusion exclusion = TextExclusions[index];
                    if (Y >= exclusion.Top && Y < exclusion.Bottom) {
                        activeExclusions.Add(exclusion);
                    }
                }

                if (activeExclusions.Count == 0) {
                    return new WordTextFlowFrame(Left, ContentWidth);
                }

                List<WordTextFlowInterval> intervals = new List<WordTextFlowInterval> {
                    new WordTextFlowInterval(Left, contentRight)
                };
                double nextAvailableY = double.MaxValue;
                foreach (WordTextExclusion exclusion in activeExclusions) {
                    nextAvailableY = Math.Min(nextAvailableY, exclusion.Bottom);
                    double exclusionLeft = exclusion.Left;
                    double exclusionRight = exclusion.Right;
                    if (exclusion.Polygon != null) {
                        if (!TryGetPolygonHorizontalSpan(exclusion.Polygon, Y, out exclusionLeft, out exclusionRight)) {
                            continue;
                        }
                    }

                    if (exclusion.WrapSide == WordTextWrapSide.Left) {
                        intervals = IntersectTextFlowIntervals(intervals, Left, Math.Min(contentRight, exclusionLeft));
                    } else if (exclusion.WrapSide == WordTextWrapSide.Right) {
                        intervals = IntersectTextFlowIntervals(intervals, Math.Max(Left, exclusionRight), contentRight);
                    } else {
                        intervals = SubtractTextExclusionInterval(
                            intervals,
                            Math.Max(Left, exclusionLeft),
                            Math.Min(contentRight, exclusionRight));
                    }
                }

                if (TrySelectTextFlowInterval(intervals, out WordTextFlowInterval selected)) {
                    return new WordTextFlowFrame(selected.Left, selected.Width);
                }

                if (nextAvailableY < double.MaxValue) {
                    Y = nextAvailableY + ParagraphGapPoints;
                    return new WordTextFlowFrame(Left, ContentWidth);
                }

                return new WordTextFlowFrame(Left, ContentWidth);
            }

            private static List<WordTextFlowInterval> IntersectTextFlowIntervals(
                IReadOnlyList<WordTextFlowInterval> intervals,
                double left,
                double right) {
                var result = new List<WordTextFlowInterval>();
                if (right <= left) {
                    return result;
                }

                for (int index = 0; index < intervals.Count; index++) {
                    WordTextFlowInterval interval = intervals[index];
                    double segmentLeft = Math.Max(interval.Left, left);
                    double segmentRight = Math.Min(interval.Right, right);
                    if (segmentRight - segmentLeft >= 1D) {
                        result.Add(new WordTextFlowInterval(segmentLeft, segmentRight));
                    }
                }

                return result;
            }

            private static List<WordTextFlowInterval> SubtractTextExclusionInterval(
                IReadOnlyList<WordTextFlowInterval> intervals,
                double left,
                double right) {
                var result = new List<WordTextFlowInterval>();
                if (right <= left) {
                    result.AddRange(intervals);
                    return result;
                }

                for (int index = 0; index < intervals.Count; index++) {
                    WordTextFlowInterval interval = intervals[index];
                    if (right <= interval.Left || left >= interval.Right) {
                        result.Add(interval);
                        continue;
                    }

                    double leftSegmentRight = Math.Min(left, interval.Right);
                    if (leftSegmentRight - interval.Left >= 1D) {
                        result.Add(new WordTextFlowInterval(interval.Left, leftSegmentRight));
                    }

                    double rightSegmentLeft = Math.Max(right, interval.Left);
                    if (interval.Right - rightSegmentLeft >= 1D) {
                        result.Add(new WordTextFlowInterval(rightSegmentLeft, interval.Right));
                    }
                }

                return result;
            }

            private static bool TrySelectTextFlowInterval(IReadOnlyList<WordTextFlowInterval> intervals, out WordTextFlowInterval selected) {
                selected = default;
                double selectedWidth = 0D;
                bool hasSelection = false;
                for (int index = 0; index < intervals.Count; index++) {
                    WordTextFlowInterval interval = intervals[index];
                    double width = interval.Width;
                    if (width >= 1D && (!hasSelection || width >= selectedWidth)) {
                        selected = interval;
                        selectedWidth = width;
                        hasSelection = true;
                    }
                }

                return hasSelection;
            }

            private static bool TryGetPolygonHorizontalSpan(IReadOnlyList<OfficePoint> polygon, double y, out double left, out double right) {
                left = 0D;
                right = 0D;
                var intersections = new List<double>();
                for (int index = 0; index < polygon.Count; index++) {
                    OfficePoint start = polygon[index];
                    OfficePoint end = polygon[(index + 1) % polygon.Count];
                    if (Math.Abs(start.Y - end.Y) < 0.000001D) {
                        continue;
                    }

                    double minY = Math.Min(start.Y, end.Y);
                    double maxY = Math.Max(start.Y, end.Y);
                    if (y < minY || y >= maxY) {
                        continue;
                    }

                    double ratio = (y - start.Y) / (end.Y - start.Y);
                    intersections.Add(start.X + ((end.X - start.X) * ratio));
                }

                if (intersections.Count < 2) {
                    return false;
                }

                intersections.Sort();
                left = intersections[0];
                right = intersections[intersections.Count - 1];
                return right - left >= 1D;
            }

            internal void SetParagraphSpacingState(WordParagraphSpacingState state) =>
                PreviousParagraphSpacingState = state;

            internal void ClearParagraphSpacingState() =>
                PreviousParagraphSpacingState = null;
        }

        private readonly struct WordTextExclusion {
            internal WordTextExclusion(double left, double top, double right, double bottom, WordTextWrapSide wrapSide, IReadOnlyList<OfficePoint>? polygon = null) {
                Left = left;
                Top = top;
                Right = right;
                Bottom = bottom;
                WrapSide = wrapSide;
                Polygon = polygon;
            }

            internal double Left { get; }

            internal double Top { get; }

            internal double Right { get; }

            internal double Bottom { get; }

            internal WordTextWrapSide WrapSide { get; }

            internal IReadOnlyList<OfficePoint>? Polygon { get; }
        }

        private readonly struct WordTextFlowInterval {
            internal WordTextFlowInterval(double left, double right) {
                Left = left;
                Right = right;
            }

            internal double Left { get; }

            internal double Right { get; }

            internal double Width => Math.Max(0D, Right - Left);
        }

        private enum WordTextWrapSide {
            Largest,
            Left,
            Right
        }

        private readonly struct WordTextFlowFrame {
            internal WordTextFlowFrame(double left, double width) {
                Left = left;
                Width = width;
            }

            internal double Left { get; }

            internal double Width { get; }
        }
    }
}
