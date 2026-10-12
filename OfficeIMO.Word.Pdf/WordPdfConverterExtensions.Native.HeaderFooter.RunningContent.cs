using DocumentFormat.OpenXml;
using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    private static bool RequiresNativeRunningStory(WordHeaderFooter? story) => story != null &&
        (story.ChildElements.Any(element => element is W.Table || element.Descendants<W.Table>().Any()) ||
        story.ChildElements.SelectMany(element => element is W.Paragraph paragraph
            ? new[] { paragraph }.Concat(element.Descendants<W.Paragraph>()) : element.Descendants<W.Paragraph>())
            .Any(paragraph => GetNativeHeaderFooterVisibleTextRuns(new WordParagraph(story.Document, paragraph))
                .Any(item => item.IsField && item.Text == "{sectionpages}")));

    private static bool UsesNativeRunningHeader(WordSection section, WordToPdfOptions? options) => UsesNativeRunningStories(section, options, false,
        section.Header?.Default, section.DifferentFirstPage ? section.Header?.First : null,
        section.DocumentOddEvenSettingEnabled ? section.Header?.Even : null);

    private static bool UsesNativeRunningFooter(WordSection section, WordToPdfOptions? options) => UsesNativeRunningStories(section, options, true,
        section.Footer?.Default, section.DifferentFirstPage ? section.Footer?.First : null,
        section.DocumentOddEvenSettingEnabled ? section.Footer?.Even : null);

    private static bool UsesNativeRunningStories(WordSection section, WordToPdfOptions? options, bool isFooter, params WordHeaderFooter?[] variants) {
        if (variants.Any(RequiresNativeRunningStory)) return true;
        bool hasNumberedParagraph = false;
        bool hasParagraph = false;
        bool ordinaryStoriesSupported = !HasNativeRunningStoryOverrides(options?.PdfOptions, isFooter);
        PdfCore.PageMargins margins = GetNativeMargins(section, options, (0D, 0D));
        double width = GetNativePageSize(section, options).Width - margins.Left - margins.Right;
        foreach (WordHeaderFooter? story in variants) {
            if (story == null) continue;
            double minimumHeight = 0D;
            double joinedLineHeight = double.PositiveInfinity;
            ordinaryStoriesSupported &= story.ChildElements.All(element => element is W.Paragraph);
            // Admission switches all active variants. Preserve the source-order
            // text-box path until flow can retain its mixed and nested content.
            if (story.ChildElements.Any(element => element.Descendants<W.TextBoxContent>().Any())) return false;
            foreach (W.Paragraph source in story.ChildElements.SelectMany(element => element is W.Paragraph paragraph
                ? new[] { paragraph }.Concat(element.Descendants<W.Paragraph>()) : element.Descendants<W.Paragraph>())) {
                var paragraph = new WordParagraph(story.Document, source);
                hasParagraph = true;
                // Drawings, text boxes and explicit pagination retain their
                // established bounded export until running flow preserves them.
                ordinaryStoriesSupported &= !source.Descendants<W.Drawing>().Any() &&
                    !source.Descendants<W.Picture>().Any() && !source.Descendants<W.EmbeddedObject>().Any() &&
                    !source.Descendants<AlternateContent>().Any() && !HasNativePageBreakBefore(paragraph) &&
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
                    (style.FontSize ?? NativeHeaderFooterFontSize) * (style.LineHeight ?? 1.2D));
                if (ShouldRenderNativeEmptyParagraphLineBox(paragraph)) {
                    minimumHeight += joinedLineHeight;
                    joinedLineHeight = double.PositiveInfinity;
                }
                if (info.HasValue) ApplyNativeInlineListIndent(paragraph, style);
                // Keep the established bounded simple-story export for indents
                // that cannot leave both paragraph-flow text frames positive.
                // Admission switches every paragraph in every active variant.
                double textWidth = width - style.LeftIndent - style.RightIndent;
                double firstLineWidth = textWidth - style.FirstLineIndent;
                if (textWidth <= 0D || firstLineWidth <= 0D ||
                    double.IsNaN(textWidth) || double.IsInfinity(textWidth) ||
                    double.IsNaN(firstLineWidth) || double.IsInfinity(firstLineWidth)) return false;
            }
            // Retain the bounded simple-story export for obviously oversized
            // ordinary stories. Its expansion clamp preserves a usable body;
            // bounded running flow cannot create continuation pages for them.
            double distance = (isFooter ? section.Margins.FooterDistance : section.Margins.HeaderDistance) / 20D;
            ordinaryStoriesSupported &= minimumHeight + distance < GetNativePageSize(section, options).Height;
        }
        return hasNumberedParagraph || (hasParagraph && ordinaryStoriesSupported);
    }

    /// <summary>
    /// Keeps caller-configured simple-story positioning and typography on the
    /// path that implements those overrides. Registered body font resources do
    /// not prevent source paragraphs from using running flow.
    /// </summary>
    private static bool HasNativeRunningStoryOverrides(PdfCore.PdfOptions? options, bool isFooter) => options != null &&
        (isFooter
            ? options.FooterOffsetY != 18D || options.HasExplicitFooterFont || options.FooterFontFamily != null ||
                options.FooterFontSize != 9D || options.FooterTextColor.HasValue || options.FooterAlign != PdfCore.PdfAlign.Center
            : options.HeaderOffsetY != 18D || options.HasExplicitHeaderFont || options.HeaderFontFamily != null ||
                options.HeaderFontSize != 9D || options.HeaderTextColor.HasValue || options.HeaderAlign != PdfCore.PdfAlign.Center);

    private static void ConfigureNativeRunningHeaderFooter(PdfCore.PdfPageBuilder page, WordSection section,
        WordToPdfOptions? options, NativeFontMap fontMap,
        IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers, NativeDocumentDefaults defaults) {
        if (UsesNativeRunningHeader(section, options)) {
            double distance = section.Margins.HeaderDistance / 20D;
            page.Header(header => {
                header.Content(CreateNativeRunningStory(section.Header?.Default, section, options, fontMap, listMarkers, defaults, false), distance);
                if (section.DifferentFirstPage)
                    header.FirstPageContent(CreateNativeRunningStory(section.Header?.First, section, options, fontMap, listMarkers, defaults, false), distance);
                if (section.DocumentOddEvenSettingEnabled)
                    header.EvenPagesContent(CreateNativeRunningStory(section.Header?.Even, section, options, fontMap, listMarkers, defaults, false), distance);
            });
            if (!new[] { section.Header?.Default, section.DifferentFirstPage ? section.Header?.First : null,
                    section.DocumentOddEvenSettingEnabled ? section.Header?.Even : null }.Any(RequiresNativeRunningStory))
                page.Options.AllowRunningContentTextFallback(isHeader: true);
        }
        if (UsesNativeRunningFooter(section, options)) {
            double distance = section.Margins.FooterDistance / 20D;
            page.Footer(footer => {
                footer.Content(CreateNativeRunningStory(section.Footer?.Default, section, options, fontMap, listMarkers, defaults, true), distance);
                if (section.DifferentFirstPage)
                    footer.FirstPageContent(CreateNativeRunningStory(section.Footer?.First, section, options, fontMap, listMarkers, defaults, true), distance);
                if (section.DocumentOddEvenSettingEnabled)
                    footer.EvenPagesContent(CreateNativeRunningStory(section.Footer?.Even, section, options, fontMap, listMarkers, defaults, true), distance);
            });
            if (!new[] { section.Footer?.Default, section.DifferentFirstPage ? section.Footer?.First : null,
                    section.DocumentOddEvenSettingEnabled ? section.Footer?.Even : null }.Any(RequiresNativeRunningStory))
                page.Options.AllowRunningContentTextFallback(isHeader: false);
        }
    }

    private static Func<PdfCore.PdfRunningContentContext, Action<PdfCore.PdfContentBuilder>> CreateNativeRunningStory(
        WordHeaderFooter? story, WordSection section, WordToPdfOptions? options, NativeFontMap fontMap,
        IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers, NativeDocumentDefaults defaults, bool isFooter) {
        IReadOnlyList<WordElement> elements = story == null ? Array.Empty<WordElement>() : CollapseNativeParagraphElements(story.Elements);
        RecordNativeRunningStoryJoinDiagnostics(elements, listMarkers, defaults, fontMap, options);
        // Running callbacks are evaluated while the PDF is written, after the
        // operation report is captured. Discover approximation diagnostics now.
        if (story != null)
            foreach (W.Paragraph source in story.ChildElements.SelectMany(element => element is W.Paragraph paragraph
                ? new[] { paragraph }.Concat(element.Descendants<W.Paragraph>()) : element.Descendants<W.Paragraph>()))
                RecordNativeMixedTextBoxDiagnostic(GetNativeRuns(new WordParagraph(story.Document, source)), options,
                    "header/footer paragraph");
        bool addPageNumber = isFooter && options?.IncludePageNumbers == true &&
            (story == null || GetNativeHeaderFooterText(story, listMarkers, fontMap)?.HasPageTokens != true);
        return context => content => {
            options?.CancellationToken.ThrowIfCancellationRequested();
            var flow = new NativeSpacingCollapseFlow(new NativePdfColumnFlow(content,
                new PdfCore.PageSize(context.PageWidth, context.PageHeight)), preserveInitialSpacingBefore: true);
            NativeDocumentDefaults runningDefaults = defaults with { RunningContentContext = context };
            var notes = new NativeNoteNumbering(section._document, options);
            notes.BeginSection(section);
            var destinations = new Dictionary<W.Paragraph, string>();
            for (int index = 0; index < elements.Count; index++) {
                options?.CancellationToken.ThrowIfCancellationRequested();
                if (TryRenderNativeJoinedParagraphs(flow, elements, ref index,
                    paragraph => listMarkers.TryGetValue(paragraph, out var joinedMarker) ? joinedMarker : null,
                    notes, options, runningDefaults, fontMap)) continue;
                RenderNativeElement(flow, elements[index], section,
                    paragraph => listMarkers.TryGetValue(paragraph, out var marker) ? marker : null,
                    Array.Empty<int>(), notes, options, Array.Empty<NativeTableOfContentsEntry>(), destinations,
                    context.ContentWidth, runningDefaults, fontMap,
                    index + 1 < elements.Count ? elements[index + 1] : null);
            }
            if (addPageNumber)
                content.Paragraph(paragraph => paragraph.Text(FormatNativeRunningPageTokens(
                    GetNativePageNumberFormat(options), context, null)), PdfCore.PdfAlign.Right);
        };
    }

    private static void RecordNativeRunningStoryJoinDiagnostics(IReadOnlyList<WordElement> elements,
        IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers, NativeDocumentDefaults defaults,
        NativeFontMap fontMap, WordToPdfOptions? options) {
        if (options == null) return;
        for (int index = 0; index < elements.Count; index++) {
            if (elements[index] is WordTable table) {
                RecordNativeRunningTableJoinDiagnostics(table, listMarkers, defaults, fontMap, options);
                continue;
            }
            if (!TryGetNativeParagraphJoinChain(elements, index, out var paragraphs, out int lastIndex)) continue;
            if (!CanJoinNativeParagraphChain(paragraphs,
                paragraph => listMarkers.TryGetValue(paragraph, out var marker) ? marker : null, defaults, fontMap))
                AddNativeExportWarning(options, "NativeHiddenParagraphJoinUnsupported", "header/footer",
                    "Hidden paragraph marks beside lists, headings, decorated paragraphs, flow breaks or objects retain separate PDF paragraph boundaries.");
            index = lastIndex;
        }
    }

    private static void RecordNativeRunningTableJoinDiagnostics(WordTable table,
        IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers, NativeDocumentDefaults defaults,
        NativeFontMap fontMap, WordToPdfOptions options, int depth = 0) {
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
                    for (int index = 0; index < elements.Count; index++) {
                        if (elements[index] is WordTable nested) {
                            RecordNativeRunningTableJoinDiagnostics(nested, listMarkers, defaults, fontMap, options, depth + 1);
                        } else if (elements[index] is WordParagraph) {
                            // Use the renderer's eligibility, including effective
                            // conditional styles and caller-configured typography.
                            index = GetNativeJoinedCellParagraphEnd(elements, index,
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
