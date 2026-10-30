using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Threading;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private const double NativeDefaultParagraphLineHeight = 1.15D;
        private const double NativeDefaultParagraphSpacingAfter = 0D;
        private const double NativeCalibriSingleLineHeight = 1.220703125D;
        private const double NativeTablePageContinuationSpacingBefore = 24D;
        private const double NativeHeaderFooterFontSize = 9D;
        private const double NativeHeaderFooterLineHeight = NativeHeaderFooterFontSize * 1.2D;
        private const double NativeHeaderFooterBodyGap = 2D;
        private const double NativeMinimumBodyHeight = 72D;
        private const double NativeHeaderFooterDefaultOffset = 18D;
        private const double NativeFooterDefaultOffset = 20D;

        // OOXML auto line spacing is expressed in 240ths of Word's single-line
        // box. Resolve that box from the effective font instead of applying the
        // Calibri/Carlito metric to Arial, Times, or arbitrary embedded faces.
        private static double ResolveNativeWordSingleLineHeight(params string?[] familyCandidates) =>
            ResolveNativeWordSingleLineHeight(null, familyCandidates);

        private static double ResolveNativeWordSingleLineHeight(
            NativeFontMap? nativeFontMap,
            params string?[] familyCandidates) {
            foreach (string? candidate in familyCandidates) {
                if (string.IsNullOrWhiteSpace(candidate)) {
                    continue;
                }

                if (nativeFontMap?.TryResolveLineSpacingRatio(candidate, out double resolvedRatio) == true) {
                    return resolvedRatio;
                }

                string normalized = NormalizeNativeFontFamily(candidate!);
                if (normalized.StartsWith("calibri", StringComparison.OrdinalIgnoreCase) ||
                    normalized.StartsWith("carlito", StringComparison.OrdinalIgnoreCase)) {
                    return NativeCalibriSingleLineHeight;
                }
            }

            if (nativeFontMap?.TryResolveDefaultLineSpacingRatio(out double defaultRatio) == true) {
                return defaultRatio;
            }

            return NativeDefaultParagraphLineHeight;
        }

        private interface INativePdfFlow {
            void ContinuedFlow(Action<PdfCore.PdfContentBuilder> build, Action<PdfCore.PdfContentBuilder>? separator,
                Action<PdfCore.PdfContentBuilder>? continuationSeparator, Action<PdfCore.PdfContentBuilder>? continuationNotice,
                bool reserveFinalNotice, bool allowContinuationSeparatorInBottomMargin);
            PdfCore.PdfFootnoteDefinition DefineFootnote(Action<PdfCore.PdfContentBuilder> build, PdfCore.PdfFootnoteOptions options);
            bool SupportsPositionedTables { get; }
            PdfCore.PdfPageSize PageSize { get; }
            void PageBreak(bool preserveEmptyPage = false);
            void ColumnBreak();
            void Spacer(double height);
            void ParagraphSpacingBefore(double height);
            void ParagraphSpacingAfter(double height);
            void Bookmark(string name);
            void TableOfContents(PdfCore.PdfTableOfContentsOptions options);
            void HR(double? thickness = null, global::OfficeIMO.OfficeColor? color = null, double? spacingBefore = null, double? spacingAfter = null, PdfCore.PdfHorizontalRuleStyle? style = null);
            void Paragraph(Action<PdfCore.PdfParagraphBuilder> build, PdfCore.PdfAlign align = PdfCore.PdfAlign.Left, global::OfficeIMO.OfficeColor? defaultColor = null, PdfCore.PdfParagraphStyle? style = null);
            void Panel(Action<PdfCore.PdfContentBuilder> build, PdfCore.PdfPanelStyle? style = null);
            void PanelParagraph(Action<PdfCore.PdfParagraphBuilder> build, PdfCore.PdfPanelStyle? style = null, PdfCore.PdfAlign align = PdfCore.PdfAlign.Left, global::OfficeIMO.OfficeColor? defaultColor = null, PdfCore.PdfParagraphStyle? paragraphStyle = null);
            void Heading(int level, string text, Action<PdfCore.PdfParagraphBuilder> build, PdfCore.PdfAlign align, global::OfficeIMO.OfficeColor? color, PdfCore.PdfHeadingStyle? style);
            void RichNumbered(IEnumerable<PdfCore.PdfListItem> items, PdfCore.PdfAlign align, global::OfficeIMO.OfficeColor? color, int startNumber, PdfCore.PdfListStyle? style);
            void RichBullets(IEnumerable<PdfCore.PdfListItem> items, PdfCore.PdfAlign align, global::OfficeIMO.OfficeColor? color, PdfCore.PdfListStyle? style);
            void TextField(string name, double width, double height, string value, PdfCore.PdfAlign align, double fontSize, double spacingBefore, double spacingAfter, PdfCore.PdfFormFieldStyle? style = null);
            void ChoiceField(string name, IEnumerable<string> options, string? value, double width, double height, PdfCore.PdfAlign align, double fontSize, double spacingBefore, double spacingAfter, bool isComboBox, PdfCore.PdfFormFieldStyle? style = null);
            void CheckBox(string name, bool isChecked, double size, PdfCore.PdfAlign align, double spacingBefore, double spacingAfter, string checkedValueName = "Yes", PdfCore.PdfFormFieldStyle? style = null);
            void Shape(OfficeShape shape, PdfCore.PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfCore.PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null);
            void Drawing(OfficeDrawing drawing, PdfCore.PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfCore.PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null);
            void Canvas(Action<PdfCore.PdfPageCanvas> build);
            void Table(IEnumerable<PdfCore.PdfTableCell[]> rows, PdfCore.PdfAlign align, PdfCore.PdfTableStyle? style);
            void Image(byte[] bytes, double width, double height, PdfCore.PdfAlign? align = null, PdfCore.PdfImageStyle? style = null);
        }

        private sealed class NativePdfDocumentFlow : INativePdfFlow {
            public void ContinuedFlow(Action<PdfCore.PdfContentBuilder> build, Action<PdfCore.PdfContentBuilder>? separator,
                Action<PdfCore.PdfContentBuilder>? continuationSeparator, Action<PdfCore.PdfContentBuilder>? continuationNotice,
                bool reserveFinalNotice, bool allowContinuationSeparatorInBottomMargin) =>
                _pdf.ContinuedFlow(build, separator, continuationSeparator, continuationNotice, reserveFinalNotice,
                    allowContinuationSeparatorInBottomMargin);
            public PdfCore.PdfFootnoteDefinition DefineFootnote(Action<PdfCore.PdfContentBuilder> build, PdfCore.PdfFootnoteOptions options) => _pdf.DefineFootnote(build, options);
            public bool SupportsPositionedTables => true;
            private readonly PdfCore.PdfDocument _pdf;

            public NativePdfDocumentFlow(PdfCore.PdfDocument pdf, PdfCore.PdfPageSize pageSize) {
                _pdf = pdf;
                PageSize = pageSize;
            }

            public PdfCore.PdfPageSize PageSize { get; }

            public void PageBreak(bool preserveEmptyPage = false) => _pdf.PageBreak(preserveEmptyPage);
            public void ColumnBreak() => _pdf.PageBreak(preserveEmptyPage: true);
            public void Spacer(double height) => _pdf.Spacer(height);
            public void ParagraphSpacingBefore(double height) { if (height > 0D) _pdf.Spacer(height); }
            public void ParagraphSpacingAfter(double height) { if (height > 0D) _pdf.Spacer(height); }
            public void Bookmark(string name) => _pdf.Bookmark(name);
            public void TableOfContents(PdfCore.PdfTableOfContentsOptions options) => _pdf.TableOfContents(options);
            public void HR(double? thickness = null, global::OfficeIMO.OfficeColor? color = null, double? spacingBefore = null, double? spacingAfter = null, PdfCore.PdfHorizontalRuleStyle? style = null) => _pdf.HR(thickness, color, spacingBefore, spacingAfter, style);
            public void Paragraph(Action<PdfCore.PdfParagraphBuilder> build, PdfCore.PdfAlign align = PdfCore.PdfAlign.Left, global::OfficeIMO.OfficeColor? defaultColor = null, PdfCore.PdfParagraphStyle? style = null) => _pdf.Paragraph(build, align, defaultColor, style);
            public void Panel(Action<PdfCore.PdfContentBuilder> build, PdfCore.PdfPanelStyle? style = null) => _pdf.Panel(build, style);
            public void PanelParagraph(Action<PdfCore.PdfParagraphBuilder> build, PdfCore.PdfPanelStyle? style = null, PdfCore.PdfAlign align = PdfCore.PdfAlign.Left, global::OfficeIMO.OfficeColor? defaultColor = null, PdfCore.PdfParagraphStyle? paragraphStyle = null) =>
                _pdf.Panel(content => content.Paragraph(build, align, defaultColor, CreateNativePanelParagraphStyle(paragraphStyle, style)), style);
            public void Heading(int level, string text, Action<PdfCore.PdfParagraphBuilder> build, PdfCore.PdfAlign align, global::OfficeIMO.OfficeColor? color, PdfCore.PdfHeadingStyle? style) {
                _pdf.Heading(level, text, build, align, color, style);
            }
            public void RichNumbered(IEnumerable<PdfCore.PdfListItem> items, PdfCore.PdfAlign align, global::OfficeIMO.OfficeColor? color, int startNumber, PdfCore.PdfListStyle? style) => _pdf.RichNumbered(items, align, color, startNumber, style);
            public void RichBullets(IEnumerable<PdfCore.PdfListItem> items, PdfCore.PdfAlign align, global::OfficeIMO.OfficeColor? color, PdfCore.PdfListStyle? style) => _pdf.RichBullets(items, align, color, style);
            public void TextField(string name, double width, double height, string value, PdfCore.PdfAlign align, double fontSize, double spacingBefore, double spacingAfter, PdfCore.PdfFormFieldStyle? style) => _pdf.TextField(name, global::OfficeIMO.OfficeLength.Points(width), global::OfficeIMO.OfficeLength.Points(height), value, align, global::OfficeIMO.OfficeLength.Points(fontSize), global::OfficeIMO.OfficeLength.Points(spacingBefore), global::OfficeIMO.OfficeLength.Points(spacingAfter), style);
            public void ChoiceField(string name, IEnumerable<string> options, string? value, double width, double height, PdfCore.PdfAlign align, double fontSize, double spacingBefore, double spacingAfter, bool isComboBox, PdfCore.PdfFormFieldStyle? style) => _pdf.ChoiceField(name, options, value, width, height, align, fontSize, spacingBefore, spacingAfter, isComboBox, style);
            public void CheckBox(string name, bool isChecked, double size, PdfCore.PdfAlign align, double spacingBefore, double spacingAfter, string checkedValueName, PdfCore.PdfFormFieldStyle? style) => _pdf.CheckBox(name, isChecked, size, align, spacingBefore, spacingAfter, checkedValueName, style);
            public void Shape(OfficeShape shape, PdfCore.PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfCore.PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) => _pdf.Shape(shape, align, spacingBefore, spacingAfter, style, linkUri, linkContents);
            public void Drawing(OfficeDrawing drawing, PdfCore.PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfCore.PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) => _pdf.Drawing(drawing, align, spacingBefore, spacingAfter, style, linkUri, linkContents);
            public void Canvas(Action<PdfCore.PdfPageCanvas> build) => _pdf.Canvas(build);
            public void Table(IEnumerable<PdfCore.PdfTableCell[]> rows, PdfCore.PdfAlign align, PdfCore.PdfTableStyle? style) => _pdf.Table(rows, align, style);
            public void Image(byte[] bytes, double width, double height, PdfCore.PdfAlign? align = null, PdfCore.PdfImageStyle? style = null) => _pdf.Image(bytes, width, height, align, style: style ?? CreateNativeImageStyle());
        }

        private sealed class NativePdfColumnFlow : INativePdfFlow {
            public void ContinuedFlow(Action<PdfCore.PdfContentBuilder> build, Action<PdfCore.PdfContentBuilder>? separator,
                Action<PdfCore.PdfContentBuilder>? continuationSeparator, Action<PdfCore.PdfContentBuilder>? continuationNotice,
                bool reserveFinalNotice, bool allowContinuationSeparatorInBottomMargin) =>
                _column.ContinuedFlow(build, separator, continuationSeparator, continuationNotice, reserveFinalNotice,
                    allowContinuationSeparatorInBottomMargin);
            public PdfCore.PdfFootnoteDefinition DefineFootnote(Action<PdfCore.PdfContentBuilder> build, PdfCore.PdfFootnoteOptions options) => _column.DefineFootnote(build, options);
            public bool SupportsPositionedTables => false;
            private readonly PdfCore.PdfPageBuilder? _page;
            private readonly PdfCore.PdfContentBuilder _column;

            public NativePdfColumnFlow(PdfCore.PdfPageBuilder page, PdfCore.PdfContentBuilder column, PdfCore.PdfPageSize pageSize) {
                _page = page;
                _column = column;
                PageSize = pageSize;
            }

            internal NativePdfColumnFlow(PdfCore.PdfContentBuilder content, PdfCore.PdfPageSize pageSize) {
                _column = content;
                PageSize = pageSize;
            }

            public PdfCore.PdfPageSize PageSize { get; }

            public void PageBreak(bool preserveEmptyPage = false) => _column.PageBreak(preserveEmptyPage);
            public void ColumnBreak() => _column.ColumnBreak();
            public void Spacer(double height) => _column.Spacer(height);
            public void ParagraphSpacingBefore(double height) { if (height > 0D) _column.Spacer(height); }
            public void ParagraphSpacingAfter(double height) { if (height > 0D) _column.Spacer(height); }
            public void Bookmark(string name) => _column.Bookmark(name);
            public void TableOfContents(PdfCore.PdfTableOfContentsOptions options) => _column.TableOfContents(options);
            public void HR(double? thickness = null, global::OfficeIMO.OfficeColor? color = null, double? spacingBefore = null, double? spacingAfter = null, PdfCore.PdfHorizontalRuleStyle? style = null) => _column.HR(thickness, color, spacingBefore, spacingAfter, style);
            public void Paragraph(Action<PdfCore.PdfParagraphBuilder> build, PdfCore.PdfAlign align = PdfCore.PdfAlign.Left, global::OfficeIMO.OfficeColor? defaultColor = null, PdfCore.PdfParagraphStyle? style = null) => _column.Paragraph(build, align, defaultColor, style);
            public void Panel(Action<PdfCore.PdfContentBuilder> build, PdfCore.PdfPanelStyle? style = null) => _column.Panel(build, style);
            public void PanelParagraph(Action<PdfCore.PdfParagraphBuilder> build, PdfCore.PdfPanelStyle? style = null, PdfCore.PdfAlign align = PdfCore.PdfAlign.Left, global::OfficeIMO.OfficeColor? defaultColor = null, PdfCore.PdfParagraphStyle? paragraphStyle = null) =>
                _column.Panel(content => content.Paragraph(build, align, defaultColor, CreateNativePanelParagraphStyle(paragraphStyle, style)), style);
            public void Heading(int level, string text, Action<PdfCore.PdfParagraphBuilder> build, PdfCore.PdfAlign align, global::OfficeIMO.OfficeColor? color, PdfCore.PdfHeadingStyle? style) {
                _column.Heading(level, text, build, align, color, style);
            }
            public void RichNumbered(IEnumerable<PdfCore.PdfListItem> items, PdfCore.PdfAlign align, global::OfficeIMO.OfficeColor? color, int startNumber, PdfCore.PdfListStyle? style) => _column.RichNumbered(items, align, color, startNumber, style);
            public void RichBullets(IEnumerable<PdfCore.PdfListItem> items, PdfCore.PdfAlign align, global::OfficeIMO.OfficeColor? color, PdfCore.PdfListStyle? style) => _column.RichBullets(items, align, color, style);
            public void TextField(string name, double width, double height, string value, PdfCore.PdfAlign align, double fontSize, double spacingBefore, double spacingAfter, PdfCore.PdfFormFieldStyle? style) => _column.TextField(name, global::OfficeIMO.OfficeLength.Points(width), global::OfficeIMO.OfficeLength.Points(height), value, align, global::OfficeIMO.OfficeLength.Points(fontSize), global::OfficeIMO.OfficeLength.Points(spacingBefore), global::OfficeIMO.OfficeLength.Points(spacingAfter), style);
            public void ChoiceField(string name, IEnumerable<string> options, string? value, double width, double height, PdfCore.PdfAlign align, double fontSize, double spacingBefore, double spacingAfter, bool isComboBox, PdfCore.PdfFormFieldStyle? style) => _column.ChoiceField(name, options, value, width, height, align, fontSize, spacingBefore, spacingAfter, isComboBox, style);
            public void CheckBox(string name, bool isChecked, double size, PdfCore.PdfAlign align, double spacingBefore, double spacingAfter, string checkedValueName, PdfCore.PdfFormFieldStyle? style) => _column.CheckBox(name, isChecked, size, align, spacingBefore, spacingAfter, checkedValueName, style);
            public void Shape(OfficeShape shape, PdfCore.PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfCore.PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) => _column.Shape(shape, align, spacingBefore, spacingAfter, style, linkUri, linkContents);
            public void Drawing(OfficeDrawing drawing, PdfCore.PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfCore.PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) => _column.Drawing(drawing, align, spacingBefore, spacingAfter, style, linkUri, linkContents);
            public void Canvas(Action<PdfCore.PdfPageCanvas> build) {
                if (_page == null) _column.Canvas(build);
                else _page.Canvas(build);
            }
            public void Table(IEnumerable<PdfCore.PdfTableCell[]> rows, PdfCore.PdfAlign align, PdfCore.PdfTableStyle? style) => _column.Table(rows, align, style);
            public void Image(byte[] bytes, double width, double height, PdfCore.PdfAlign? align = null, PdfCore.PdfImageStyle? style = null) => _column.Image(bytes, width, height, align, style: style ?? CreateNativeImageStyle());
        }

        /// <summary>Preserves text layout inside decorations while the surrounding panel owns paragraph spacing.</summary>
        private static PdfCore.PdfParagraphStyle? CreateNativePanelParagraphStyle(PdfCore.PdfParagraphStyle? style, PdfCore.PdfPanelStyle? panelStyle) {
            if (style == null) return null;
            PdfCore.PdfParagraphStyle result = style.Clone();
            result.SpacingBefore = 0D;
            result.SpacingAfter = panelStyle == null ? 0D : Math.Max(0D, (style.SpacingAfter ?? 0D) - panelStyle.SpacingAfter);
            return result;
        }

        private static PdfCore.PdfDocument CreateOfficeIMOPdfDocument(WordDocument document, WordPdfExportOptions? options) {
            CancellationToken cancellationToken = options?.CancellationToken ?? CancellationToken.None;
            cancellationToken.ThrowIfCancellationRequested();
            ResetNativeStyleLookupCache(document);
            WordBuiltinDocumentProperties properties = document.BuiltinDocumentProperties;
            var nativeFontMap = new NativeFontMap(options?.Report);
            Dictionary<WordParagraph, WordDocumentTraversal.ResolvedListMarker> resolvedMarkers =
                WordDocumentTraversal.BuildResolvedListMarkers(document, cancellationToken);
            using IDisposable listInfoScope = WordDocumentTraversal.UseResolvedListMarkers(document, resolvedMarkers);
            Dictionary<WordParagraph, (int Level, string Marker)> listMarkers = resolvedMarkers.ToDictionary(
                pair => pair.Key, pair => (pair.Value.Level, pair.Value.Marker), resolvedMarkers.Comparer);
            if (options != null) {
                IEnumerable<WordDocumentTraversal.ResolvedListMarker> renderedMarkers = resolvedMarkers
                    .Where(pair => IsNativeRenderedListStoryParagraph(document, pair.Key))
                    .Select(pair => pair.Value);
                foreach (int pictureBulletId in WordDocumentTraversal.GetPictureBulletFallbackIds(renderedMarkers)) {
                    AddNativeExportWarning(options, "native-picture-bullet-text-fallback", "list marker",
                        "Picture bullet " + pictureBulletId.ToString(CultureInfo.InvariantCulture) + " is represented by a portable text bullet in PDF output.");
                }
            }
            PdfCore.PdfDocument pdf = PdfCore.PdfDocument.CreateForImportedLayout(CreateNativeOptions(
                    document,
                    options,
                    nativeFontMap,
                    listMarkers))
                .Meta(
                    title: options?.Title ?? properties.Title,
                    author: options?.Author ?? properties.Creator,
                    subject: options?.Subject ?? properties.Subject,
                    keywords: BuildNativeKeywords(options, properties));

            Dictionary<WordParagraph, (int Level, int Index)> listIndices = WordDocumentTraversal.BuildListIndices(document);
            Dictionary<W.Paragraph, string> headingDestinations = BuildNativeHeadingDestinations(document);
            IReadOnlyList<NativeTableOfContentsEntry> tableOfContentsEntries = BuildNativeTableOfContentsEntries(document, headingDestinations);
            NativeDocumentDefaults nativeDefaults = GetNativeDocumentDefaults(document, nativeFontMap);
            nativeDefaults = nativeDefaults with {
                BodyTextInkTopOffset = GetNativeBodyTextInkTopOffset(document, nativeDefaults, nativeFontMap, options),
                RunningTextBoxStoryPeers = GetNativeRunningTextBoxStoryPeers(document, options)
            };
            var footnoteNumbersById = new NativeNoteNumbering(document, options);
            nativeDefaults = nativeDefaults with { NoteNumbering = footnoteNumbersById };
            IReadOnlyList<WordSection> sections = document.Sections;
            bool canContinueRunningStories = SupportsNativeContinuousRunningStoryDocument(document, sections);
            string? pageFootnoteBodyProfileGap = null;
            bool pageFootnoteBodyProfileChecked = false;
            var documentEndnotes = new List<PdfFootnote>();
            for (int sectionIndex = 0; sectionIndex < sections.Count;) {
                cancellationToken.ThrowIfCancellationRequested();
                int sectionGroupEnd = GetNativePdfSectionGroupEnd(sections, sectionIndex, options, listMarkers, canContinueRunningStories);
                bool changesRunningStories = NativeSectionGroupChangesRunningStories(sections, sectionIndex, sectionGroupEnd);
                WordSection firstSection = sections[sectionIndex];
                if (sectionIndex > 0) ReportNativeContinuousStoryPageFallback(
                    sections[sectionIndex - 1], firstSection, sectionIndex, options);
                bool preserveLogicalRunningSections = changesRunningStories || UsesNativeSectionPageCounts(firstSection) || firstSection.DifferentFirstPage;
                PdfCore.PdfPageSize sectionPageSize = GetNativePageSize(firstSection, options);
                (double Header, double Footer) headerFooterMarginExpansion = GetNativeHeaderFooterMarginExpansion(firstSection, options, listMarkers);
                PdfCore.PdfPageMargins sectionMargins = GetNativeMargins(firstSection, options, headerFooterMarginExpansion);
                double sectionContentWidth = Math.Max(1D, sectionPageSize.Width.ToPoints() - sectionMargins.Left.ToPoints() - sectionMargins.Right.ToPoints());
                pdf.Section(page => {
                    page.Size(sectionPageSize);
                    page.Margin(sectionMargins);
                    page.MirrorMargins(ShouldMirrorNativeMargins(firstSection, options));
                    page.StartOnPageParity(GetNativeSectionStartParity(firstSection), useContinuingPageNumber: true);
                    ConfigureNativePageNumbering(page, firstSection);
                    ConfigureNativeHeaderFooter(page, firstSection, options, headerFooterMarginExpansion.Header, headerFooterMarginExpansion.Footer, nativeFontMap, listMarkers);
                    ConfigureNativeRunningHeaderFooter(page, firstSection, options, nativeFontMap, listMarkers, nativeDefaults);
                    INativePdfFlow flow = new NativeSpacingCollapseFlow(new NativePdfDocumentFlow(pdf, sectionPageSize));

                    for (int currentSectionIndex = sectionIndex; currentSectionIndex < sectionGroupEnd; currentSectionIndex++) {
                        cancellationToken.ThrowIfCancellationRequested();
                        WordSection section = sections[currentSectionIndex];
                        if (preserveLogicalRunningSections) {
                            if (changesRunningStories && currentSectionIndex > sectionIndex) {
                                PdfCore.PdfCreateOptions continuationOptions = CreateNativeContinuousRunningOptions(
                                    pdf, section, options, nativeFontMap, listMarkers, nativeDefaults);
                                pdf.BeginRunningContentSection(currentSectionIndex, continuationOptions);
                            } else pdf.BeginRunningContentSection(currentSectionIndex);
                        }
                        footnoteNumbersById.BeginSection(section);
                        IReadOnlyList<WordElement> elements = CollapseNativeParagraphElements(section.Elements);
                        List<PdfFootnote> footnotes = CollectNativeFootnotes(elements, footnoteNumbersById, documentEndnotes);
                        if (!pageFootnoteBodyProfileChecked && footnotes.Any(note => !note.IsEndnote)) {
                            pageFootnoteBodyProfileGap = GetNativePageFootnoteBodyProfileGap(sections, options);
                            pageFootnoteBodyProfileChecked = true;
                        }
                        if (TryPrepareNativePageFootnotes(flow, section, elements, footnoteNumbersById, options, nativeDefaults, nativeFontMap,
                            pageFootnoteBodyProfileGap, HasNativeStableFootnoteContinuationWidth(sections, currentSectionIndex, options)))
                            footnotes.RemoveAll(note => !note.IsEndnote);
                        else if (footnotes.Any(note => !note.IsEndnote))
                            footnoteNumbersById.ReportPlacementFallback();
                        CollectNativeRunningFootnoteFallback(section, footnotes, footnoteNumbersById, options);
                        if (currentSectionIndex + 1 == sections.Count) footnotes.AddRange(documentEndnotes);

                        if (TryRenderNativeEmptySectionMark(flow, section, elements, nativeDefaults, nativeFontMap, options)) {
                            RenderNativeFootnotes(flow, footnotes, section, footnoteNumbersById, options, nativeDefaults, nativeFontMap);
                            continue;
                        }

                        if (TryRenderNativeSectionColumns(
                            page,
                            section,
                            elements,
                            listMarkers,
                            listIndices,
                            footnoteNumbersById,
                            options,
                            tableOfContentsEntries,
                            headingDestinations,
                            nativeDefaults,
                            nativeFontMap,
                            balanceColumns: currentSectionIndex + 1 < sectionGroupEnd)) {
                            RenderNativeFootnotes(flow, footnotes, section, footnoteNumbersById, options, nativeDefaults, nativeFontMap);
                            continue;
                        }

                        for (int i = 0; i < elements.Count; i++) {
                            cancellationToken.ThrowIfCancellationRequested();
                            WordElement element = elements[i];
                            if (element is WordFootNote) {
                                continue;
                            }

                            if (TryRenderNativeJoinedParagraphs(flow, elements, ref i,
                                paragraph => listMarkers.TryGetValue(paragraph, out var joinMarker) ? joinMarker : null,
                                footnoteNumbersById, options, nativeDefaults, nativeFontMap)) continue;

                            if (TryRenderNativeList(
                                flow,
                                elements,
                                ref i,
                                listMarkers,
                                listIndices,
                                footnoteNumbersById,
                                nativeDefaults,
                                nativeFontMap,
                                headingDestinations)) {
                                continue;
                            }

                            RenderNativeElement(
                                flow,
                                element,
                                section,
                                paragraph => listMarkers.TryGetValue(paragraph, out var marker) ? marker : null,
                                GetNativeFootnoteNumbersForElement(elements, i, footnoteNumbersById),
                                footnoteNumbersById,
                                options,
                                tableOfContentsEntries,
                                headingDestinations,
                                sectionContentWidth,
                                nativeDefaults,
                                nativeFontMap,
                                nextElement: GetNextNativeRenderableElement(elements, i));
                        }

                        RenderNativeImplicitTerminalBodyBarParagraph(flow, section, options, nativeDefaults, nativeFontMap);
                        RenderNativeFootnotes(flow, footnotes, section, footnoteNumbersById, options, nativeDefaults, nativeFontMap);
                    }
                });

                sectionIndex = sectionGroupEnd;
            }

            // The lookup cache is conversion-scoped. Releasing it here preserves the pre-existing
            // post-conversion behavior and prevents disposed document state from being reused.
            ResetNativeStyleLookupCache(document);
            cancellationToken.ThrowIfCancellationRequested();
            return pdf;
        }

        private static bool IsNativeRenderedListStoryParagraph(WordDocument document, WordParagraph paragraph) {
            if (IsNativeDescendantOf(paragraph._paragraph, document._wordprocessingDocument.MainDocumentPart?.Document?.Body)) {
                return true;
            }

            foreach (WordSection section in document.Sections) {
                if (IsNativeDescendantOf(paragraph._paragraph, section.ResolveDefaultHeader()?._header) ||
                    IsNativeDescendantOf(paragraph._paragraph, section.ResolveDefaultFooter()?._footer)) {
                    return true;
                }

                if (section.DifferentFirstPage &&
                    (IsNativeDescendantOf(paragraph._paragraph, section.ResolveFirstHeader()?._header) ||
                     IsNativeDescendantOf(paragraph._paragraph, section.ResolveFirstFooter()?._footer))) {
                    return true;
                }

                if (section.DifferentOddAndEvenPages &&
                    (IsNativeDescendantOf(paragraph._paragraph, section.ResolveEvenHeader()?._header) ||
                     IsNativeDescendantOf(paragraph._paragraph, section.ResolveEvenFooter()?._footer))) {
                    return true;
                }
            }

            return false;
        }

        private static bool IsNativeDescendantOf(DocumentFormat.OpenXml.OpenXmlElement element, DocumentFormat.OpenXml.OpenXmlElement? ancestor) {
            if (ancestor == null) {
                return false;
            }

            for (DocumentFormat.OpenXml.OpenXmlElement? current = element.Parent; current != null; current = current.Parent) {
                if (ReferenceEquals(current, ancestor)) {
                    return true;
                }
            }

            return false;
        }

        private static int GetNativePdfSectionGroupEnd(IReadOnlyList<WordSection> sections, int startIndex, WordPdfExportOptions? options, IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers, bool canContinueRunningStories) {
            int endIndex = startIndex + 1;
            while (endIndex < sections.Count && CanMergeNativeContinuousSection(sections[endIndex - 1], sections[endIndex], options, listMarkers, canContinueRunningStories)) {
                endIndex++;
            }

            return endIndex;
        }

        private static bool CanMergeNativeContinuousSection(WordSection previous, WordSection current, WordPdfExportOptions? options, IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers, bool canContinueRunningStories) {
            if (current.BreakType != WordSectionBreakType.Continuous) {
                return false;
            }

            if (!NativePageSizesEquivalent(GetNativePageSize(previous, options), GetNativePageSize(current, options))) {
                return false;
            }

            (double Header, double Footer) previousExpansion = GetNativeHeaderFooterMarginExpansion(previous, options, listMarkers);
            (double Header, double Footer) currentExpansion = GetNativeHeaderFooterMarginExpansion(current, options, listMarkers);
            if (!NativeMarginsEquivalent(GetNativeMargins(previous, options, previousExpansion), GetNativeMargins(current, options, currentExpansion))) {
                return false;
            }

            if (current._sectionProperties?.GetFirstChild<W.PageNumberType>()?.Start?.HasValue == true) {
                return false;
            }

            return NativeSectionPageNumberingEquivalent(previous, current) &&
                (NativeSectionHeaderFooterReferencesEquivalent(previous, current) ||
                 (canContinueRunningStories && CanContinueNativeDefaultHeaderStory(previous, current, options)));
        }

        private static PdfCore.PdfPageParity? GetNativeSectionStartParity(WordSection section) {
            if (section.BreakType == WordSectionBreakType.OddPage) return PdfCore.PdfPageParity.Odd;
            if (section.BreakType == WordSectionBreakType.EvenPage) return PdfCore.PdfPageParity.Even;
            if (section.BreakType == WordSectionBreakType.NextPage &&
                section._sectionProperties.GetFirstChild<W.PageNumberType>()?.Start?.Value is int start && start > 0) {
                return start % 2 == 0 ? PdfCore.PdfPageParity.Even : PdfCore.PdfPageParity.Odd;
            }
            return null;
        }

        private static bool NativePageSizesEquivalent(PdfCore.PdfPageSize first, PdfCore.PdfPageSize second) =>
            NativeDoubleEquals(first.Width.ToPoints(), second.Width.ToPoints()) &&
            NativeDoubleEquals(first.Height.ToPoints(), second.Height.ToPoints());

        private static bool NativeMarginsEquivalent(PdfCore.PdfPageMargins first, PdfCore.PdfPageMargins second) =>
            NativeDoubleEquals(first.Left.ToPoints(), second.Left.ToPoints()) &&
            NativeDoubleEquals(first.Top.ToPoints(), second.Top.ToPoints()) &&
            NativeDoubleEquals(first.Right.ToPoints(), second.Right.ToPoints()) &&
            NativeDoubleEquals(first.Bottom.ToPoints(), second.Bottom.ToPoints());

        private static bool NativeDoubleEquals(double first, double second) =>
            Math.Abs(first - second) < 0.001D;

        private static bool NativeSectionHeaderFooterReferencesEquivalent(WordSection first, WordSection second) =>
            first.DifferentFirstPage == second.DifferentFirstPage &&
            (!HasNativeSectionStory(first, isFooter: false) || NativeDoubleEquals(checked((uint)(first.Margins.HeaderDistance.GetValueOrDefault().ToTwips())), checked((uint)(second.Margins.HeaderDistance.GetValueOrDefault().ToTwips())))) &&
            (!HasNativeSectionStory(first, isFooter: true) || NativeDoubleEquals(checked((uint)(first.Margins.FooterDistance.GetValueOrDefault().ToTwips())), checked((uint)(second.Margins.FooterDistance.GetValueOrDefault().ToTwips())))) &&
            ReferenceEquals(first.ResolveDefaultHeader()?._header, second.ResolveDefaultHeader()?._header) &&
            ReferenceEquals(first.ResolveDefaultFooter()?._footer, second.ResolveDefaultFooter()?._footer) &&
            (!first.DifferentFirstPage ||
                (ReferenceEquals(first.ResolveFirstHeader()?._header, second.ResolveFirstHeader()?._header) &&
                 ReferenceEquals(first.ResolveFirstFooter()?._footer, second.ResolveFirstFooter()?._footer))) &&
            (!first.DocumentOddEvenSettingEnabled ||
                (ReferenceEquals(first.ResolveEvenHeader()?._header, second.ResolveEvenHeader()?._header) &&
                 ReferenceEquals(first.ResolveEvenFooter()?._footer, second.ResolveEvenFooter()?._footer)));

        private static bool HasNativeSectionStory(WordSection section, bool isFooter) => isFooter
            ? section.ResolveDefaultFooter() != null ||
                (section.DifferentFirstPage && section.ResolveFirstFooter() != null) ||
                (section.DocumentOddEvenSettingEnabled && section.ResolveEvenFooter() != null)
            : section.ResolveDefaultHeader() != null ||
                (section.DifferentFirstPage && section.ResolveFirstHeader() != null) ||
                (section.DocumentOddEvenSettingEnabled && section.ResolveEvenHeader() != null);

        private static bool NativeSectionPageNumberingEquivalent(WordSection first, WordSection second) =>
            NativeOpenXmlChildEquivalent<W.PageNumberType>(first, second);

        private static bool NativeOpenXmlChildEquivalent<T>(WordSection first, WordSection second)
            where T : DocumentFormat.OpenXml.OpenXmlElement =>
            string.Equals(
                first._sectionProperties?.GetFirstChild<T>()?.OuterXml,
                second._sectionProperties?.GetFirstChild<T>()?.OuterXml,
                StringComparison.Ordinal);

        private static IReadOnlyList<WordElement> CollapseNativeParagraphElements(IEnumerable<WordElement> elements) {
            var collapsed = new List<WordElement>();
            var paragraphIndexes = new Dictionary<W.Paragraph, int>();

            foreach (WordElement element in elements) {
                if (element is WordParagraph paragraph && paragraph._paragraph != null) {
                    if (paragraphIndexes.TryGetValue(paragraph._paragraph, out int existingIndex)) {
                        if (ShouldReplaceNativeParagraphElement(collapsed[existingIndex], paragraph)) {
                            collapsed[existingIndex] = paragraph;
                        }

                        continue;
                    }

                    paragraphIndexes.Add(paragraph._paragraph, collapsed.Count);
                }

                collapsed.Add(element);
            }

            return collapsed;
        }

        private static bool ShouldReplaceNativeParagraphElement(WordElement existing, WordParagraph candidate) {
            if (existing is not WordParagraph existingParagraph) {
                return false;
            }

            if (string.IsNullOrEmpty(existingParagraph.Bookmark?.Name) &&
                !string.IsNullOrEmpty(candidate.Bookmark?.Name)) {
                return true;
            }

            return false;
        }

        private static WordElement? GetNextNativeRenderableElement(IReadOnlyList<WordElement> elements, int index) {
            for (int nextIndex = index + 1; nextIndex < elements.Count; nextIndex++) {
                WordElement next = elements[nextIndex];
                if (next is WordFootNote) {
                    continue;
                }

                return next;
            }

            return null;
        }

    }
}
