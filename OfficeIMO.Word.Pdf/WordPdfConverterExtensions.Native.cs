using System.Collections.Generic;
using System.Globalization;
using System.Text;
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
        private const double NativeDefaultParagraphSpacingAfter = 8D;
        private const double NativeWordSingleLineHeight = 1.22D;
        private const double NativeWordAutoLineSpacingHeight = NativeDefaultParagraphLineHeight;
        private const double NativeWordTableSingleLineHeight = 1.22D;
        private const double NativeTablePageContinuationSpacingBefore = 24D;
        private const double NativeHeaderFooterFontSize = 9D;
        private const double NativeHeaderFooterLineHeight = NativeHeaderFooterFontSize * 1.2D;
        private const double NativeHeaderFooterBodyGap = 2D;
        private const double NativeHeaderFooterDefaultOffset = 18D;
        private const double NativeFooterDefaultOffset = 20D;

        private interface INativePdfFlow {
            void PageBreak();
            void Spacer(double height);
            void Bookmark(string name);
            void HR(double? thickness = null, PdfCore.PdfColor? color = null, double? spacingBefore = null, double? spacingAfter = null, PdfCore.PdfHorizontalRuleStyle? style = null);
            void Paragraph(Action<PdfCore.PdfParagraphBuilder> build, PdfCore.PdfAlign align = PdfCore.PdfAlign.Left, PdfCore.PdfColor? defaultColor = null, PdfCore.PdfParagraphStyle? style = null);
            void PanelParagraph(Action<PdfCore.PdfParagraphBuilder> build, PdfCore.PanelStyle? style = null, PdfCore.PdfAlign align = PdfCore.PdfAlign.Left, PdfCore.PdfColor? defaultColor = null);
            void Heading(int level, string text, PdfCore.PdfAlign align, PdfCore.PdfColor? color, PdfCore.PdfHeadingStyle? style, string? linkUri, string? linkDestinationName, string? linkContents);
            void RichNumbered(IEnumerable<PdfCore.PdfListItem> items, PdfCore.PdfAlign align, PdfCore.PdfColor? color, int startNumber, PdfCore.PdfListStyle? style);
            void RichBullets(IEnumerable<PdfCore.PdfListItem> items, PdfCore.PdfAlign align, PdfCore.PdfColor? color, PdfCore.PdfListStyle? style);
            void TextField(string name, double width, double height, string value, PdfCore.PdfAlign align, double fontSize, double spacingBefore, double spacingAfter, PdfCore.PdfFormFieldStyle? style = null);
            void ChoiceField(string name, IEnumerable<string> options, string? value, double width, double height, PdfCore.PdfAlign align, double fontSize, double spacingBefore, double spacingAfter, bool isComboBox, PdfCore.PdfFormFieldStyle? style = null);
            void CheckBox(string name, bool isChecked, double size, PdfCore.PdfAlign align, double spacingBefore, double spacingAfter, string checkedValueName = "Yes", PdfCore.PdfFormFieldStyle? style = null);
            void Shape(OfficeShape shape, PdfCore.PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfCore.PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null);
            void Drawing(OfficeDrawing drawing, PdfCore.PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfCore.PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null);
            void Canvas(Action<PdfCore.PdfPageCanvas> build);
            void Table(IEnumerable<PdfCore.PdfTableCell[]> rows, PdfCore.PdfAlign align, PdfCore.PdfTableStyle? style);
            void Image(byte[] bytes, double width, double height, PdfCore.PdfAlign? align = null);
        }

        private sealed class NativePdfDocumentFlow : INativePdfFlow {
            private readonly PdfCore.PdfDocument _pdf;

            public NativePdfDocumentFlow(PdfCore.PdfDocument pdf) {
                _pdf = pdf;
            }

            public void PageBreak() => _pdf.PageBreak();
            public void Spacer(double height) => _pdf.Spacer(height);
            public void Bookmark(string name) => _pdf.Bookmark(name);
            public void HR(double? thickness = null, PdfCore.PdfColor? color = null, double? spacingBefore = null, double? spacingAfter = null, PdfCore.PdfHorizontalRuleStyle? style = null) => _pdf.HR(thickness, color, spacingBefore, spacingAfter, style);
            public void Paragraph(Action<PdfCore.PdfParagraphBuilder> build, PdfCore.PdfAlign align = PdfCore.PdfAlign.Left, PdfCore.PdfColor? defaultColor = null, PdfCore.PdfParagraphStyle? style = null) => _pdf.Paragraph(build, align, defaultColor, style);
            public void PanelParagraph(Action<PdfCore.PdfParagraphBuilder> build, PdfCore.PanelStyle? style = null, PdfCore.PdfAlign align = PdfCore.PdfAlign.Left, PdfCore.PdfColor? defaultColor = null) => _pdf.PanelParagraph(build, style, align, defaultColor);
            public void Heading(int level, string text, PdfCore.PdfAlign align, PdfCore.PdfColor? color, PdfCore.PdfHeadingStyle? style, string? linkUri, string? linkDestinationName, string? linkContents) {
                if (level == 1) _pdf.H1(text, align, color, linkUri: linkUri, style: style, linkContents: linkContents, linkDestinationName: linkDestinationName);
                else if (level == 2) _pdf.H2(text, align, color, linkUri: linkUri, style: style, linkContents: linkContents, linkDestinationName: linkDestinationName);
                else _pdf.H3(text, align, color, linkUri: linkUri, style: style, linkContents: linkContents, linkDestinationName: linkDestinationName);
            }
            public void RichNumbered(IEnumerable<PdfCore.PdfListItem> items, PdfCore.PdfAlign align, PdfCore.PdfColor? color, int startNumber, PdfCore.PdfListStyle? style) => _pdf.RichNumbered(items, align, color, startNumber, style);
            public void RichBullets(IEnumerable<PdfCore.PdfListItem> items, PdfCore.PdfAlign align, PdfCore.PdfColor? color, PdfCore.PdfListStyle? style) => _pdf.RichBullets(items, align, color, style);
            public void TextField(string name, double width, double height, string value, PdfCore.PdfAlign align, double fontSize, double spacingBefore, double spacingAfter, PdfCore.PdfFormFieldStyle? style) => _pdf.TextField(name, width, height, value, align, fontSize, spacingBefore, spacingAfter, style);
            public void ChoiceField(string name, IEnumerable<string> options, string? value, double width, double height, PdfCore.PdfAlign align, double fontSize, double spacingBefore, double spacingAfter, bool isComboBox, PdfCore.PdfFormFieldStyle? style) => _pdf.ChoiceField(name, options, value, width, height, align, fontSize, spacingBefore, spacingAfter, isComboBox, style);
            public void CheckBox(string name, bool isChecked, double size, PdfCore.PdfAlign align, double spacingBefore, double spacingAfter, string checkedValueName, PdfCore.PdfFormFieldStyle? style) => _pdf.CheckBox(name, isChecked, size, align, spacingBefore, spacingAfter, checkedValueName, style);
            public void Shape(OfficeShape shape, PdfCore.PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfCore.PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) => _pdf.Shape(shape, align, spacingBefore, spacingAfter, style, linkUri, linkContents);
            public void Drawing(OfficeDrawing drawing, PdfCore.PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfCore.PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) => _pdf.Drawing(drawing, align, spacingBefore, spacingAfter, style, linkUri, linkContents);
            public void Canvas(Action<PdfCore.PdfPageCanvas> build) => _pdf.Canvas(build);
            public void Table(IEnumerable<PdfCore.PdfTableCell[]> rows, PdfCore.PdfAlign align, PdfCore.PdfTableStyle? style) => _pdf.Table(rows, align, style);
            public void Image(byte[] bytes, double width, double height, PdfCore.PdfAlign? align = null) => _pdf.Image(bytes, width, height, align, style: CreateNativeImageStyle());
        }

        private sealed class NativePdfColumnFlow : INativePdfFlow {
            private readonly PdfCore.PdfPageCompose _page;
            private readonly PdfCore.PdfRowColumnCompose _column;

            public NativePdfColumnFlow(PdfCore.PdfPageCompose page, PdfCore.PdfRowColumnCompose column) {
                _page = page;
                _column = column;
            }

            public void PageBreak() => _column.PageBreak();
            public void Spacer(double height) => _column.Spacer(height);
            public void Bookmark(string name) => _column.Bookmark(name);
            public void HR(double? thickness = null, PdfCore.PdfColor? color = null, double? spacingBefore = null, double? spacingAfter = null, PdfCore.PdfHorizontalRuleStyle? style = null) => _column.HR(thickness, color, spacingBefore, spacingAfter, style);
            public void Paragraph(Action<PdfCore.PdfParagraphBuilder> build, PdfCore.PdfAlign align = PdfCore.PdfAlign.Left, PdfCore.PdfColor? defaultColor = null, PdfCore.PdfParagraphStyle? style = null) => _column.Paragraph(build, align, defaultColor, style);
            public void PanelParagraph(Action<PdfCore.PdfParagraphBuilder> build, PdfCore.PanelStyle? style = null, PdfCore.PdfAlign align = PdfCore.PdfAlign.Left, PdfCore.PdfColor? defaultColor = null) => _column.PanelParagraph(build, style, align, defaultColor);
            public void Heading(int level, string text, PdfCore.PdfAlign align, PdfCore.PdfColor? color, PdfCore.PdfHeadingStyle? style, string? linkUri, string? linkDestinationName, string? linkContents) {
                if (level == 1) _column.H1(text, align, color, linkUri: linkUri, style: style, linkContents: linkContents, linkDestinationName: linkDestinationName);
                else if (level == 2) _column.H2(text, align, color, linkUri: linkUri, style: style, linkContents: linkContents, linkDestinationName: linkDestinationName);
                else _column.H3(text, align, color, linkUri: linkUri, style: style, linkContents: linkContents, linkDestinationName: linkDestinationName);
            }
            public void RichNumbered(IEnumerable<PdfCore.PdfListItem> items, PdfCore.PdfAlign align, PdfCore.PdfColor? color, int startNumber, PdfCore.PdfListStyle? style) => _column.RichNumbered(items, align, color, startNumber, style);
            public void RichBullets(IEnumerable<PdfCore.PdfListItem> items, PdfCore.PdfAlign align, PdfCore.PdfColor? color, PdfCore.PdfListStyle? style) => _column.RichBullets(items, align, color, style);
            public void TextField(string name, double width, double height, string value, PdfCore.PdfAlign align, double fontSize, double spacingBefore, double spacingAfter, PdfCore.PdfFormFieldStyle? style) => _column.TextField(name, width, height, value, align, fontSize, spacingBefore, spacingAfter, style);
            public void ChoiceField(string name, IEnumerable<string> options, string? value, double width, double height, PdfCore.PdfAlign align, double fontSize, double spacingBefore, double spacingAfter, bool isComboBox, PdfCore.PdfFormFieldStyle? style) => _column.ChoiceField(name, options, value, width, height, align, fontSize, spacingBefore, spacingAfter, isComboBox, style);
            public void CheckBox(string name, bool isChecked, double size, PdfCore.PdfAlign align, double spacingBefore, double spacingAfter, string checkedValueName, PdfCore.PdfFormFieldStyle? style) => _column.CheckBox(name, isChecked, size, align, spacingBefore, spacingAfter, checkedValueName, style);
            public void Shape(OfficeShape shape, PdfCore.PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfCore.PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) => _column.Shape(shape, align, spacingBefore, spacingAfter, style, linkUri, linkContents);
            public void Drawing(OfficeDrawing drawing, PdfCore.PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfCore.PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) => _column.Drawing(drawing, align, spacingBefore, spacingAfter, style, linkUri, linkContents);
            public void Canvas(Action<PdfCore.PdfPageCanvas> build) => _page.Canvas(build);
            public void Table(IEnumerable<PdfCore.PdfTableCell[]> rows, PdfCore.PdfAlign align, PdfCore.PdfTableStyle? style) => _column.Table(rows, align, style);
            public void Image(byte[] bytes, double width, double height, PdfCore.PdfAlign? align = null) => _column.Image(bytes, width, height, align, style: CreateNativeImageStyle());
        }

        private static PdfCore.PdfImageStyle CreateNativeImageStyle() => new() {
            ScaleDownToFit = true
        };

        private static PdfCore.PdfDocument CreateOfficeIMOPdfDocument(WordDocument document, PdfSaveOptions? options) {
            ResetNativeStyleLookupCache(document);
            BuiltinDocumentProperties properties = document.BuiltinDocumentProperties;
            var nativeFontMap = new NativeFontMap(options?.Report);
            PdfCore.PdfDocument pdf = PdfCore.PdfDocument.Create(CreateNativeOptions(document, options, nativeFontMap))
                .Meta(
                    title: options?.Title ?? properties.Title,
                    author: options?.Author ?? properties.Creator,
                    subject: options?.Subject ?? properties.Subject,
                    keywords: BuildNativeKeywords(options, properties));

            Dictionary<WordParagraph, (int Level, string Marker)> listMarkers = DocumentTraversal.BuildListMarkers(document);
            Dictionary<WordParagraph, (int Level, int Index)> listIndices = DocumentTraversal.BuildListIndices(document);
            Dictionary<W.Paragraph, string> headingDestinations = BuildNativeHeadingDestinations(document);
            IReadOnlyList<NativeTableOfContentsEntry> tableOfContentsEntries = BuildNativeTableOfContentsEntries(document, options, headingDestinations);
            NativeDocumentDefaults nativeDefaults = GetNativeDocumentDefaults(document);
            var footnoteNumbersById = new Dictionary<long, int>();
            IReadOnlyList<WordSection> sections = document.Sections;
            for (int sectionIndex = 0; sectionIndex < sections.Count;) {
                int sectionGroupEnd = GetNativePdfSectionGroupEnd(sections, sectionIndex, options);
                WordSection firstSection = sections[sectionIndex];
                PdfCore.PageSize sectionPageSize = GetNativePageSize(firstSection, options);
                (double Header, double Footer) headerFooterMarginExpansion = GetNativeHeaderFooterMarginExpansion(firstSection, options);
                PdfCore.PageMargins sectionMargins = GetNativeMargins(firstSection, options, headerFooterMarginExpansion);
                double sectionContentWidth = Math.Max(72D, sectionPageSize.Width - sectionMargins.Left - sectionMargins.Right);
                pdf.Section(page => {
                    page.Size(sectionPageSize);
                    page.Margin(sectionMargins);
                    ConfigureNativePageNumbering(page, firstSection);
                    ConfigureNativeHeaderFooter(page, firstSection, options, headerFooterMarginExpansion.Header, headerFooterMarginExpansion.Footer, nativeFontMap);
                    INativePdfFlow flow = new NativeSpacingCollapseFlow(new NativePdfDocumentFlow(pdf));

                    for (int currentSectionIndex = sectionIndex; currentSectionIndex < sectionGroupEnd; currentSectionIndex++) {
                        WordSection section = sections[currentSectionIndex];
                        IReadOnlyList<WordElement> elements = CollapseNativeParagraphElements(section.Elements);
                        List<PdfFootnote> footnotes = CollectNativeFootnotes(elements, footnoteNumbersById);

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
                            nativeFontMap)) {
                            RenderNativeFootnotes(flow, footnotes);
                            continue;
                        }

                        for (int i = 0; i < elements.Count; i++) {
                            WordElement element = elements[i];
                            if (element is WordFootNote) {
                                continue;
                            }

                            if (TryRenderNativeList(
                                flow,
                                elements,
                                ref i,
                                listMarkers,
                                listIndices,
                                footnoteNumbersById,
                                nativeDefaults,
                                nativeFontMap)) {
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
                                renderSpacingOnlyEmptyParagraphLineBox: IsPreviousNativeElementTable(elements, i),
                                nextElement: GetNextNativeRenderableElement(elements, i));
                        }

                        RenderNativeFootnotes(flow, footnotes);
                    }
                });

                sectionIndex = sectionGroupEnd;
            }

            return pdf;
        }

        private static int GetNativePdfSectionGroupEnd(IReadOnlyList<WordSection> sections, int startIndex, PdfSaveOptions? options) {
            int endIndex = startIndex + 1;
            while (endIndex < sections.Count && CanMergeNativeContinuousSection(sections[endIndex - 1], sections[endIndex], options)) {
                endIndex++;
            }

            return endIndex;
        }

        private static bool CanMergeNativeContinuousSection(WordSection previous, WordSection current, PdfSaveOptions? options) {
            if (GetNativeSectionBreakAfter(previous) != W.SectionMarkValues.Continuous) {
                return false;
            }

            if (HasNativeSectionColumns(previous) || HasNativeSectionColumns(current)) {
                return false;
            }

            if (!NativePageSizesEquivalent(GetNativePageSize(previous, options), GetNativePageSize(current, options))) {
                return false;
            }

            (double Header, double Footer) previousExpansion = GetNativeHeaderFooterMarginExpansion(previous, options);
            (double Header, double Footer) currentExpansion = GetNativeHeaderFooterMarginExpansion(current, options);
            if (!NativeMarginsEquivalent(GetNativeMargins(previous, options, previousExpansion), GetNativeMargins(current, options, currentExpansion))) {
                return false;
            }

            return NativeSectionHeaderFooterReferencesEquivalent(previous, current) &&
                NativeSectionPageNumberingEquivalent(previous, current);
        }

        private static W.SectionMarkValues? GetNativeSectionBreakAfter(WordSection section) =>
            section._sectionProperties?.GetFirstChild<W.SectionType>()?.Val?.Value;

        private static bool HasNativeSectionColumns(WordSection section) =>
            (section.ColumnCount ?? 1) > 1 || section.HasColumnSeparator;

        private static bool NativePageSizesEquivalent(PdfCore.PageSize first, PdfCore.PageSize second) =>
            NativeDoubleEquals(first.Width, second.Width) &&
            NativeDoubleEquals(first.Height, second.Height);

        private static bool NativeMarginsEquivalent(PdfCore.PageMargins first, PdfCore.PageMargins second) =>
            NativeDoubleEquals(first.Left, second.Left) &&
            NativeDoubleEquals(first.Top, second.Top) &&
            NativeDoubleEquals(first.Right, second.Right) &&
            NativeDoubleEquals(first.Bottom, second.Bottom);

        private static bool NativeDoubleEquals(double first, double second) =>
            Math.Abs(first - second) < 0.001D;

        private static bool NativeSectionHeaderFooterReferencesEquivalent(WordSection first, WordSection second) =>
            NativeOpenXmlChildrenEquivalent<W.HeaderReference>(first, second) &&
            NativeOpenXmlChildrenEquivalent<W.FooterReference>(first, second) &&
            NativeOpenXmlChildEquivalent<W.TitlePage>(first, second);

        private static bool NativeSectionPageNumberingEquivalent(WordSection first, WordSection second) =>
            NativeOpenXmlChildEquivalent<W.PageNumberType>(first, second);

        private static bool NativeOpenXmlChildEquivalent<T>(WordSection first, WordSection second)
            where T : DocumentFormat.OpenXml.OpenXmlElement =>
            string.Equals(
                first._sectionProperties?.GetFirstChild<T>()?.OuterXml,
                second._sectionProperties?.GetFirstChild<T>()?.OuterXml,
                StringComparison.Ordinal);

        private static bool NativeOpenXmlChildrenEquivalent<T>(WordSection first, WordSection second)
            where T : DocumentFormat.OpenXml.OpenXmlElement {
            string[] firstValues = first._sectionProperties?.Elements<T>().Select(element => element.OuterXml).ToArray() ?? Array.Empty<string>();
            string[] secondValues = second._sectionProperties?.Elements<T>().Select(element => element.OuterXml).ToArray() ?? Array.Empty<string>();
            return firstValues.SequenceEqual(secondValues, StringComparer.Ordinal);
        }

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

        private static bool IsPreviousNativeElementTable(IReadOnlyList<WordElement> elements, int index) {
            for (int previousIndex = index - 1; previousIndex >= 0; previousIndex--) {
                WordElement previous = elements[previousIndex];
                if (previous is WordFootNote) {
                    continue;
                }

                return previous is WordTable;
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
