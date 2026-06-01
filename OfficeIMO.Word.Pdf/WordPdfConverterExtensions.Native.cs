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

        private interface INativePdfFlow {
            void PageBreak();
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
            void Table(IEnumerable<PdfCore.PdfTableCell[]> rows, PdfCore.PdfAlign align, PdfCore.PdfTableStyle? style);
            void Image(byte[] bytes, double width, double height, PdfCore.PdfAlign? align = null);
        }

        private sealed class NativePdfDocFlow : INativePdfFlow {
            private readonly PdfCore.PdfDoc _pdf;

            public NativePdfDocFlow(PdfCore.PdfDoc pdf) {
                _pdf = pdf;
            }

            public void PageBreak() => _pdf.PageBreak();
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
            public void Table(IEnumerable<PdfCore.PdfTableCell[]> rows, PdfCore.PdfAlign align, PdfCore.PdfTableStyle? style) => _pdf.Table(rows, align, style);
            public void Image(byte[] bytes, double width, double height, PdfCore.PdfAlign? align = null) => _pdf.Image(bytes, width, height, align);
        }

        private sealed class NativePdfColumnFlow : INativePdfFlow {
            private readonly PdfCore.PdfRowColumnCompose _column;

            public NativePdfColumnFlow(PdfCore.PdfRowColumnCompose column) {
                _column = column;
            }

            public void PageBreak() => _column.PageBreak();
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
            public void Table(IEnumerable<PdfCore.PdfTableCell[]> rows, PdfCore.PdfAlign align, PdfCore.PdfTableStyle? style) => _column.Table(rows, align, style);
            public void Image(byte[] bytes, double width, double height, PdfCore.PdfAlign? align = null) => _column.Image(bytes, width, height, align);
        }

        private static PdfCore.PdfDoc CreateOfficeIMOPdfDocument(WordDocument document, PdfSaveOptions? options) {
            options?.Warnings.Clear();

            BuiltinDocumentProperties properties = document.BuiltinDocumentProperties;
            PdfCore.PdfDoc pdf = PdfCore.PdfDoc.Create(CreateNativeOptions(document, options))
                .Meta(
                    title: options?.Title ?? properties.Title,
                    author: options?.Author ?? properties.Creator,
                    subject: options?.Subject ?? properties.Subject,
                    keywords: BuildNativeKeywords(options, properties));

            Dictionary<WordParagraph, (int Level, string Marker)> listMarkers = DocumentTraversal.BuildListMarkers(document);
            Dictionary<WordParagraph, (int Level, int Index)> listIndices = DocumentTraversal.BuildListIndices(document);
            Dictionary<W.Paragraph, string> headingDestinations = BuildNativeHeadingDestinations(document);
            IReadOnlyList<NativeTableOfContentsEntry> tableOfContentsEntries = BuildNativeTableOfContentsEntries(document, options, headingDestinations);
            foreach (WordSection section in document.Sections) {
                IReadOnlyList<WordElement> elements = CollapseNativeParagraphElements(section.Elements);
                List<PdfFootnote> footnotes = CollectNativeFootnotes(elements, out Dictionary<long, int> footnoteNumbersById);
                pdf.Section(page => {
                    page.Size(GetNativePageSize(section, options));
                    page.Margin(GetNativeMargins(section, options));
                    ConfigureNativePageNumbering(page, section);
                    ConfigureNativeHeaderFooter(page, section, options);
                    var flow = new NativePdfDocFlow(pdf);

                    if (TryRenderNativeSectionColumns(
                        page,
                        section,
                        elements,
                        listMarkers,
                        listIndices,
                        footnoteNumbersById,
                        options,
                        tableOfContentsEntries,
                        headingDestinations)) {
                        RenderNativeFootnotes(flow, footnotes);
                        return;
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
                            footnoteNumbersById)) {
                            continue;
                        }

                        RenderNativeElement(
                            flow,
                            element,
                            paragraph => listMarkers.TryGetValue(paragraph, out var marker) ? marker : null,
                            GetNativeFootnoteNumbersForElement(elements, i, footnoteNumbersById),
                            footnoteNumbersById,
                            options,
                            tableOfContentsEntries,
                            headingDestinations);
                    }

                    RenderNativeFootnotes(flow, footnotes);
                });
            }

            return pdf;
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

        private static bool TryRenderNativeSectionColumns(
            PdfCore.PdfPageCompose page,
            WordSection section,
            IReadOnlyList<WordElement> elements,
            Dictionary<WordParagraph, (int Level, string Marker)> listMarkers,
            Dictionary<WordParagraph, (int Level, int Index)> listIndices,
            Dictionary<long, int> footnoteNumbersById,
            PdfSaveOptions? options,
            IReadOnlyList<NativeTableOfContentsEntry> tableOfContentsEntries,
            IReadOnlyDictionary<W.Paragraph, string> headingDestinations) {
            IReadOnlyList<double> columnWidthPercents = GetNativeSectionColumnWidthPercents(section);
            int columnCount = columnWidthPercents.Count;
            if (columnCount <= 1) {
                return false;
            }

            IReadOnlyList<IReadOnlyList<WordElement>> columns = SplitNativeElementsByColumnBreaks(elements, columnCount);
            double gap = GetNativeSectionColumnGap(section);
            page.Content(content => content.Row(row => {
                row.Gap(gap);
                if (section.HasColumnSeparator) {
                    row.ColumnSeparator(PdfCore.PdfColor.Black, 0.5D);
                }

                for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                    IReadOnlyList<WordElement> columnElements = columns[columnIndex];
                    row.Column(columnWidthPercents[columnIndex], column => {
                        var flow = new NativePdfColumnFlow(column);
                        bool hasContent = false;
                        for (int i = 0; i < columnElements.Count; i++) {
                            WordElement element = columnElements[i];
                            if (element is WordFootNote) {
                                continue;
                            }

                            if (TryRenderNativeList(
                                flow,
                                columnElements,
                                ref i,
                                listMarkers,
                                listIndices,
                                footnoteNumbersById)) {
                                hasContent = true;
                                continue;
                            }

                            RenderNativeElement(
                                flow,
                                element,
                                paragraph => listMarkers.TryGetValue(paragraph, out var marker) ? marker : null,
                                GetNativeFootnoteNumbersForElement(columnElements, i, footnoteNumbersById),
                                footnoteNumbersById,
                                options,
                                tableOfContentsEntries,
                                headingDestinations);
                            hasContent = true;
                        }

                        if (!hasContent) {
                            column.Spacer(0);
                        }
                    });
                }
            }));
            return true;
        }

        private static int GetNativeSectionColumnCount(WordSection section) {
            int? explicitColumnCount = section._sectionProperties
                .GetFirstChild<W.Columns>()?
                .Elements<W.Column>()
                .Count();
            int count = section.ColumnCount ?? explicitColumnCount ?? 1;
            if (count < 1) {
                return 1;
            }

            return Math.Min(count, 8);
        }

        private static IReadOnlyList<double> GetNativeSectionColumnWidthPercents(WordSection section) {
            int columnCount = GetNativeSectionColumnCount(section);
            if (columnCount <= 1) {
                return new[] { 100D };
            }

            List<int>? explicitWidths = GetNativeExplicitSectionColumnWidths(section, columnCount);
            if (explicitWidths == null || explicitWidths.Count == 0) {
                return CreateEqualNativeColumnWidths(columnCount);
            }

            int total = explicitWidths.Sum();
            if (total <= 0) {
                return CreateEqualNativeColumnWidths(columnCount);
            }

            var widths = new List<double>(explicitWidths.Count);
            double accumulated = 0D;
            for (int i = 0; i < explicitWidths.Count; i++) {
                double percent = i == explicitWidths.Count - 1
                    ? 100D - accumulated
                    : explicitWidths[i] * 100D / total;
                widths.Add(percent);
                accumulated += percent;
            }

            return widths;
        }

        private static List<double> CreateEqualNativeColumnWidths(int columnCount) {
            var widths = new List<double>(columnCount);
            for (int i = 0; i < columnCount; i++) {
                widths.Add(100D / columnCount);
            }

            return widths;
        }

        private static List<int>? GetNativeExplicitSectionColumnWidths(WordSection section, int columnCount) {
            W.Columns? columns = section._sectionProperties.GetFirstChild<W.Columns>();
            if (columns == null) {
                return null;
            }

            var widths = new List<int>(columnCount);
            foreach (W.Column column in columns.Elements<W.Column>().Take(columnCount)) {
                if (!TryParseNativeTwips(column.Width?.Value, out int width) || width <= 0) {
                    return null;
                }

                widths.Add(width);
            }

            return widths.Count == columnCount ? widths : null;
        }

        private static double GetNativeSectionColumnGap(WordSection section) {
            double? gap = section.ColumnsSpace.HasValue ? ConvertNativeTwipsToPoints(section.ColumnsSpace.Value) : null;
            if (!gap.HasValue) {
                W.Column? firstColumn = section._sectionProperties.GetFirstChild<W.Columns>()?.Elements<W.Column>().FirstOrDefault();
                if (TryParseNativeTwips(firstColumn?.Space?.Value, out int columnGap)) {
                    gap = ConvertNativeTwipsToPoints(columnGap);
                }
            }

            if (!gap.HasValue || gap.Value < 0D) {
                return 36D;
            }

            return gap.Value;
        }

        private static bool TryParseNativeTwips(string? value, out int twips) {
            return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out twips);
        }

        private static IReadOnlyList<IReadOnlyList<WordElement>> SplitNativeElementsByColumnBreaks(IReadOnlyList<WordElement> elements, int columnCount) {
            var columns = new List<List<WordElement>>(columnCount);
            for (int i = 0; i < columnCount; i++) {
                columns.Add(new List<WordElement>());
            }

            bool sawColumnBreak = false;
            int currentColumn = 0;
            foreach (WordElement element in elements) {
                if (IsNativeColumnBreakElement(element)) {
                    sawColumnBreak = true;
                    AdvanceNativeColumn(columns, ref currentColumn);
                    continue;
                }

                if (element is WordParagraph paragraph) {
                    if (TrySplitNativeParagraphAtColumnBreak(paragraph, out WordParagraph? beforeColumnBreak, out WordParagraph? afterColumnBreak)) {
                        sawColumnBreak = true;
                        if (beforeColumnBreak != null) {
                            columns[currentColumn].Add(beforeColumnBreak);
                        }

                        AdvanceNativeColumn(columns, ref currentColumn);

                        if (afterColumnBreak != null) {
                            columns[currentColumn].Add(afterColumnBreak);
                        }

                        continue;
                    }

                    NativeColumnBreakPlacement columnBreakPlacement = GetNativeParagraphColumnBreakPlacement(paragraph);
                    if (columnBreakPlacement != NativeColumnBreakPlacement.None) {
                        sawColumnBreak = true;
                    }

                    if (columnBreakPlacement == NativeColumnBreakPlacement.StartsWithBreak) {
                        AdvanceNativeColumn(columns, ref currentColumn);
                    }

                    columns[currentColumn].Add(element);
                    if (columnBreakPlacement == NativeColumnBreakPlacement.EndsWithBreak ||
                        columnBreakPlacement == NativeColumnBreakPlacement.ContainsBreak) {
                        AdvanceNativeColumn(columns, ref currentColumn);
                    }

                    continue;
                }

                columns[currentColumn].Add(element);
            }

            if (!sawColumnBreak) {
                return SplitNativeElementsAcrossAutomaticColumns(elements, columnCount);
            }

            return columns;
        }

        private static bool TrySplitNativeParagraphAtColumnBreak(WordParagraph paragraph, out WordParagraph? before, out WordParagraph? after) {
            before = null;
            after = null;
            if (paragraph._paragraph == null) {
                return false;
            }

            var beforeParagraph = new W.Paragraph();
            var afterParagraph = new W.Paragraph();
            if (paragraph._paragraph.ParagraphProperties != null) {
                beforeParagraph.Append((W.ParagraphProperties)paragraph._paragraph.ParagraphProperties.CloneNode(true));
                afterParagraph.Append((W.ParagraphProperties)paragraph._paragraph.ParagraphProperties.CloneNode(true));
            }

            bool sawColumnBreak = false;
            foreach (DocumentFormat.OpenXml.OpenXmlElement child in paragraph._paragraph.ChildElements) {
                if (child is W.ParagraphProperties) {
                    continue;
                }

                if (!sawColumnBreak &&
                    TrySplitNativeOpenXmlAtColumnBreak(child, out DocumentFormat.OpenXml.OpenXmlElement? beforeChild, out DocumentFormat.OpenXml.OpenXmlElement? afterChild)) {
                    if (beforeChild != null) {
                        beforeParagraph.Append(beforeChild);
                    }

                    if (afterChild != null) {
                        afterParagraph.Append(afterChild);
                    }

                    sawColumnBreak = true;
                    continue;
                }

                if (sawColumnBreak) {
                    afterParagraph.Append(child.CloneNode(true));
                } else {
                    beforeParagraph.Append(child.CloneNode(true));
                }
            }

            if (!sawColumnBreak) {
                return false;
            }

            if (HasNativeRenderableOpenXmlContent(beforeParagraph)) {
                before = new WordParagraph(paragraph._document, beforeParagraph);
            }

            if (HasNativeRenderableOpenXmlContent(afterParagraph)) {
                after = new WordParagraph(paragraph._document, afterParagraph);
            }

            return true;
        }

        private static bool TrySplitNativeOpenXmlAtColumnBreak(
            DocumentFormat.OpenXml.OpenXmlElement element,
            out DocumentFormat.OpenXml.OpenXmlElement? before,
            out DocumentFormat.OpenXml.OpenXmlElement? after) {
            before = null;
            after = null;
            if (IsNativeColumnBreakOpenXml(element)) {
                return true;
            }

            if (!element.HasChildren) {
                return false;
            }

            DocumentFormat.OpenXml.OpenXmlElement beforeElement = element.CloneNode(false);
            DocumentFormat.OpenXml.OpenXmlElement afterElement = element.CloneNode(false);
            bool sawColumnBreak = false;
            bool containsColumnBreak = false;
            foreach (DocumentFormat.OpenXml.OpenXmlElement child in element.ChildElements) {
                if (!sawColumnBreak &&
                    TrySplitNativeOpenXmlAtColumnBreak(child, out DocumentFormat.OpenXml.OpenXmlElement? beforeChild, out DocumentFormat.OpenXml.OpenXmlElement? afterChild)) {
                    containsColumnBreak = true;
                    if (beforeChild != null) {
                        beforeElement.Append(beforeChild);
                    }

                    if (afterChild != null) {
                        afterElement.Append(afterChild);
                    }

                    sawColumnBreak = true;
                    continue;
                }

                if (sawColumnBreak) {
                    afterElement.Append(child.CloneNode(true));
                } else {
                    beforeElement.Append(child.CloneNode(true));
                }
            }

            if (!containsColumnBreak) {
                return false;
            }

            if (HasNativeRenderableOpenXmlContent(beforeElement)) {
                before = beforeElement;
            }

            if (HasNativeRenderableOpenXmlContent(afterElement)) {
                after = afterElement;
            }

            return true;
        }

        private static bool IsNativeColumnBreakOpenXml(DocumentFormat.OpenXml.OpenXmlElement element) =>
            element is W.Break wordBreak && wordBreak.Type?.Value == W.BreakValues.Column;

        private static bool HasNativeRenderableOpenXmlContent(DocumentFormat.OpenXml.OpenXmlElement element) {
            if (element.Descendants<W.Text>().Any(text => !string.IsNullOrEmpty(text.Text)) ||
                element.Descendants<W.TabChar>().Any() ||
                element.Descendants<W.Drawing>().Any() ||
                element.Descendants<W.FootnoteReference>().Any() ||
                element.Descendants<W.EndnoteReference>().Any() ||
                element.Descendants<DocumentFormat.OpenXml.Math.OfficeMath>().Any() ||
                element.Descendants<DocumentFormat.OpenXml.Vml.Shape>().Any()) {
                return true;
            }

            return element.Descendants<W.Break>().Any(wordBreak => wordBreak.Type?.Value != W.BreakValues.Column);
        }

        private static IReadOnlyList<IReadOnlyList<WordElement>> SplitNativeElementsAcrossAutomaticColumns(IReadOnlyList<WordElement> elements, int columnCount) {
            var columns = new List<List<WordElement>>(columnCount);
            for (int i = 0; i < columnCount; i++) {
                columns.Add(new List<WordElement>());
            }

            if (elements.Count == 0) {
                return columns;
            }

            int totalWeight = 0;
            var weights = new int[elements.Count];
            for (int i = 0; i < elements.Count; i++) {
                int weight = GetNativeAutomaticColumnWeight(elements[i]);
                weights[i] = weight;
                totalWeight += weight;
            }

            int currentColumn = 0;
            int currentWeight = 0;
            for (int i = 0; i < elements.Count; i++) {
                int remainingElements = elements.Count - i;
                int remainingColumnsAfterCurrent = columnCount - currentColumn - 1;
                if (currentColumn < columnCount - 1 &&
                    columns[currentColumn].Count > 0) {
                    double targetWeight = (double)totalWeight * (currentColumn + 1) / columnCount;
                    if (remainingElements <= remainingColumnsAfterCurrent ||
                        currentWeight >= targetWeight) {
                        if (!TryAdvanceNativeAutomaticColumnKeepingTrailingContent(columns, ref currentColumn)) {
                            currentColumn++;
                        }
                    }
                }

                columns[currentColumn].Add(elements[i]);
                currentWeight += weights[i];
            }

            return columns;
        }

        private static bool TryAdvanceNativeAutomaticColumnKeepingTrailingContent(List<List<WordElement>> columns, ref int currentColumn) {
            if (currentColumn >= columns.Count - 1) {
                return false;
            }

            List<WordElement> current = columns[currentColumn];
            if (current.Count <= 1) {
                return false;
            }

            int moveStart = current.Count - 1;
            if (!ShouldKeepNativeElementWithFollowingContent(current[moveStart])) {
                return false;
            }

            while (moveStart > 0 && ShouldKeepNativeElementWithFollowingContent(current[moveStart - 1])) {
                moveStart--;
            }

            if (moveStart == 0) {
                return false;
            }

            List<WordElement> next = columns[currentColumn + 1];
            for (int i = moveStart; i < current.Count; i++) {
                next.Add(current[i]);
            }

            current.RemoveRange(moveStart, current.Count - moveStart);
            currentColumn++;
            return true;
        }

        private static bool ShouldKeepNativeElementWithFollowingContent(WordElement element) =>
            element is WordParagraph paragraph &&
            (paragraph.KeepWithNext || GetHeadingLevel(paragraph) > 0);

        private static int GetNativeAutomaticColumnWeight(WordElement element) {
            if (element is WordParagraph paragraph) {
                return Math.Max(1, (paragraph.Text?.Length ?? 0) / 80 + 1);
            }

            if (element is WordTable table) {
                return Math.Max(2, table.Rows.Count * 2);
            }

            return 1;
        }

        private static void AdvanceNativeColumn(List<List<WordElement>> columns, ref int currentColumn) {
            if (columns[currentColumn].Count > 0) {
                currentColumn = Math.Min(columns.Count - 1, currentColumn + 1);
            }
        }

        private static bool IsNativeColumnBreakElement(WordElement element) {
            if (element is WordBreak wordBreak) {
                return wordBreak.BreakType == W.BreakValues.Column;
            }

            return element is WordParagraph paragraph &&
                paragraph.Break?.BreakType == W.BreakValues.Column &&
                string.IsNullOrWhiteSpace(paragraph.Text);
        }

        private enum NativeColumnBreakPlacement {
            None,
            StartsWithBreak,
            EndsWithBreak,
            ContainsBreak
        }

        private static NativeColumnBreakPlacement GetNativeParagraphColumnBreakPlacement(WordParagraph paragraph) {
            if (paragraph._paragraph == null) {
                return NativeColumnBreakPlacement.None;
            }

            bool sawColumnBreak = false;
            bool hasContentBefore = false;
            bool hasContentAfter = false;
            InspectNativeColumnBreakFlow(paragraph._paragraph, ref sawColumnBreak, ref hasContentBefore, ref hasContentAfter);
            if (!sawColumnBreak) {
                return NativeColumnBreakPlacement.None;
            }

            if (hasContentBefore && hasContentAfter) {
                return NativeColumnBreakPlacement.ContainsBreak;
            }

            if (hasContentBefore) {
                return NativeColumnBreakPlacement.EndsWithBreak;
            }

            if (hasContentAfter) {
                return NativeColumnBreakPlacement.StartsWithBreak;
            }

            return NativeColumnBreakPlacement.None;
        }

        private static void InspectNativeColumnBreakFlow(DocumentFormat.OpenXml.OpenXmlElement element, ref bool sawColumnBreak, ref bool hasContentBefore, ref bool hasContentAfter) {
            foreach (DocumentFormat.OpenXml.OpenXmlElement child in element.ChildElements) {
                if (child is W.Break wordBreak && wordBreak.Type?.Value == W.BreakValues.Column) {
                    sawColumnBreak = true;
                    continue;
                }

                if (child is W.Text text && !string.IsNullOrEmpty(text.Text)) {
                    if (sawColumnBreak) {
                        hasContentAfter = true;
                    } else {
                        hasContentBefore = true;
                    }
                } else {
                    InspectNativeColumnBreakFlow(child, ref sawColumnBreak, ref hasContentBefore, ref hasContentAfter);
                }
            }
        }

        private static bool TryRenderNativeList(
            INativePdfFlow pdf,
            IReadOnlyList<WordElement> elements,
            ref int index,
            Dictionary<WordParagraph, (int Level, string Marker)> listMarkers,
            Dictionary<WordParagraph, (int Level, int Index)> listIndices,
            Dictionary<long, int> footnoteNumbersById) {
            if (elements[index] is not WordParagraph firstParagraph ||
                !TryGetNativeListItem(firstParagraph, listMarkers, listIndices, footnoteNumbersById, out bool ordered, out int level, out int startNumber, out PdfCore.PdfListItem? item, out PdfCore.PdfAlign align, out PdfCore.PdfColor? color, out PdfCore.PdfListStyle? style)) {
                return false;
            }

            var items = new List<PdfCore.PdfListItem> { item! };
            int nextIndex = index + 1;
            int expectedNumber = startNumber + 1;
            while (nextIndex < elements.Count &&
                   elements[nextIndex] is WordParagraph paragraph &&
                   TryGetNativeListItem(paragraph, listMarkers, listIndices, footnoteNumbersById, out bool nextOrdered, out int nextLevel, out int nextNumber, out PdfCore.PdfListItem? nextItem, out PdfCore.PdfAlign nextAlign, out PdfCore.PdfColor? nextColor, out PdfCore.PdfListStyle? nextStyle) &&
                   nextOrdered == ordered &&
                   nextLevel == level &&
                   nextAlign == align &&
                   nextColor.Equals(color) &&
                   NativeListStylesEquivalent(nextStyle, style) &&
                   (!ordered || nextNumber == expectedNumber)) {
                items.Add(nextItem!);
                nextIndex++;
                expectedNumber++;
            }

            if (ordered) {
                pdf.RichNumbered(items, align, color, startNumber, style);
            } else {
                pdf.RichBullets(items, align, color, style);
            }

            index = nextIndex - 1;
            return true;
        }

        private static bool TryGetNativeListItem(
            WordParagraph paragraph,
            Dictionary<WordParagraph, (int Level, string Marker)> listMarkers,
            Dictionary<WordParagraph, (int Level, int Index)> listIndices,
            Dictionary<long, int> footnoteNumbersById,
            out bool ordered,
            out int level,
            out int index,
            out PdfCore.PdfListItem? item,
            out PdfCore.PdfAlign align,
            out PdfCore.PdfColor? color,
            out PdfCore.PdfListStyle? style) {
            ordered = false;
            level = 0;
            index = 1;
            item = null;
            align = PdfCore.PdfAlign.Left;
            color = null;
            style = null;

            if (!listMarkers.TryGetValue(paragraph, out var marker) ||
                !listIndices.TryGetValue(paragraph, out var listIndex)) {
                return false;
            }

            DocumentTraversal.ListInfo? info = DocumentTraversal.GetListInfo(paragraph);
            if (info == null || marker.Level != info.Value.Level || listIndex.Level != info.Value.Level) {
                return false;
            }

            if (paragraph.PageBreakBefore ||
                paragraph.IsPageBreak ||
                paragraph.Shape != null ||
                paragraph.TextBox != null ||
                paragraph.Image != null) {
                return false;
            }

            List<WordParagraph> runs = GetNativeRuns(paragraph);
            if (runs.Any(run => run.IsImage)) {
                return false;
            }

            List<PdfCore.TextRun> richRuns = CreateNativeCellParagraphRuns(paragraph, footnoteNumbersById);
            string content = string.Concat(richRuns.Select(run => run.Text));
            if (string.IsNullOrWhiteSpace(content)) {
                return false;
            }

            bool itemOrdered = info.Value.Ordered;
            string displayMarker = itemOrdered
                ? marker.Marker
                : NormalizeNativeBulletMarker(marker.Marker);
            ordered = itemOrdered;
            level = info.Value.Level;
            index = listIndex.Index;
            item = new PdfCore.PdfListItem(richRuns, paragraph.Bookmark?.Name, string.IsNullOrWhiteSpace(displayMarker) ? null : displayMarker);
            align = MapNativeParagraphAlign(paragraph.ParagraphAlignment, allowJustify: false);
            color = ParseNativeColor(paragraph.ColorHex);
            style = CreateNativeListStyle(paragraph, info.Value, displayMarker);
            return true;
        }

        private static string NormalizeNativeBulletMarker(string marker) {
            if (string.IsNullOrWhiteSpace(marker)) {
                return "•";
            }

            return marker.Trim() switch {
                "\uf0b7" => "•",
                "\u00b7" => "•",
                "\u25cf" => "•",
                "\u006f" => "o",
                _ => marker
            };
        }

        private static PdfCore.PdfListStyle CreateNativeListStyle(WordParagraph paragraph, DocumentTraversal.ListInfo info, string marker) {
            const double defaultLevelTextIndent = 36D;
            const double defaultHangingIndent = 18D;

            double textIndent = ConvertNativeTwipsToPoints(info.LeftIndentTwips ?? ((info.Level + 1) * 720)) ?? ((info.Level + 1) * defaultLevelTextIndent);
            double hangingIndent = ConvertNativeTwipsToPoints(info.HangingIndentTwips ?? 360) ?? defaultHangingIndent;
            double markerIndent = Math.Max(0D, textIndent - hangingIndent);
            double fontSize = paragraph.FontSize.HasValue && paragraph.FontSize.Value > 0D ? paragraph.FontSize.Value : 11D;
            double markerWidth = EstimateNativeListMarkerWidth(marker, fontSize);
            double markerGap = Math.Max(0D, textIndent - markerIndent - markerWidth);

            var style = new PdfCore.PdfListStyle {
                LeftIndent = markerIndent,
                MarkerGap = markerGap
            };

            if (paragraph.FontSize.HasValue && paragraph.FontSize.Value > 0D) {
                style.FontSize = paragraph.FontSize.Value;
            }

            style.LineHeight = ResolveNativeParagraphLineHeight(paragraph, fontSize);

            if (paragraph.LineSpacingBeforePoints.HasValue) {
                style.SpacingBefore = paragraph.LineSpacingBeforePoints.Value;
            }

            if (paragraph.LineSpacingAfterPoints.HasValue) {
                style.SpacingAfter = paragraph.LineSpacingAfterPoints.Value;
            }

            style.KeepTogether = paragraph.KeepLinesTogether;
            style.KeepWithNext = paragraph.KeepWithNext;
            return style;
        }

        private static double EstimateNativeListMarkerWidth(string marker, double fontSize) {
            if (string.IsNullOrEmpty(marker)) {
                return 0D;
            }

            double width = 0D;
            foreach (char ch in marker) {
                if (char.IsDigit(ch) || char.IsLetter(ch)) {
                    width += fontSize * 0.56D;
                } else if (char.IsWhiteSpace(ch)) {
                    width += fontSize * 0.28D;
                } else if (ch == '.' || ch == ')' || ch == '(') {
                    width += fontSize * 0.28D;
                } else if (ch == '\u2022' || ch == '\u25CF' || ch == '\u25E6') {
                    width += fontSize * 0.36D;
                } else {
                    width += fontSize * 0.5D;
                }
            }

            return width;
        }

        private static bool NativeListStylesEquivalent(PdfCore.PdfListStyle? left, PdfCore.PdfListStyle? right) {
            if (ReferenceEquals(left, right)) {
                return true;
            }

            if (left == null || right == null) {
                return false;
            }

            return NullableDoubleEquals(left.FontSize, right.FontSize) &&
                   NullableDoubleEquals(left.LineHeight, right.LineHeight) &&
                   DoubleEquals(left.LeftIndent, right.LeftIndent) &&
                   NullableDoubleEquals(left.MarkerGap, right.MarkerGap) &&
                   DoubleEquals(left.SpacingBefore, right.SpacingBefore) &&
                   NullableDoubleEquals(left.SpacingAfter, right.SpacingAfter) &&
                   NullableDoubleEquals(left.ItemSpacing, right.ItemSpacing) &&
                   left.Color.Equals(right.Color) &&
                   left.KeepTogether == right.KeepTogether &&
                   left.KeepWithNext == right.KeepWithNext;
        }

        private static bool NullableDoubleEquals(double? left, double? right) {
            if (left.HasValue != right.HasValue) {
                return false;
            }

            return !left.HasValue || DoubleEquals(left.Value, right!.Value);
        }

        private static bool DoubleEquals(double left, double right) =>
            Math.Abs(left - right) < 0.001D;

        private static List<WordParagraph> GetNativeRuns(WordParagraph paragraph) {
            if (paragraph._paragraph == null) {
                return new List<WordParagraph>();
            }

            var runs = new List<WordParagraph>();
            foreach (var element in paragraph._paragraph.ChildElements) {
                if (element is W.Run run) {
                    runs.Add(new WordParagraph(paragraph._document, paragraph._paragraph, run));
                } else if (element is W.Hyperlink hyperlink) {
                    AddNativeHyperlinkRuns(runs, paragraph, hyperlink);
                } else if (element is W.SdtRun sdtRun && IsNativeSimpleTextContentControl(sdtRun)) {
                    foreach (var childElement in sdtRun.SdtContentRun!.ChildElements) {
                        if (childElement is W.Run sdtContentRun) {
                            runs.Add(new WordParagraph(paragraph._document, paragraph._paragraph, sdtContentRun));
                        } else if (childElement is W.Hyperlink sdtHyperlink) {
                            AddNativeHyperlinkRuns(runs, paragraph, sdtHyperlink);
                        }
                    }
                }
            }

            return runs;
        }

        private static void AddNativeHyperlinkRuns(List<WordParagraph> runs, WordParagraph paragraph, W.Hyperlink hyperlink) {
            foreach (W.Run childRun in hyperlink.Elements<W.Run>()) {
                var run = new WordParagraph(paragraph._document, paragraph._paragraph!, childRun) { _hyperlink = hyperlink };
                runs.Add(run);
            }
        }

        private static void ConfigureNativePageNumbering(PdfCore.PdfPageCompose page, WordSection section) {
            W.PageNumberType? pageNumberType = section._sectionProperties.GetFirstChild<W.PageNumberType>();
            if (pageNumberType?.Start?.Value is int start && start > 0) {
                page.PageNumberStart(start);
            }

            PdfCore.PdfPageNumberStyle? style = MapNativePageNumberStyle(pageNumberType?.Format?.Value);
            if (style.HasValue) {
                page.PageNumberStyle(style.Value);
            }
        }

        private static PdfCore.PdfPageNumberStyle? MapNativePageNumberStyle(W.NumberFormatValues? format) {
            if (format == W.NumberFormatValues.LowerRoman) {
                return PdfCore.PdfPageNumberStyle.LowerRoman;
            }

            if (format == W.NumberFormatValues.UpperRoman) {
                return PdfCore.PdfPageNumberStyle.UpperRoman;
            }

            if (format == W.NumberFormatValues.LowerLetter) {
                return PdfCore.PdfPageNumberStyle.LowerLetter;
            }

            if (format == W.NumberFormatValues.UpperLetter) {
                return PdfCore.PdfPageNumberStyle.UpperLetter;
            }

            if (format == W.NumberFormatValues.Decimal || format == W.NumberFormatValues.DecimalZero) {
                return PdfCore.PdfPageNumberStyle.Arabic;
            }

            return null;
        }

        private static void ConfigureNativeHeaderFooter(PdfCore.PdfPageCompose page, WordSection section, PdfSaveOptions? options) {
            RecordNativeHeaderFooterDiagnostics(section.Header?.Default, options, "default header");
            RecordNativeHeaderFooterDiagnostics(section.Header?.First, options, "first header");
            RecordNativeHeaderFooterDiagnostics(section.Header?.Even, options, "even header");
            RecordNativeHeaderFooterDiagnostics(section.Footer?.Default, options, "default footer");
            RecordNativeHeaderFooterDiagnostics(section.Footer?.First, options, "first footer");
            RecordNativeHeaderFooterDiagnostics(section.Footer?.Even, options, "even footer");

            NativeHeaderFooterText? defaultHeader = GetNativeHeaderFooterText(section.Header?.Default);
            NativeHeaderFooterText? firstHeader = GetNativeHeaderFooterText(section.Header?.First);
            NativeHeaderFooterText? evenHeader = GetNativeHeaderFooterText(section.Header?.Even);
            NativeHeaderFooterText? defaultFooter = GetNativeHeaderFooterText(section.Footer?.Default);
            NativeHeaderFooterText? firstFooter = GetNativeHeaderFooterText(section.Footer?.First);
            NativeHeaderFooterText? evenFooter = GetNativeHeaderFooterText(section.Footer?.Even);
            IReadOnlyList<NativeHeaderFooterImage> defaultHeaderImages = GetNativeHeaderFooterImages(section.Header?.Default, options, "default header image");
            IReadOnlyList<NativeHeaderFooterImage> firstHeaderImages = GetNativeHeaderFooterImages(section.Header?.First, options, "first header image");
            IReadOnlyList<NativeHeaderFooterImage> evenHeaderImages = GetNativeHeaderFooterImages(section.Header?.Even, options, "even header image");
            IReadOnlyList<NativeHeaderFooterImage> defaultFooterImages = GetNativeHeaderFooterImages(section.Footer?.Default, options, "default footer image");
            IReadOnlyList<NativeHeaderFooterImage> firstFooterImages = GetNativeHeaderFooterImages(section.Footer?.First, options, "first footer image");
            IReadOnlyList<NativeHeaderFooterImage> evenFooterImages = GetNativeHeaderFooterImages(section.Footer?.Even, options, "even footer image");
            IReadOnlyList<NativeHeaderFooterShape> defaultHeaderShapes = GetNativeHeaderFooterShapes(section.Header?.Default);
            IReadOnlyList<NativeHeaderFooterShape> firstHeaderShapes = GetNativeHeaderFooterShapes(section.Header?.First);
            IReadOnlyList<NativeHeaderFooterShape> evenHeaderShapes = GetNativeHeaderFooterShapes(section.Header?.Even);
            IReadOnlyList<NativeHeaderFooterShape> defaultFooterShapes = GetNativeHeaderFooterShapes(section.Footer?.Default);
            IReadOnlyList<NativeHeaderFooterShape> firstFooterShapes = GetNativeHeaderFooterShapes(section.Footer?.First);
            IReadOnlyList<NativeHeaderFooterShape> evenFooterShapes = GetNativeHeaderFooterShapes(section.Footer?.Even);
            ApplyNativeHeaderFooterPageNumberStyle(page, defaultHeader, firstHeader, evenHeader, defaultFooter, firstFooter, evenFooter);
            bool hasFirstHeaderVariant = section.DifferentFirstPage || firstHeader != null || firstHeaderImages.Count > 0 || firstHeaderShapes.Count > 0;
            bool hasEvenHeaderVariant = section.DifferentOddAndEvenPages || evenHeader != null || evenHeaderImages.Count > 0 || evenHeaderShapes.Count > 0;
            bool hasFirstFooterVariant = section.DifferentFirstPage || firstFooter != null || firstFooterImages.Count > 0 || firstFooterShapes.Count > 0;
            bool hasEvenFooterVariant = section.DifferentOddAndEvenPages || evenFooter != null || evenFooterImages.Count > 0 || evenFooterShapes.Count > 0;
            if (defaultHeader != null || hasFirstHeaderVariant || hasEvenHeaderVariant ||
                defaultHeaderImages.Count > 0 || firstHeaderImages.Count > 0 || evenHeaderImages.Count > 0 ||
                defaultHeaderShapes.Count > 0 || firstHeaderShapes.Count > 0 || evenHeaderShapes.Count > 0) {
                page.Header(header => {
                    if (defaultHeader != null) {
                        header.Zones(defaultHeader.Left, defaultHeader.Center, defaultHeader.Right);
                    }

                    AddNativeHeaderImages(header, defaultHeaderImages, W.HeaderFooterValues.Default);
                    AddNativeHeaderShapes(header, defaultHeaderShapes, W.HeaderFooterValues.Default);

                    if (firstHeader != null) {
                        header.FirstPageZones(firstHeader.Left, firstHeader.Center, firstHeader.Right);
                    } else if (hasFirstHeaderVariant) {
                        header.FirstPageText(string.Empty);
                    }

                    AddNativeHeaderImages(header, firstHeaderImages, W.HeaderFooterValues.First);
                    AddNativeHeaderShapes(header, firstHeaderShapes, W.HeaderFooterValues.First);

                    if (evenHeader != null) {
                        header.EvenPagesZones(evenHeader.Left, evenHeader.Center, evenHeader.Right);
                    } else if (hasEvenHeaderVariant) {
                        header.EvenPagesText(string.Empty);
                    }

                    AddNativeHeaderImages(header, evenHeaderImages, W.HeaderFooterValues.Even);
                    AddNativeHeaderShapes(header, evenHeaderShapes, W.HeaderFooterValues.Even);
                });
            }

            bool includePageNumbers = options?.IncludePageNumbers ?? true;
            if (!includePageNumbers && defaultFooter == null && !hasFirstFooterVariant && !hasEvenFooterVariant &&
                defaultFooterImages.Count == 0 && firstFooterImages.Count == 0 && evenFooterImages.Count == 0 &&
                defaultFooterShapes.Count == 0 && firstFooterShapes.Count == 0 && evenFooterShapes.Count == 0) {
                return;
            }

            string pageNumberFormat = GetNativePageNumberFormat(options);
            page.Footer(footer => {
                NativeHeaderFooterText? resolvedDefaultFooter = WithNativeFooterPageNumber(defaultFooter, includePageNumbers, pageNumberFormat);
                if (resolvedDefaultFooter != null) {
                    footer.Zones(resolvedDefaultFooter.Left, resolvedDefaultFooter.Center, resolvedDefaultFooter.Right);
                }

                AddNativeFooterImages(footer, defaultFooterImages, W.HeaderFooterValues.Default);
                AddNativeFooterShapes(footer, defaultFooterShapes, W.HeaderFooterValues.Default);

                NativeHeaderFooterText? resolvedFirstFooter = WithNativeFooterPageNumber(firstFooter, includePageNumbers && firstFooter != null, pageNumberFormat);
                if (resolvedFirstFooter != null) {
                    footer.FirstPageZones(resolvedFirstFooter.Left, resolvedFirstFooter.Center, resolvedFirstFooter.Right);
                } else if (hasFirstFooterVariant) {
                    footer.FirstPageText(string.Empty);
                }

                AddNativeFooterImages(footer, firstFooterImages, W.HeaderFooterValues.First);
                AddNativeFooterShapes(footer, firstFooterShapes, W.HeaderFooterValues.First);

                NativeHeaderFooterText? resolvedEvenFooter = WithNativeFooterPageNumber(evenFooter, includePageNumbers && evenFooter != null, pageNumberFormat);
                if (resolvedEvenFooter != null) {
                    footer.EvenPagesZones(resolvedEvenFooter.Left, resolvedEvenFooter.Center, resolvedEvenFooter.Right);
                } else if (hasEvenFooterVariant) {
                    footer.EvenPagesText(string.Empty);
                }

                AddNativeFooterImages(footer, evenFooterImages, W.HeaderFooterValues.Even);
                AddNativeFooterShapes(footer, evenFooterShapes, W.HeaderFooterValues.Even);
            });
        }

        private static void AddNativeHeaderImages(PdfCore.PdfHeaderCompose header, IReadOnlyList<NativeHeaderFooterImage> images, W.HeaderFooterValues variant) {
            foreach (NativeHeaderFooterImage image in images) {
                if (variant == W.HeaderFooterValues.First) {
                    header.FirstPageImage(image.Data, image.Width, image.Height, image.Align);
                } else if (variant == W.HeaderFooterValues.Even) {
                    header.EvenPagesImage(image.Data, image.Width, image.Height, image.Align);
                } else {
                    header.Image(image.Data, image.Width, image.Height, image.Align);
                }
            }
        }

        private static void AddNativeHeaderShapes(PdfCore.PdfHeaderCompose header, IReadOnlyList<NativeHeaderFooterShape> shapes, W.HeaderFooterValues variant) {
            foreach (NativeHeaderFooterShape shape in shapes) {
                if (variant == W.HeaderFooterValues.First) {
                    header.FirstPageShape(shape.Shape, shape.Align);
                } else if (variant == W.HeaderFooterValues.Even) {
                    header.EvenPagesShape(shape.Shape, shape.Align);
                } else {
                    header.Shape(shape.Shape, shape.Align);
                }
            }
        }

        private static void AddNativeFooterImages(PdfCore.PdfFooterCompose footer, IReadOnlyList<NativeHeaderFooterImage> images, W.HeaderFooterValues variant) {
            foreach (NativeHeaderFooterImage image in images) {
                if (variant == W.HeaderFooterValues.First) {
                    footer.FirstPageImage(image.Data, image.Width, image.Height, image.Align);
                } else if (variant == W.HeaderFooterValues.Even) {
                    footer.EvenPagesImage(image.Data, image.Width, image.Height, image.Align);
                } else {
                    footer.Image(image.Data, image.Width, image.Height, image.Align);
                }
            }
        }

        private static void AddNativeFooterShapes(PdfCore.PdfFooterCompose footer, IReadOnlyList<NativeHeaderFooterShape> shapes, W.HeaderFooterValues variant) {
            foreach (NativeHeaderFooterShape shape in shapes) {
                if (variant == W.HeaderFooterValues.First) {
                    footer.FirstPageShape(shape.Shape, shape.Align);
                } else if (variant == W.HeaderFooterValues.Even) {
                    footer.EvenPagesShape(shape.Shape, shape.Align);
                } else {
                    footer.Shape(shape.Shape, shape.Align);
                }
            }
        }

        private static void RecordNativeHeaderFooterDiagnostics(WordHeaderFooter? headerFooter, PdfSaveOptions? options, string source) {
            if (headerFooter == null || options == null) {
                return;
            }

            foreach (WordElement element in headerFooter.Elements) {
                RecordNativeHeaderFooterElementDiagnostics(element, options, source);
            }
        }

        private static void RecordNativeHeaderFooterElementDiagnostics(WordElement element, PdfSaveOptions options, string source) {
            switch (element) {
                case WordParagraph paragraph:
                    RecordNativeHeaderFooterParagraphDiagnostics(paragraph, options, source);
                    break;
                case WordTable table:
                    RecordNativeHeaderFooterTableDiagnostics(table, options, source + " table");
                    break;
                case WordEmbeddedDocument:
                    AddNativeExportWarning(
                        options,
                        "NativeHeaderFooterEmbeddedDocumentUnsupported",
                        source,
                        "Embedded documents in Word headers and footers are not mapped by the OfficeIMO PDF engine yet.");
                    break;
            }
        }

        private static void RecordNativeHeaderFooterTableDiagnostics(WordTable table, PdfSaveOptions options, string source) {
            foreach (WordTableRow row in table.Rows) {
                foreach (WordTableCell cell in row.Cells) {
                    foreach (WordElement element in cell.Elements) {
                        RecordNativeHeaderFooterElementDiagnostics(element, options, source);
                    }
                }
            }
        }

        private static void RecordNativeHeaderFooterParagraphDiagnostics(WordParagraph paragraph, PdfSaveOptions options, string source) {
            if (paragraph.Shape != null && CreateNativeShape(paragraph.Shape) == null) {
                AddNativeExportWarning(
                    options,
                    "NativeHeaderFooterShapeUnsupported",
                    source,
                    "Word header and footer shapes without supported geometry are not mapped by the OfficeIMO PDF engine yet.");
            }

            if (HasNativeUnsupportedHeaderFooterTextBox(paragraph)) {
                AddNativeExportWarning(
                    options,
                    "NativeHeaderFooterTextBoxUnsupported",
                    source,
                    "Word header and footer text boxes without extractable text are not mapped by the OfficeIMO PDF engine yet.");
            }

            if (paragraph.IsSmartArt) {
                AddNativeExportWarning(
                    options,
                    "NativeHeaderFooterSmartArtUnsupported",
                    source,
                    "SmartArt in Word headers and footers is not mapped by the OfficeIMO PDF engine yet.");
            }

            if (paragraph.IsEquation && string.IsNullOrWhiteSpace(GetNativeEquationText(paragraph))) {
                AddNativeExportWarning(
                    options,
                    "NativeHeaderFooterEquationUnsupported",
                    source,
                    "Equations in Word headers and footers are not mapped by the OfficeIMO PDF engine yet.");
            }

            if (HasNativeUnsupportedHeaderFooterContentControl(paragraph)) {
                AddNativeExportWarning(
                    options,
                    "NativeHeaderFooterContentControlUnsupported",
                    source,
                    "Content controls in Word headers and footers are not mapped by the OfficeIMO PDF engine yet.");
            }
        }

        private static bool HasNativeUnsupportedHeaderFooterTextBox(WordParagraph paragraph) =>
            paragraph.TextBox != null &&
            string.IsNullOrWhiteSpace(GetNativeParagraphTextBoxPlainText(paragraph));

        private static void RecordNativeBodyTableDiagnostics(WordTable table, PdfSaveOptions? options, string source) {
            if (options == null) {
                return;
            }

            foreach (WordTableRow row in table.Rows) {
                foreach (WordTableCell cell in row.Cells) {
                    foreach (WordParagraph paragraph in cell.Paragraphs) {
                        RecordNativeBodyParagraphDiagnostics(paragraph, options, source, mapsCheckBoxes: true, mapsFormFields: true, mapsPictureControls: true, mapsRepeatingSections: true);
                    }

                    foreach (WordElement element in cell.Elements) {
                        if (element is not WordParagraph) {
                            RecordNativeBodyElementDiagnostics(element, options, source);
                        }
                    }
                }
            }
        }

        private static void RecordNativeBodyElementDiagnostics(WordElement element, PdfSaveOptions options, string source) {
            switch (element) {
                case WordParagraph paragraph:
                    RecordNativeBodyParagraphDiagnostics(paragraph, options, source, mapsCheckBoxes: false, mapsFormFields: false, mapsPictureControls: false, mapsRepeatingSections: false);
                    break;
                case WordTable table:
                    RecordNativeBodyTableDiagnostics(table, options, source + " table");
                    break;
                case WordEmbeddedDocument:
                    AddNativeExportWarning(
                        options,
                        "NativeBodyEmbeddedDocumentUnsupported",
                        source,
                        "Embedded documents in Word body content are not mapped by the OfficeIMO PDF engine yet.");
                    break;
            }
        }

        private static void RecordNativeBodyParagraphDiagnostics(WordParagraph paragraph, PdfSaveOptions? options, string source, bool mapsCheckBoxes, bool mapsFormFields, bool mapsPictureControls, bool mapsRepeatingSections) {
            if (options == null) {
                return;
            }

            if (paragraph.IsSmartArt) {
                AddNativeExportWarning(
                    options,
                    "NativeBodySmartArtUnsupported",
                    source,
                    "SmartArt in Word body content is not mapped by the OfficeIMO PDF engine yet.");
            }

            if (paragraph.IsEquation && string.IsNullOrWhiteSpace(GetNativeEquationText(paragraph))) {
                AddNativeExportWarning(
                    options,
                    "NativeBodyEquationUnsupported",
                    source,
                    "Equations in Word body content are not mapped by the OfficeIMO PDF engine yet.");
            }

            if (HasNativeUnsupportedBodyContentControl(paragraph, mapsCheckBoxes, mapsFormFields, mapsPictureControls, mapsRepeatingSections)) {
                AddNativeExportWarning(
                    options,
                    "NativeBodyContentControlUnsupported",
                    source,
                    "Content controls in Word body content are not mapped by the OfficeIMO PDF engine yet.");
            }
        }

        private static bool HasNativeUnsupportedHeaderFooterContentControl(WordParagraph paragraph) =>
            (paragraph.IsCheckBox && GetNativeCheckBoxControls(paragraph).Count == 0) ||
            ((paragraph.IsDatePicker || paragraph.IsDropDownList || paragraph.IsComboBox) && GetNativeFormFieldControls(paragraph).Count == 0) ||
            (paragraph.IsPictureControl && paragraph.PictureControl?.Image == null) ||
            (paragraph.IsRepeatingSection && paragraph.RepeatingSection?.TextItems.Count == 0) ||
            paragraph._paragraph?.Descendants<W.SdtRun>().Any(sdtRun =>
                !IsNativeSimpleTextContentControl(sdtRun) &&
                !IsNativeCheckBoxControl(sdtRun) &&
                !IsNativeSupportedFormFieldContentControl(sdtRun) &&
                !IsNativePictureControlWithImage(paragraph, sdtRun) &&
                !IsNativeRepeatingSectionWithText(sdtRun) &&
                !IsNativeRepeatingSectionChildControl(sdtRun)) == true ||
            paragraph._paragraph?.Descendants<W.SdtBlock>().Any() == true ||
            paragraph._paragraph?.Descendants<W.SdtCell>().Any() == true;

        private static bool HasNativeUnsupportedBodyContentControl(WordParagraph paragraph, bool mapsCheckBoxes, bool mapsFormFields, bool mapsPictureControls, bool mapsRepeatingSections) =>
            (!mapsCheckBoxes && paragraph.IsCheckBox) ||
            (!mapsFormFields && (paragraph.IsDatePicker || paragraph.IsDropDownList || paragraph.IsComboBox)) ||
            (paragraph.IsPictureControl && (!mapsPictureControls || paragraph.PictureControl?.Image == null)) ||
            (paragraph.IsRepeatingSection && (!mapsRepeatingSections || paragraph.RepeatingSection?.TextItems.Count == 0)) ||
            paragraph._paragraph?.Descendants<W.SdtRun>().Any(sdtRun =>
                (!mapsCheckBoxes || !IsNativeCheckBoxControl(sdtRun)) &&
                (!mapsFormFields || !IsNativeSupportedFormFieldContentControl(sdtRun)) &&
                (!mapsPictureControls || !IsNativePictureControl(sdtRun)) &&
                (!mapsRepeatingSections || !IsNativeRepeatingSectionControl(sdtRun) && !IsNativeRepeatingSectionChildControl(sdtRun)) &&
                !IsNativeSimpleTextContentControl(sdtRun)) == true ||
            paragraph._paragraph?.Descendants<W.SdtBlock>().Any() == true ||
            paragraph._paragraph?.Descendants<W.SdtCell>().Any() == true;

        private static IReadOnlyList<W.SdtRun> GetNativeCheckBoxControls(WordParagraph paragraph) {
            if (paragraph._paragraph == null) {
                return Array.Empty<W.SdtRun>();
            }

            return paragraph._paragraph.Descendants<W.SdtRun>()
                .Where(IsNativeCheckBoxControl)
                .ToList();
        }

        private static bool IsNativeCheckBoxControl(W.SdtRun sdtRun) =>
            sdtRun.SdtProperties?.Elements<W14.SdtContentCheckBox>().Any() == true;

        private static bool IsNativePictureControl(W.SdtRun sdtRun) =>
            sdtRun.SdtProperties?.Elements<W.SdtContentPicture>().Any() == true;

        private static bool IsNativePictureControlWithImage(WordParagraph paragraph, W.SdtRun sdtRun) {
            if (!IsNativePictureControl(sdtRun) || paragraph._paragraph == null) {
                return false;
            }

            var pictureParagraph = new WordParagraph(paragraph._document, paragraph._paragraph, sdtRun);
            return pictureParagraph.PictureControl?.Image != null;
        }

        private static IReadOnlyList<W.SdtRun> GetNativePictureControls(WordParagraph paragraph) {
            if (paragraph._paragraph == null) {
                return Array.Empty<W.SdtRun>();
            }

            return paragraph._paragraph.Descendants<W.SdtRun>()
                .Where(IsNativePictureControl)
                .ToList();
        }

        private static bool IsNativeRepeatingSectionControl(W.SdtRun sdtRun) =>
            sdtRun.SdtProperties?.Elements<W15.SdtRepeatedSection>().Any() == true;

        private static bool IsNativeRepeatingSectionChildControl(W.SdtRun sdtRun) =>
            sdtRun.Ancestors<W.SdtRun>().Any(IsNativeRepeatingSectionControl);

        private static bool IsNativeRepeatingSectionWithText(W.SdtRun sdtRun) =>
            IsNativeRepeatingSectionControl(sdtRun) &&
            GetNativeRepeatingSectionItems(sdtRun).Count > 0;

        private static IReadOnlyList<W.SdtRun> GetNativeFormFieldControls(WordParagraph paragraph) {
            if (paragraph._paragraph == null) {
                return Array.Empty<W.SdtRun>();
            }

            return paragraph._paragraph.Descendants<W.SdtRun>()
                .Where(IsNativeSupportedFormFieldContentControl)
                .ToList();
        }

        private static IReadOnlyList<W.SdtRun> GetNativeRepeatingSectionControls(WordParagraph paragraph) {
            if (paragraph._paragraph == null) {
                return Array.Empty<W.SdtRun>();
            }

            return paragraph._paragraph.Descendants<W.SdtRun>()
                .Where(IsNativeRepeatingSectionControl)
                .ToList();
        }

        private static bool IsNativeSupportedFormFieldContentControl(W.SdtRun sdtRun) =>
            IsNativeDatePickerControl(sdtRun) ||
            GetNativeChoiceFieldOptions(sdtRun).Count > 0;

        private static bool IsNativeDatePickerControl(W.SdtRun sdtRun) =>
            sdtRun.SdtProperties?.Elements<W.SdtContentDate>().Any() == true;

        private static IReadOnlyList<string> GetNativeChoiceFieldOptions(W.SdtRun sdtRun) {
            W.SdtProperties? properties = sdtRun.SdtProperties;
            if (properties == null) {
                return Array.Empty<string>();
            }

            IEnumerable<W.ListItem> items = properties.Elements<W.SdtContentDropDownList>().FirstOrDefault()?.Elements<W.ListItem>() ??
                properties.Elements<W.SdtContentComboBox>().FirstOrDefault()?.Elements<W.ListItem>() ??
                Enumerable.Empty<W.ListItem>();

            var options = new List<string>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (W.ListItem item in items) {
                string? option = item.DisplayText?.Value ?? item.Value?.Value;
                if (string.IsNullOrWhiteSpace(option) || !seen.Add(option!)) {
                    continue;
                }

                options.Add(option!);
            }

            return options;
        }

        private static string? GetNativeChoiceFieldValue(W.SdtRun sdtRun, IReadOnlyList<string> options) {
            W.SdtContentComboBox? comboBox = sdtRun.SdtProperties?.Elements<W.SdtContentComboBox>().FirstOrDefault();
            string? lastValue = comboBox?.LastValue?.Value;
            if (!string.IsNullOrWhiteSpace(lastValue)) {
                string? displayValue = GetNativeChoiceDisplayValue(sdtRun, lastValue!);
                if (!string.IsNullOrWhiteSpace(displayValue) && options.Contains(displayValue!, StringComparer.Ordinal)) {
                    return displayValue;
                }

                if (options.Contains(lastValue!, StringComparer.Ordinal)) {
                    return lastValue;
                }
            }

            string? contentText = GetNativeSdtText(sdtRun);
            if (!string.IsNullOrWhiteSpace(contentText) && options.Contains(contentText!, StringComparer.Ordinal)) {
                return contentText;
            }

            return options.Count > 0 ? options[0] : null;
        }

        private static string? GetNativeChoiceDisplayValue(W.SdtRun sdtRun, string value) {
            IEnumerable<W.ListItem> items = sdtRun.SdtProperties?.Elements<W.SdtContentDropDownList>().FirstOrDefault()?.Elements<W.ListItem>() ??
                sdtRun.SdtProperties?.Elements<W.SdtContentComboBox>().FirstOrDefault()?.Elements<W.ListItem>() ??
                Enumerable.Empty<W.ListItem>();

            W.ListItem? match = items.FirstOrDefault(item => string.Equals(item.Value?.Value, value, StringComparison.Ordinal));
            return match?.DisplayText?.Value ?? match?.Value?.Value;
        }

        private static string GetNativeDatePickerValue(W.SdtRun sdtRun) {
            W.SdtContentDate? datePicker = sdtRun.SdtProperties?.Elements<W.SdtContentDate>().FirstOrDefault();
            if (datePicker?.FullDate?.Value is DateTime value) {
                return value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            }

            return GetNativeSdtText(sdtRun) ?? string.Empty;
        }

        private static string? GetNativeSdtText(W.SdtRun sdtRun) {
            if (sdtRun.SdtContentRun == null) {
                return null;
            }

            string text = string.Concat(sdtRun.SdtContentRun.Descendants<W.Text>().Select(runText => runText.Text));
            return string.IsNullOrWhiteSpace(text) ? null : text;
        }

        private static bool IsNativeSimpleTextContentControl(W.SdtRun sdtRun) {
            W.SdtProperties? properties = sdtRun.SdtProperties;
            if (properties == null) {
                return false;
            }

            if (properties.Elements<W14.SdtContentCheckBox>().Any() ||
                properties.Elements<W.SdtContentDate>().Any() ||
                properties.Elements<W.SdtContentDropDownList>().Any() ||
                properties.Elements<W.SdtContentComboBox>().Any() ||
                properties.Elements<W.SdtContentPicture>().Any() ||
                properties.Elements<W15.SdtRepeatedSection>().Any()) {
                return false;
            }

            return sdtRun.SdtContentRun?.Descendants<W.Text>().Any() == true;
        }

        private static bool IsNativeCheckBoxChecked(W.SdtRun sdtRun) {
            W14.SdtContentCheckBox? checkBox = sdtRun.SdtProperties?.Elements<W14.SdtContentCheckBox>().FirstOrDefault();
            W14.Checked? checkedState = checkBox?.Elements<W14.Checked>().FirstOrDefault();
            return checkedState?.Val?.Value == W14.OnOffValues.One;
        }

        private static string GetNativeCheckBoxFieldName(W.SdtRun sdtRun, int index, string fallbackPrefix = "WordCheckBox") {
            return GetNativeContentControlFieldName(sdtRun, index, fallbackPrefix);
        }

        private static string GetNativeContentControlFieldName(W.SdtRun sdtRun, int index, string fallbackPrefix) {
            string? tag = sdtRun.SdtProperties?.Elements<W.Tag>().FirstOrDefault()?.Val?.Value;
            if (!string.IsNullOrWhiteSpace(tag)) {
                return tag!;
            }

            string? alias = sdtRun.SdtProperties?.Elements<W.SdtAlias>().FirstOrDefault()?.Val?.Value;
            if (!string.IsNullOrWhiteSpace(alias)) {
                return alias!;
            }

            int? sdtId = sdtRun.SdtProperties?.Elements<W.SdtId>().FirstOrDefault()?.Val?.Value;
            return sdtId.HasValue
                ? fallbackPrefix + "." + sdtId.Value.ToString(CultureInfo.InvariantCulture)
                : fallbackPrefix + "." + (index + 1).ToString(CultureInfo.InvariantCulture);
        }

        private static void AddNativeExportWarning(PdfSaveOptions options, string code, string source, string message) {
            options.Warnings.Add(new PdfExportWarning(code, source, message));
        }

        private static void ApplyNativeHeaderFooterPageNumberStyle(PdfCore.PdfPageCompose page, params NativeHeaderFooterText?[] parts) {
            PdfCore.PdfPageNumberStyle? style = null;
            foreach (NativeHeaderFooterText? part in parts) {
                if (part?.PageNumberStyle == null) {
                    continue;
                }

                if (style.HasValue && style.Value != part.PageNumberStyle.Value) {
                    return;
                }

                style = part.PageNumberStyle.Value;
            }

            if (style.HasValue) {
                page.PageNumberStyle(style.Value);
            }
        }

        private static NativeHeaderFooterText? WithNativeFooterPageNumber(NativeHeaderFooterText? footer, bool includePageNumber, string pageNumberFormat) {
            if (!includePageNumber) {
                return footer;
            }

            if (footer?.HasPageTokens == true) {
                return footer;
            }

            NativeHeaderFooterText result = footer?.Clone() ?? new NativeHeaderFooterText();
            result.AppendRight(pageNumberFormat);
            return result;
        }

        private static NativeHeaderFooterText? GetNativeHeaderFooterText(WordHeaderFooter? headerFooter) {
            if (headerFooter == null) {
                return null;
            }

            var parts = new NativeHeaderFooterText();
            foreach (WordElement element in CollapseNativeParagraphElements(headerFooter.Elements)) {
                switch (element) {
                    case WordParagraph paragraph:
                        AddNativeHeaderFooterParagraphText(parts, paragraph);
                        break;
                    case WordTable table:
                        AddNativeHeaderFooterTableText(parts, table);
                        break;
                    case WordHyperLink link when !string.IsNullOrWhiteSpace(link.Text):
                        parts.AppendLeft(link.Text);
                        break;
                }
            }

            return parts.HasContent ? parts : null;
        }

        private static IReadOnlyList<NativeHeaderFooterImage> GetNativeHeaderFooterImages(WordHeaderFooter? headerFooter, PdfSaveOptions? options, string source) {
            if (headerFooter == null) {
                return Array.Empty<NativeHeaderFooterImage>();
            }

            var images = new List<NativeHeaderFooterImage>();
            foreach (WordElement element in headerFooter.Elements) {
                switch (element) {
                    case WordParagraph paragraph:
                        AddNativeHeaderFooterParagraphImage(images, paragraph, null, options, source);
                        break;
                    case WordTable table:
                        AddNativeHeaderFooterTableImages(images, table, options, source);
                        break;
                }
            }

            return images;
        }

        private static IReadOnlyList<NativeHeaderFooterShape> GetNativeHeaderFooterShapes(WordHeaderFooter? headerFooter) {
            if (headerFooter == null) {
                return Array.Empty<NativeHeaderFooterShape>();
            }

            var shapes = new List<NativeHeaderFooterShape>();
            foreach (WordElement element in headerFooter.Elements) {
                switch (element) {
                    case WordParagraph paragraph:
                        AddNativeHeaderFooterParagraphShape(shapes, paragraph, null);
                        break;
                    case WordTable table:
                        AddNativeHeaderFooterTableShapes(shapes, table);
                        break;
                }
            }

            return shapes;
        }

        private static void AddNativeHeaderFooterParagraphText(NativeHeaderFooterText parts, WordParagraph paragraph) {
            string? text = GetNativeHeaderFooterParagraphText(paragraph, out PdfCore.PdfPageNumberStyle? pageNumberStyle, out NativeHeaderFooterZone? zoneOverride);

            if (string.IsNullOrWhiteSpace(text)) {
                return;
            }

            string resolvedText = text!;
            if (zoneOverride.HasValue) {
                parts.Append(zoneOverride.Value, resolvedText, pageNumberStyle);
                return;
            }

            W.JustificationValues? alignment = paragraph.ParagraphAlignment;
            if (alignment == W.JustificationValues.Center) {
                parts.AppendCenter(resolvedText, pageNumberStyle);
            } else if (alignment == W.JustificationValues.Right) {
                parts.AppendRight(resolvedText, pageNumberStyle);
            } else {
                parts.AppendLeft(resolvedText, pageNumberStyle);
            }
        }

        private static void AddNativeHeaderFooterTableImages(List<NativeHeaderFooterImage> images, WordTable table, PdfSaveOptions? options, string source) {
            foreach (WordTableRow row in table.Rows) {
                IReadOnlyList<WordTableCell> cells = row.Cells;
                if (cells.Count == 1) {
                    foreach (WordParagraph paragraph in GetNativeCellParagraphs(cells[0])) {
                        AddNativeHeaderFooterParagraphImage(images, paragraph, null, options, source);
                    }

                    continue;
                }

                for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                    PdfCore.PdfAlign align = cellIndex == 0
                        ? PdfCore.PdfAlign.Left
                        : cellIndex == cells.Count - 1
                            ? PdfCore.PdfAlign.Right
                            : PdfCore.PdfAlign.Center;

                    foreach (WordParagraph paragraph in GetNativeCellParagraphs(cells[cellIndex])) {
                        AddNativeHeaderFooterParagraphImage(images, paragraph, align, options, source);
                    }
                }
            }
        }

        private static void AddNativeHeaderFooterParagraphImage(List<NativeHeaderFooterImage> images, WordParagraph paragraph, PdfCore.PdfAlign? alignOverride, PdfSaveOptions? options, string source) {
            PdfCore.PdfAlign align = alignOverride ?? MapNativeParagraphAlign(paragraph.ParagraphAlignment, allowJustify: false);
            if (paragraph.Image != null) {
                AddNativeHeaderFooterImage(images, paragraph.Image, align, options, source);
            }

            foreach (W.SdtRun pictureControl in GetNativePictureControls(paragraph)) {
                var pictureParagraph = new WordParagraph(paragraph._document, paragraph._paragraph!, pictureControl);
                WordImage? pictureControlImage = pictureParagraph.PictureControl?.Image;
                if (pictureControlImage == null) {
                    continue;
                }

                AddNativeHeaderFooterImage(images, pictureControlImage, align, options, source);
            }
        }

        private static void AddNativeHeaderFooterImage(List<NativeHeaderFooterImage> images, WordImage image, PdfCore.PdfAlign align, PdfSaveOptions? options, string source) {
            byte[] bytes = ImageEmbedder.GetImageBytes(image);
            if (!IsNativePdfSupportedImageBytes(bytes, out string? unsupportedReason)) {
                if (options != null) {
                    AddNativeExportWarning(
                        options,
                        "NativeHeaderFooterImageUnsupported",
                        source,
                        "Word header/footer image was not exported because the first-party PDF image writer supports JPEG and simple PNG images only. " + unsupportedReason);
                }

                return;
            }

            double width = image.Width.HasValue ? image.Width.Value * 72D / 96D : 144D;
            double height = image.Height.HasValue ? image.Height.Value * 72D / 96D : 144D;
            images.Add(new NativeHeaderFooterImage(bytes, width, height, align));
        }

        private static void AddNativeHeaderFooterTableShapes(List<NativeHeaderFooterShape> shapes, WordTable table) {
            foreach (WordTableRow row in table.Rows) {
                IReadOnlyList<WordTableCell> cells = row.Cells;
                if (cells.Count == 1) {
                    foreach (WordParagraph paragraph in GetNativeCellParagraphs(cells[0])) {
                        AddNativeHeaderFooterParagraphShape(shapes, paragraph, null);
                    }

                    continue;
                }

                for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                    PdfCore.PdfAlign align = cellIndex == 0
                        ? PdfCore.PdfAlign.Left
                        : cellIndex == cells.Count - 1
                            ? PdfCore.PdfAlign.Right
                            : PdfCore.PdfAlign.Center;

                    foreach (WordParagraph paragraph in GetNativeCellParagraphs(cells[cellIndex])) {
                        AddNativeHeaderFooterParagraphShape(shapes, paragraph, align);
                    }
                }
            }
        }

        private static void AddNativeHeaderFooterParagraphShape(List<NativeHeaderFooterShape> shapes, WordParagraph paragraph, PdfCore.PdfAlign? alignOverride) {
            if (paragraph.Shape == null) {
                return;
            }

            OfficeShape? shape = CreateNativeShape(paragraph.Shape);
            if (shape == null) {
                return;
            }

            PdfCore.PdfAlign align = alignOverride ?? MapNativeParagraphAlign(paragraph.ParagraphAlignment, allowJustify: false);
            shapes.Add(new NativeHeaderFooterShape(shape, align));
        }

        private static string? GetNativeHeaderFooterParagraphText(WordParagraph paragraph, out PdfCore.PdfPageNumberStyle? pageNumberStyle) {
            return GetNativeHeaderFooterParagraphText(paragraph, out pageNumberStyle, out _);
        }

        private static string? GetNativeHeaderFooterParagraphText(WordParagraph paragraph, out PdfCore.PdfPageNumberStyle? pageNumberStyle, out NativeHeaderFooterZone? zoneOverride) {
            zoneOverride = null;
            if (TryBuildNativeHeaderFooterParagraphText(paragraph, out string? mixedText, out pageNumberStyle)) {
                return AppendNativeHeaderFooterSupplementalText(mixedText, paragraph);
            }

            if (TryGetNativeHeaderFooterFieldToken(paragraph, out string? fieldToken, out pageNumberStyle)) {
                return AppendNativeHeaderFooterSupplementalText(fieldToken, paragraph);
            }

            pageNumberStyle = null;
            if (paragraph.IsHyperLink && paragraph.Hyperlink != null) {
                return AppendNativeHeaderFooterSupplementalText(paragraph.Hyperlink.Text, paragraph);
            }

            List<WordParagraph> runs = GetNativeRuns(paragraph);
            string? text = runs.Count > 0
                ? string.Concat(runs.Select(run => run.Text))
                : paragraph.Text;
            text = AppendNativeHeaderFooterSupplementalText(text, paragraph);
            if (!string.IsNullOrWhiteSpace(text)) {
                return text;
            }

            string? textBoxText = GetNativeParagraphTextBoxPlainText(paragraph);
            if (string.IsNullOrWhiteSpace(textBoxText)) {
                return text;
            }

            WordTextBox? textBox = GetNativeParagraphTextBox(paragraph, out _);
            zoneOverride = MapNativeTextBoxHeaderFooterZone(textBox?.HorizontalAlignment ?? WordHorizontalAlignmentValues.Center);
            return textBoxText;
        }

        private static string? AppendNativeHeaderFooterSupplementalText(string? text, WordParagraph paragraph) {
            text = AppendNativeHeaderFooterEquationText(text, paragraph);
            text = AppendNativeHeaderFooterFormControlText(text, paragraph);
            return AppendNativeHeaderFooterRepeatingSectionText(text, paragraph);
        }

        private static string? AppendNativeHeaderFooterEquationText(string? text, WordParagraph paragraph) {
            string? equationText = GetNativeEquationText(paragraph);
            if (string.IsNullOrWhiteSpace(equationText)) {
                return text;
            }

            var builder = new StringBuilder(text ?? string.Empty);
            string currentText = builder.ToString();
            AppendNativeHeaderFooterSupplementalValue(builder, ref currentText, equationText, skipIfPresent: true);
            return builder.Length == 0 ? text : builder.ToString();
        }

        private static string? AppendNativeHeaderFooterFormControlText(string? text, WordParagraph paragraph) {
            IReadOnlyList<W.SdtRun> checkBoxes = GetNativeCheckBoxControls(paragraph);
            IReadOnlyList<W.SdtRun> formFields = GetNativeFormFieldControls(paragraph);
            if (checkBoxes.Count == 0 && formFields.Count == 0) {
                return text;
            }

            var builder = new StringBuilder(text ?? string.Empty);
            string currentText = builder.ToString();
            foreach (W.SdtRun checkBox in checkBoxes) {
                AppendNativeHeaderFooterSupplementalValue(
                    builder,
                    ref currentText,
                    IsNativeCheckBoxChecked(checkBox) ? "[x]" : "[ ]",
                    skipIfPresent: false);
            }

            foreach (W.SdtRun formField in formFields) {
                string? value;
                if (IsNativeDatePickerControl(formField)) {
                    value = GetNativeDatePickerValue(formField);
                } else {
                    IReadOnlyList<string> options = GetNativeChoiceFieldOptions(formField);
                    value = GetNativeChoiceFieldValue(formField, options);
                }

                AppendNativeHeaderFooterSupplementalValue(builder, ref currentText, value, skipIfPresent: true);
            }

            return builder.Length == 0 ? text : builder.ToString();
        }

        private static string? AppendNativeHeaderFooterRepeatingSectionText(string? text, WordParagraph paragraph) {
            IReadOnlyList<W.SdtRun> controls = GetNativeRepeatingSectionControls(paragraph);
            if (controls.Count == 0) {
                return text;
            }

            var builder = new StringBuilder(text ?? string.Empty);
            string currentText = builder.ToString();
            foreach (W.SdtRun control in controls) {
                foreach (string itemText in GetNativeRepeatingSectionItems(control)) {
                    AppendNativeHeaderFooterSupplementalValue(builder, ref currentText, itemText, skipIfPresent: true);
                }
            }

            return builder.Length == 0 ? text : builder.ToString();
        }

        private static void AppendNativeHeaderFooterSupplementalValue(StringBuilder builder, ref string currentText, string? value, bool skipIfPresent) {
            if (string.IsNullOrWhiteSpace(value) ||
                skipIfPresent && currentText.IndexOf(value!, StringComparison.Ordinal) >= 0) {
                return;
            }

            if (builder.Length > 0 && !char.IsWhiteSpace(builder[builder.Length - 1])) {
                builder.Append(' ');
            }

            builder.Append(value);
            currentText = builder.ToString();
        }

        private static string AppendNativeTextWithEquation(string text, WordParagraph paragraph) {
            string? equationText = GetNativeEquationText(paragraph);
            if (string.IsNullOrWhiteSpace(equationText) ||
                text.IndexOf(equationText!, StringComparison.Ordinal) >= 0) {
                return text;
            }

            if (string.IsNullOrEmpty(text)) {
                return equationText!;
            }

            return char.IsWhiteSpace(text[text.Length - 1])
                ? text + equationText
                : text + " " + equationText;
        }

        private static string? GetNativeEquationText(WordParagraph paragraph) {
            var parts = new List<string>();
            AddNativeEquationText(parts, paragraph._officeMath);
            AddNativeEquationText(parts, paragraph._mathParagraph);
            if (parts.Count == 0) {
                AddNativeParagraphEquationText(parts, paragraph._paragraph);
            }

            string text = string.Concat(parts);
            return string.IsNullOrWhiteSpace(text) ? null : text;
        }

        private static void AddNativeParagraphEquationText(List<string> parts, W.Paragraph? paragraph) {
            if (paragraph == null ||
                (!paragraph.Descendants<DocumentFormat.OpenXml.Math.OfficeMath>().Any() &&
                 !paragraph.Descendants<DocumentFormat.OpenXml.Math.Paragraph>().Any())) {
                return;
            }

            int startCount = parts.Count;
            foreach (DocumentFormat.OpenXml.Math.Text text in paragraph.Descendants<DocumentFormat.OpenXml.Math.Text>()) {
                if (!string.IsNullOrEmpty(text.Text)) {
                    parts.Add(text.Text);
                }
            }

            if (parts.Count > startCount) {
                return;
            }

            foreach (DocumentFormat.OpenXml.Math.OfficeMath officeMath in paragraph.Descendants<DocumentFormat.OpenXml.Math.OfficeMath>()) {
                AddNativeEquationText(parts, officeMath);
                if (parts.Count > startCount) {
                    return;
                }
            }

            foreach (DocumentFormat.OpenXml.Math.Paragraph mathParagraph in paragraph.Descendants<DocumentFormat.OpenXml.Math.Paragraph>()) {
                AddNativeEquationText(parts, mathParagraph);
                if (parts.Count > startCount) {
                    return;
                }
            }
        }

        private static void AddNativeEquationText(List<string> parts, DocumentFormat.OpenXml.OpenXmlElement? equationElement) {
            if (equationElement == null) {
                return;
            }

            int startCount = parts.Count;
            foreach (DocumentFormat.OpenXml.Math.Text text in equationElement.Descendants<DocumentFormat.OpenXml.Math.Text>()) {
                if (!string.IsNullOrEmpty(text.Text)) {
                    parts.Add(text.Text);
                }
            }

            if (parts.Count > startCount) {
                return;
            }

            AddNativeEquationXmlText(parts, equationElement.OuterXml);
            if (parts.Count > startCount) {
                return;
            }

            AddNativeEquationXmlText(parts, equationElement.InnerXml);
            if (parts.Count == startCount && !string.IsNullOrWhiteSpace(equationElement.InnerText) && equationElement.InnerText.IndexOf('<') < 0) {
                parts.Add(equationElement.InnerText);
            }
        }

        private static void AddNativeEquationXmlText(List<string> parts, string? xml) {
            if (string.IsNullOrWhiteSpace(xml)) {
                return;
            }

            try {
                System.Xml.Linq.XElement root = System.Xml.Linq.XElement.Parse(xml!);
                foreach (System.Xml.Linq.XElement textElement in root.Descendants().Where(element =>
                    string.Equals(element.Name.LocalName, "t", StringComparison.Ordinal) &&
                    element.Name.NamespaceName.IndexOf("officeDocument/2006/math", StringComparison.OrdinalIgnoreCase) >= 0)) {
                    if (!string.IsNullOrEmpty(textElement.Value)) {
                        parts.Add(textElement.Value);
                    }
                }
            } catch (System.Xml.XmlException) {
                // Some legacy equation wrappers may expose typed math nodes without valid standalone XML.
            }
        }

        private static NativeHeaderFooterZone MapNativeTextBoxHeaderFooterZone(WordHorizontalAlignmentValues alignment) {
            switch (alignment) {
                case WordHorizontalAlignmentValues.Center:
                    return NativeHeaderFooterZone.Center;
                case WordHorizontalAlignmentValues.Right:
                case WordHorizontalAlignmentValues.Outside:
                    return NativeHeaderFooterZone.Right;
                default:
                    return NativeHeaderFooterZone.Left;
            }
        }

        private static bool TryBuildNativeHeaderFooterParagraphText(WordParagraph paragraph, out string? text, out PdfCore.PdfPageNumberStyle? pageNumberStyle) {
            text = null;
            pageNumberStyle = null;
            if (paragraph._paragraph == null) {
                return false;
            }

            var builder = new StringBuilder();
            var state = new NativeHeaderFooterFieldState();
            bool hasFieldToken = false;
            bool hasConflictingStyles = false;
            foreach (var element in paragraph._paragraph.ChildElements) {
                AppendNativeHeaderFooterElementText(element, builder, state, ref pageNumberStyle, ref hasConflictingStyles, ref hasFieldToken);
            }

            if (!hasFieldToken) {
                pageNumberStyle = null;
                return false;
            }

            text = builder.ToString();
            return !string.IsNullOrWhiteSpace(text);
        }

        private static void AppendNativeHeaderFooterElementText(DocumentFormat.OpenXml.OpenXmlElement element, StringBuilder builder, NativeHeaderFooterFieldState state, ref PdfCore.PdfPageNumberStyle? pageNumberStyle, ref bool hasConflictingStyles, ref bool hasFieldToken) {
            if (element is W.Run run) {
                AppendNativeHeaderFooterRunText(run, builder, state, ref pageNumberStyle, ref hasConflictingStyles, ref hasFieldToken);
                return;
            }

            if (element is W.Hyperlink hyperlink) {
                foreach (W.Run childRun in hyperlink.Elements<W.Run>()) {
                    AppendNativeHeaderFooterRunText(childRun, builder, state, ref pageNumberStyle, ref hasConflictingStyles, ref hasFieldToken);
                }

                return;
            }

            if (element is W.SdtRun sdtRun) {
                foreach (var child in sdtRun.SdtContentRun?.ChildElements ?? Enumerable.Empty<DocumentFormat.OpenXml.OpenXmlElement>()) {
                    AppendNativeHeaderFooterElementText(child, builder, state, ref pageNumberStyle, ref hasConflictingStyles, ref hasFieldToken);
                }

                return;
            }

            if (element is W.SimpleField simpleField) {
                string fieldCode = simpleField.Instruction?.Value ?? string.Empty;
                if (TryGetNativeHeaderFooterFieldToken(fieldCode, out string? token, out PdfCore.PdfPageNumberStyle? style)) {
                    builder.Append(token);
                    MergeNativeHeaderFooterPageNumberStyle(ref pageNumberStyle, ref hasConflictingStyles, style);
                    hasFieldToken = true;
                    return;
                }

                foreach (var child in simpleField.ChildElements) {
                    AppendNativeHeaderFooterElementText(child, builder, state, ref pageNumberStyle, ref hasConflictingStyles, ref hasFieldToken);
                }
            }
        }

        private static void AppendNativeHeaderFooterRunText(W.Run run, StringBuilder builder, NativeHeaderFooterFieldState state, ref PdfCore.PdfPageNumberStyle? pageNumberStyle, ref bool hasConflictingStyles, ref bool hasFieldToken) {
            foreach (var child in run.ChildElements) {
                if (child is W.FieldChar fieldChar) {
                    W.FieldCharValues? fieldCharType = fieldChar.FieldCharType?.Value;
                    if (fieldCharType == W.FieldCharValues.Begin) {
                        state.CollectingFieldCode = true;
                        state.SkippingFieldResult = false;
                        state.FieldCode.Clear();
                    } else if (fieldCharType == W.FieldCharValues.Separate) {
                        if (TryGetNativeHeaderFooterFieldToken(state.FieldCode.ToString(), out string? token, out PdfCore.PdfPageNumberStyle? style)) {
                            builder.Append(token);
                            MergeNativeHeaderFooterPageNumberStyle(ref pageNumberStyle, ref hasConflictingStyles, style);
                            hasFieldToken = true;
                            state.SkippingFieldResult = true;
                        }

                        state.CollectingFieldCode = false;
                    } else if (fieldCharType == W.FieldCharValues.End) {
                        state.CollectingFieldCode = false;
                        state.SkippingFieldResult = false;
                        state.FieldCode.Clear();
                    }

                    continue;
                }

                if (child is W.FieldCode fieldCode) {
                    if (state.CollectingFieldCode) {
                        state.FieldCode.Append(fieldCode.Text);
                    }

                    continue;
                }

                if (state.CollectingFieldCode || state.SkippingFieldResult) {
                    continue;
                }

                if (child is W.Text text) {
                    builder.Append(text.Text);
                } else if (child is W.TabChar) {
                    builder.Append('\t');
                } else if (child is W.Break) {
                    builder.AppendLine();
                }
            }
        }

        private static bool TryGetNativeHeaderFooterFieldToken(WordParagraph paragraph, out string? token, out PdfCore.PdfPageNumberStyle? style) {
            token = null;
            style = null;
            WordField? field = paragraph.Field;
            if (field?.FieldType == WordFieldType.Page) {
                token = "{page}";
                style = MapNativePageNumberFieldStyle(field.Field);
                return true;
            }

            if (field?.FieldType == WordFieldType.NumPages) {
                token = "{documentpages}";
                style = MapNativePageNumberFieldStyle(field.Field);
                return true;
            }

            return false;
        }

        private static bool TryGetNativeHeaderFooterFieldToken(string fieldCode, out string? token, out PdfCore.PdfPageNumberStyle? style) {
            token = null;
            style = null;
            string trimmed = fieldCode.Trim();
            if (trimmed.Length == 0) {
                return false;
            }

            int end = 0;
            while (end < trimmed.Length && !char.IsWhiteSpace(trimmed[end])) {
                end++;
            }

            string fieldType = trimmed.Substring(0, end);
            if (string.Equals(fieldType, "PAGE", StringComparison.OrdinalIgnoreCase)) {
                token = "{page}";
                style = MapNativePageNumberFieldStyle(trimmed);
                return true;
            }

            if (string.Equals(fieldType, "NUMPAGES", StringComparison.OrdinalIgnoreCase)) {
                token = "{documentpages}";
                style = MapNativePageNumberFieldStyle(trimmed);
                return true;
            }

            return false;
        }

        private static PdfCore.PdfPageNumberStyle? MapNativePageNumberFieldStyle(string fieldCode) {
            string? format = GetNativePageNumberFieldFormatSwitch(fieldCode);
            if (format == "roman") {
                return PdfCore.PdfPageNumberStyle.LowerRoman;
            }

            if (format == "Roman") {
                return PdfCore.PdfPageNumberStyle.UpperRoman;
            }

            if (format == "Alphabetical") {
                return PdfCore.PdfPageNumberStyle.LowerLetter;
            }

            if (format == "ALPHABETICAL") {
                return PdfCore.PdfPageNumberStyle.UpperLetter;
            }

            if (format == "Arabic") {
                return PdfCore.PdfPageNumberStyle.Arabic;
            }

            return null;
        }

        private static string? GetNativePageNumberFieldFormatSwitch(string fieldCode) {
            int markerIndex = fieldCode.IndexOf(@"\*", StringComparison.Ordinal);
            while (markerIndex >= 0) {
                int index = markerIndex + 2;
                while (index < fieldCode.Length && char.IsWhiteSpace(fieldCode[index])) {
                    index++;
                }

                int start = index;
                while (index < fieldCode.Length && (char.IsLetter(fieldCode[index]) || fieldCode[index] == '_')) {
                    index++;
                }

                if (index > start) {
                    return fieldCode.Substring(start, index - start);
                }

                markerIndex = fieldCode.IndexOf(@"\*", markerIndex + 2, StringComparison.Ordinal);
            }

            return null;
        }

        private static void AddNativeHeaderFooterTableText(NativeHeaderFooterText parts, WordTable table) {
            foreach (WordTableRow row in table.Rows) {
                IReadOnlyList<WordTableCell> cells = row.Cells;
                if (cells.Count == 1) {
                    AddNativeHeaderFooterSingleCellText(parts, cells[0]);
                    continue;
                }

                for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                    NativeHeaderFooterZone zone = cellIndex == 0
                        ? NativeHeaderFooterZone.Left
                        : cellIndex == cells.Count - 1
                            ? NativeHeaderFooterZone.Right
                            : NativeHeaderFooterZone.Center;

                    AddNativeHeaderFooterCellText(parts, cells[cellIndex], zone);
                }
            }
        }

        private static void AddNativeHeaderFooterSingleCellText(NativeHeaderFooterText parts, WordTableCell cell) {
            foreach (WordParagraph paragraph in GetNativeCellParagraphs(cell)) {
                AddNativeHeaderFooterParagraphText(parts, paragraph);
            }
        }

        private static void AddNativeHeaderFooterCellText(NativeHeaderFooterText parts, WordTableCell cell, NativeHeaderFooterZone zone) {
            string cellText = GetNativeHeaderFooterCellText(cell, out PdfCore.PdfPageNumberStyle? pageNumberStyle);
            if (!string.IsNullOrWhiteSpace(cellText)) {
                parts.Append(zone, cellText, pageNumberStyle);
            }
        }

        private static string GetNativeHeaderFooterCellText(WordTableCell cell, out PdfCore.PdfPageNumberStyle? pageNumberStyle) {
            var parts = new List<string>();
            pageNumberStyle = null;
            bool hasConflictingStyles = false;
            foreach (WordParagraph paragraph in GetNativeCellParagraphs(cell)) {
                string? text = GetNativeHeaderFooterParagraphText(paragraph, out PdfCore.PdfPageNumberStyle? paragraphStyle);
                if (!string.IsNullOrEmpty(text)) {
                    parts.Add(text!);
                    MergeNativeHeaderFooterPageNumberStyle(ref pageNumberStyle, ref hasConflictingStyles, paragraphStyle);
                }
            }

            return string.Join(Environment.NewLine, parts);
        }

        private static void MergeNativeHeaderFooterPageNumberStyle(ref PdfCore.PdfPageNumberStyle? current, ref bool hasConflict, PdfCore.PdfPageNumberStyle? candidate) {
            if (!candidate.HasValue || hasConflict) {
                return;
            }

            if (current.HasValue && current.Value != candidate.Value) {
                current = null;
                hasConflict = true;
                return;
            }

            current = candidate.Value;
        }

        private enum NativeHeaderFooterZone {
            Left,
            Center,
            Right
        }

        private sealed class NativeHeaderFooterFieldState {
            public bool CollectingFieldCode { get; set; }
            public bool SkippingFieldResult { get; set; }
            public StringBuilder FieldCode { get; } = new StringBuilder();
        }

        private sealed class NativeHeaderFooterImage {
            public NativeHeaderFooterImage(byte[] data, double width, double height, PdfCore.PdfAlign align) {
                Data = data;
                Width = width;
                Height = height;
                Align = align;
            }

            public byte[] Data { get; }
            public double Width { get; }
            public double Height { get; }
            public PdfCore.PdfAlign Align { get; }
        }

        private sealed class NativeHeaderFooterShape {
            public NativeHeaderFooterShape(OfficeShape shape, PdfCore.PdfAlign align) {
                Shape = shape.Clone();
                Align = align;
            }

            public OfficeShape Shape { get; }
            public PdfCore.PdfAlign Align { get; }
        }

        private sealed class NativeHeaderFooterText {
            public string? Left { get; private set; }
            public string? Center { get; private set; }
            public string? Right { get; private set; }
            public bool HasPageTokens { get; private set; }
            public PdfCore.PdfPageNumberStyle? PageNumberStyle { get; private set; }
            private bool _hasConflictingPageNumberStyles;
            public bool HasContent =>
                !string.IsNullOrWhiteSpace(Left) ||
                !string.IsNullOrWhiteSpace(Center) ||
                !string.IsNullOrWhiteSpace(Right);

            public void AppendLeft(string text) => Left = Append(Left, text, null);
            public void AppendCenter(string text) => Center = Append(Center, text, null);
            public void AppendRight(string text) => Right = Append(Right, text, null);
            public void AppendLeft(string text, PdfCore.PdfPageNumberStyle? pageNumberStyle) => Left = Append(Left, text, pageNumberStyle);
            public void AppendCenter(string text, PdfCore.PdfPageNumberStyle? pageNumberStyle) => Center = Append(Center, text, pageNumberStyle);
            public void AppendRight(string text, PdfCore.PdfPageNumberStyle? pageNumberStyle) => Right = Append(Right, text, pageNumberStyle);

            public void Append(NativeHeaderFooterZone zone, string text) => Append(zone, text, null);

            public void Append(NativeHeaderFooterZone zone, string text, PdfCore.PdfPageNumberStyle? pageNumberStyle) {
                switch (zone) {
                    case NativeHeaderFooterZone.Center:
                        AppendCenter(text, pageNumberStyle);
                        break;
                    case NativeHeaderFooterZone.Right:
                        AppendRight(text, pageNumberStyle);
                        break;
                    default:
                        AppendLeft(text, pageNumberStyle);
                        break;
                }
            }

            public NativeHeaderFooterText Clone() {
                return new NativeHeaderFooterText {
                    Left = Left,
                    Center = Center,
                    Right = Right,
                    HasPageTokens = HasPageTokens,
                    PageNumberStyle = PageNumberStyle,
                    _hasConflictingPageNumberStyles = _hasConflictingPageNumberStyles
                };
            }

            private string Append(string? current, string text, PdfCore.PdfPageNumberStyle? pageNumberStyle) {
                if (text.IndexOf("{page}", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    text.IndexOf("{pages}", StringComparison.OrdinalIgnoreCase) >= 0) {
                    HasPageTokens = true;
                }

                RecordPageNumberStyle(pageNumberStyle);
                return string.IsNullOrWhiteSpace(current) ? text : current + " " + text;
            }

            private void RecordPageNumberStyle(PdfCore.PdfPageNumberStyle? style) {
                if (!style.HasValue || _hasConflictingPageNumberStyles) {
                    return;
                }

                if (PageNumberStyle.HasValue && PageNumberStyle.Value != style.Value) {
                    PageNumberStyle = null;
                    _hasConflictingPageNumberStyles = true;
                    return;
                }

                PageNumberStyle = style.Value;
            }
        }

        private static PdfCore.PdfOptions CreateNativeOptions(WordDocument document, PdfSaveOptions? options) {
            WordSection? firstSection = document.Sections.FirstOrDefault();
            PdfCore.PdfStandardFont defaultFont = GetNativeDefaultFont(document, options);
            return new PdfCore.PdfOptions {
                PageSize = firstSection == null ? PdfCore.PageSizes.A4 : GetNativePageSize(firstSection, options),
                Margins = firstSection == null ? PdfCore.PageMargins.Uniform(72) : GetNativeMargins(firstSection, options),
                DefaultFont = defaultFont,
                HeaderFont = defaultFont,
                FooterFont = defaultFont,
                BackgroundColor = ParseNativeColor(document.Background?.Color),
                CreateOutlineFromHeadings = true
            };
        }

        private static PdfCore.PdfStandardFont GetNativeDefaultFont(WordDocument document, PdfSaveOptions? options) {
            if (TryMapNativeFontFamily(options?.FontFamily, out PdfCore.PdfStandardFont optionFont)) {
                return optionFont;
            }

            if (TryMapNativeFontFamily(document.Settings.FontFamily, out PdfCore.PdfStandardFont settingsFont) ||
                TryMapNativeFontFamily(document.Settings.FontFamilyHighAnsi, out settingsFont) ||
                TryMapNativeFontFamily(document.Settings.FontFamilyEastAsia, out settingsFont) ||
                TryMapNativeFontFamily(document.Settings.FontFamilyComplexScript, out settingsFont)) {
                return settingsFont;
            }

            return PdfCore.PdfStandardFont.Helvetica;
        }

        private static bool TryMapNativeFontFamily(string? fontFamily, out PdfCore.PdfStandardFont font) {
            font = PdfCore.PdfStandardFont.Helvetica;
            if (string.IsNullOrWhiteSpace(fontFamily)) {
                return false;
            }

            string normalized = NormalizeNativeFontFamily(fontFamily!);
            switch (normalized) {
                case "timesnewroman":
                case "times":
                case "timesroman":
                case "georgia":
                case "cambria":
                case "serif":
                    font = PdfCore.PdfStandardFont.TimesRoman;
                    return true;
                case "couriernew":
                case "courier":
                case "consolas":
                case "lucidaconsole":
                case "monospace":
                    font = PdfCore.PdfStandardFont.Courier;
                    return true;
                case "arial":
                case "helvetica":
                case "calibri":
                case "aptos":
                case "segoeui":
                case "tahoma":
                case "verdana":
                case "sans":
                case "sansserif":
                    font = PdfCore.PdfStandardFont.Helvetica;
                    return true;
                default:
                    return false;
            }
        }

        private static string NormalizeNativeFontFamily(string fontFamily) {
            string firstFamily = fontFamily.Split(new[] { ',', ';' }, 2)[0];
            var builder = new StringBuilder(firstFamily.Length);
            foreach (char ch in firstFamily) {
                if (char.IsLetterOrDigit(ch)) {
                    builder.Append(char.ToLowerInvariant(ch));
                }
            }

            return builder.ToString();
        }

        private sealed class NativeTableOfContentsEntry {
            public NativeTableOfContentsEntry(string text, int level, int pageNumber, string? destinationName) {
                Text = text;
                Level = level;
                PageNumber = pageNumber;
                DestinationName = destinationName;
            }

            public string Text { get; }
            public int Level { get; }
            public int PageNumber { get; }
            public string? DestinationName { get; }
        }

        private static Dictionary<W.Paragraph, string> BuildNativeHeadingDestinations(WordDocument document) {
            var destinations = new Dictionary<W.Paragraph, string>();
            var used = new HashSet<string>(StringComparer.Ordinal);
            int headingIndex = 0;

            foreach (WordSection section in document.Sections) {
                foreach (WordElement element in CollapseNativeParagraphElements(section.Elements)) {
                    if (element is not WordParagraph paragraph ||
                        paragraph._paragraph == null ||
                        GetNativeTableOfContentsHeadingLevel(paragraph) <= 0) {
                        continue;
                    }

                    string headingText = GetNativeParagraphDisplayText(paragraph);
                    if (string.IsNullOrWhiteSpace(headingText)) {
                        continue;
                    }

                    string? bookmarkName = string.IsNullOrWhiteSpace(paragraph.Bookmark?.Name)
                        ? null
                        : paragraph.Bookmark!.Name;
                    string destinationName = bookmarkName ?? CreateNativeHeadingDestinationName(headingText, ++headingIndex, used);
                    destinations[paragraph._paragraph] = destinationName;
                    used.Add(destinationName);
                }
            }

            return destinations;
        }

        private static string CreateNativeHeadingDestinationName(string text, int headingIndex, HashSet<string> used) {
            var builder = new StringBuilder("officeimo-heading-");
            foreach (char ch in text) {
                if (char.IsLetterOrDigit(ch)) {
                    builder.Append(char.ToLowerInvariant(ch));
                } else if (builder[builder.Length - 1] != '-') {
                    builder.Append('-');
                }

                if (builder.Length >= 80) {
                    break;
                }
            }

            string baseName = builder.ToString().TrimEnd('-');
            if (baseName.Length <= "officeimo-heading".Length) {
                baseName = "officeimo-heading-" + headingIndex.ToString(CultureInfo.InvariantCulture);
            }

            string name = baseName;
            int suffix = 2;
            while (used.Contains(name)) {
                name = baseName + "-" + suffix.ToString(CultureInfo.InvariantCulture);
                suffix++;
            }

            return name;
        }

        private static IReadOnlyList<NativeTableOfContentsEntry> BuildNativeTableOfContentsEntries(WordDocument document, PdfSaveOptions? options, IReadOnlyDictionary<W.Paragraph, string> headingDestinations) {
            var entries = new List<NativeTableOfContentsEntry>();
            int headingCount = CountNativeDocumentHeadings(document);
            int currentPage = 1;
            double consumedOnPage = 0D;
            bool firstSection = true;

            foreach (WordSection section in document.Sections) {
                if (!firstSection) {
                    currentPage++;
                    consumedOnPage = 0D;
                }

                firstSection = false;
                PdfCore.PageSize pageSize = GetNativePageSize(section, options);
                PdfCore.PageMargins margins = GetNativeMargins(section, options);
                double contentHeight = Math.Max(72D, pageSize.Height - margins.Top - margins.Bottom);
                double contentWidth = Math.Max(72D, pageSize.Width - margins.Left - margins.Right);

                foreach (WordElement element in CollapseNativeParagraphElements(section.Elements)) {
                    if (element is WordParagraph paragraph && paragraph.PageBreakBefore) {
                        currentPage++;
                        consumedOnPage = 0D;
                    }

                    if (element is WordParagraph pageBreakParagraph && pageBreakParagraph.IsPageBreak) {
                        currentPage++;
                        consumedOnPage = 0D;
                        continue;
                    }

                    if (element is WordBreak wordBreak && wordBreak.BreakType == W.BreakValues.Page) {
                        currentPage++;
                        consumedOnPage = 0D;
                        continue;
                    }

                    double estimatedHeight = EstimateNativeElementHeight(element, contentWidth, headingCount);
                    if (estimatedHeight <= 0D) {
                        continue;
                    }

                    if (consumedOnPage > 0D && consumedOnPage + estimatedHeight > contentHeight) {
                        currentPage++;
                        consumedOnPage = 0D;
                    }

                    if (element is WordParagraph headingParagraph) {
                        int headingLevel = GetNativeTableOfContentsHeadingLevel(headingParagraph);
                        if (headingLevel > 0) {
                            string headingText = GetNativeParagraphDisplayText(headingParagraph);
                            if (!string.IsNullOrWhiteSpace(headingText)) {
                                string? destinationName = headingParagraph._paragraph != null &&
                                    headingDestinations.TryGetValue(headingParagraph._paragraph, out string? foundDestination)
                                        ? foundDestination
                                        : null;
                                entries.Add(new NativeTableOfContentsEntry(headingText, headingLevel, currentPage, destinationName));
                            }
                        }
                    }

                    consumedOnPage += estimatedHeight;
                    while (consumedOnPage > contentHeight) {
                        currentPage++;
                        consumedOnPage -= contentHeight;
                    }
                }
            }

            return entries;
        }

        private static int CountNativeDocumentHeadings(WordDocument document) {
            int count = 0;
            foreach (WordSection section in document.Sections) {
                foreach (WordElement element in CollapseNativeParagraphElements(section.Elements)) {
                    if (element is WordParagraph paragraph &&
                        GetNativeTableOfContentsHeadingLevel(paragraph) > 0 &&
                        !string.IsNullOrWhiteSpace(GetNativeParagraphDisplayText(paragraph))) {
                        count++;
                    }
                }
            }

            return count;
        }

        private static double EstimateNativeElementHeight(WordElement element, double contentWidth, int headingCount) {
            switch (element) {
                case WordTableOfContent:
                    return 18D + Math.Max(1, headingCount) * 15D + 10D;
                case WordTable table:
                    return EstimateNativeTableHeight(table, contentWidth);
                case WordImage image:
                    return image.Height.HasValue ? image.Height.Value * 72D / 96D + 6D : 150D;
                case WordParagraph paragraph:
                    return EstimateNativeParagraphHeight(paragraph, contentWidth);
                default:
                    return 0D;
            }
        }

        private static double EstimateNativeTableHeight(WordTable table, double contentWidth) {
            int rowCount = Math.Max(1, table.Rows.Count);
            int columnCount = Math.Max(1, table.Rows.Select(row => row.Cells.Count).DefaultIfEmpty(1).Max());
            double cellWidth = Math.Max(48D, contentWidth / columnCount);
            double height = 0D;
            foreach (WordTableRow row in table.Rows) {
                int rowLines = 1;
                foreach (WordTableCell cell in row.Cells) {
                    string cellText = GetNativeCellText(cell);
                    rowLines = Math.Max(rowLines, EstimateNativeLineCount(cellText, cellWidth, 10D));
                }

                height += rowLines * 14D + 12D;
            }

            return Math.Max(rowCount * 22D, height) + 6D;
        }

        private static double EstimateNativeParagraphHeight(WordParagraph paragraph, double contentWidth) {
            if (paragraph.IsPageBreak) {
                return 0D;
            }

            string text = GetNativeParagraphDisplayText(paragraph);
            if (string.IsNullOrWhiteSpace(text) &&
                paragraph.Image == null &&
                paragraph.Shape == null &&
                paragraph.PictureControl?.Image == null) {
                return 0D;
            }

            int headingLevel = GetNativeTableOfContentsHeadingLevel(paragraph);
            if (headingLevel > 0) {
                double headingSize = headingLevel == 1 ? 18D : headingLevel == 2 ? 15D : 13D;
                return EstimateNativeLineCount(text, contentWidth, headingSize) * headingSize * 1.25D + 8D;
            }

            double fontSize = paragraph.FontSize.HasValue && paragraph.FontSize.Value > 0 ? paragraph.FontSize.Value : 11D;
            double height = EstimateNativeLineCount(text, contentWidth, fontSize) * fontSize * NativeDefaultParagraphLineHeight + NativeDefaultParagraphSpacingAfter;
            if (!string.IsNullOrWhiteSpace(paragraph.ShadingFillColorHex) ||
                HasNativeBorder(paragraph.Borders.TopStyle) ||
                HasNativeBorder(paragraph.Borders.BottomStyle) ||
                HasNativeBorder(paragraph.Borders.LeftStyle) ||
                HasNativeBorder(paragraph.Borders.RightStyle)) {
                height += 8D;
            }

            return height;
        }

        private static int EstimateNativeLineCount(string? text, double contentWidth, double fontSize) {
            if (string.IsNullOrEmpty(text)) {
                return 1;
            }

            double averageCharacterWidth = Math.Max(3D, fontSize * 0.48D);
            int charactersPerLine = Math.Max(12, (int)Math.Floor(contentWidth / averageCharacterWidth));
            int lines = 0;
            foreach (string part in text!.Replace("\r\n", "\n").Split('\n')) {
                lines += Math.Max(1, (int)Math.Ceiling(part.Length / (double)charactersPerLine));
            }

            return Math.Max(1, lines);
        }

        private static string GetNativeParagraphDisplayText(WordParagraph paragraph) {
            if (paragraph.IsHyperLink && paragraph.Hyperlink != null) {
                return paragraph.Hyperlink.Text;
            }

            List<WordParagraph> runs = GetNativeRuns(paragraph);
            string text = runs.Count > 0
                ? string.Concat(runs.Where(run => !run.IsImage).Select(run => run.Text))
                : paragraph.Text;
            return AppendNativeTextWithEquation(text, paragraph);
        }

        private static int GetNativeTableOfContentsHeadingLevel(WordParagraph paragraph) {
            if (!paragraph.Style.HasValue) {
                return 0;
            }

            return paragraph.Style.Value switch {
                WordParagraphStyles.Heading1 => 1,
                WordParagraphStyles.Heading2 => 2,
                WordParagraphStyles.Heading3 => 3,
                WordParagraphStyles.Heading4 => 4,
                WordParagraphStyles.Heading5 => 5,
                WordParagraphStyles.Heading6 => 6,
                WordParagraphStyles.Heading7 => 7,
                WordParagraphStyles.Heading8 => 8,
                WordParagraphStyles.Heading9 => 9,
                _ => 0
            };
        }

        private static void RenderNativeTableOfContents(INativePdfFlow pdf, WordTableOfContent tableOfContent, IReadOnlyList<NativeTableOfContentsEntry> entries) {
            string title = string.IsNullOrWhiteSpace(tableOfContent.Text) ? "Table of Contents" : tableOfContent.Text;
            pdf.Paragraph(builder => builder.FontSize(11D).Text(title), PdfCore.PdfAlign.Left, null, new PdfCore.PdfParagraphStyle {
                SpacingAfter = 5D,
                KeepWithNext = true
            });

            int minLevel = tableOfContent.MinLevel;
            int maxLevel = tableOfContent.MaxLevel;
            int rendered = 0;
            foreach (NativeTableOfContentsEntry entry in entries) {
                if (entry.Level < minLevel || entry.Level > maxLevel) {
                    continue;
                }

                int relativeLevel = Math.Max(0, entry.Level - minLevel);
                var style = new PdfCore.PdfParagraphStyle {
                    LeftIndent = relativeLevel * 14D,
                    SpacingAfter = 1D,
                    DefaultTabStopWidth = 432D,
                    KeepWithNext = true
                };
                pdf.Paragraph(
                    builder => {
                        builder.FontSize(10.5D);
                        if (string.IsNullOrEmpty(entry.DestinationName)) {
                            builder.Text(entry.Text);
                        } else {
                            builder.LinkToBookmark(entry.Text, entry.DestinationName!, underline: false, contents: "Table of contents: " + entry.Text);
                        }

                        builder
                            .Tab(PdfCore.PdfTabLeaderStyle.Dots, PdfCore.PdfTabAlignment.Right)
                            .Text(entry.PageNumber.ToString(CultureInfo.InvariantCulture));
                    },
                    PdfCore.PdfAlign.Left,
                    null,
                    style);
                rendered++;
            }

            if (rendered == 0) {
                string fallback = string.IsNullOrWhiteSpace(tableOfContent.TextNoContent)
                    ? "No table of contents entries found."
                    : tableOfContent.TextNoContent;
                pdf.Paragraph(builder => builder.FontSize(10.5D).Text(fallback));
            }
        }

        private static void RenderNativeElement(INativePdfFlow pdf, WordElement element, Func<WordParagraph, (int Level, string Marker)?> getMarker, IReadOnlyList<int> footnoteNumbers, Dictionary<long, int> footnoteNumbersById, PdfSaveOptions? options, IReadOnlyList<NativeTableOfContentsEntry> tableOfContentsEntries, IReadOnlyDictionary<W.Paragraph, string> headingDestinations) {
            switch (element) {
                case WordParagraph paragraph:
                    RenderNativeParagraph(pdf, paragraph, getMarker(paragraph), footnoteNumbers, footnoteNumbersById, options, headingDestinations);
                    break;
                case WordTableOfContent tableOfContent:
                    RenderNativeTableOfContents(pdf, tableOfContent, tableOfContentsEntries);
                    break;
                case WordTable table:
                    RenderNativeTable(pdf, table, getMarker, footnoteNumbersById, options);
                    break;
                case WordImage image:
                    RenderNativeImage(pdf, image, options: options, source: "body image");
                    break;
                case WordHyperLink link:
                    RenderNativeHyperLink(pdf, link);
                    break;
                case WordBreak wordBreak:
                    RenderNativeBreak(pdf, wordBreak);
                    break;
                case WordShape shape:
                    RenderNativeShape(pdf, shape);
                    break;
                case WordEmbeddedDocument:
                    if (options != null) {
                        AddNativeExportWarning(
                            options,
                            "NativeBodyEmbeddedDocumentUnsupported",
                            "body",
                            "Embedded documents in Word body content are not mapped by the OfficeIMO PDF engine yet.");
                    }

                    break;
                default:
                    if (options != null) {
                        AddNativeExportWarning(
                            options,
                            "NativeBodyElementUnsupported",
                            "body",
                            "Word body element '" + element.GetType().Name + "' is not mapped by the OfficeIMO PDF engine yet.");
                    }

                    break;
            }
        }

        private static void RenderNativeBreak(INativePdfFlow pdf, WordBreak wordBreak) {
            if (wordBreak.BreakType == W.BreakValues.Page) {
                pdf.PageBreak();
            }
        }

        private static void RenderNativeParagraph(INativePdfFlow pdf, WordParagraph paragraph, (int Level, string Marker)? marker, IReadOnlyList<int> footnoteNumbers, Dictionary<long, int> footnoteNumbersById, PdfSaveOptions? options, IReadOnlyDictionary<W.Paragraph, string> headingDestinations) {
            if (paragraph == null) {
                return;
            }

            if (paragraph.PageBreakBefore) {
                pdf.PageBreak();
            }

            if (paragraph.IsPageBreak) {
                pdf.PageBreak();
                return;
            }

            RecordNativeBodyParagraphDiagnostics(paragraph, options, "body paragraph", mapsCheckBoxes: true, mapsFormFields: true, mapsPictureControls: true, mapsRepeatingSections: true);
            IReadOnlyList<W.SdtRun> checkboxControls = GetNativeCheckBoxControls(paragraph);
            IReadOnlyList<W.SdtRun> formFieldControls = GetNativeFormFieldControls(paragraph);
            IReadOnlyList<W.SdtRun> repeatingSectionControls = GetNativeRepeatingSectionControls(paragraph);

            if (!string.IsNullOrEmpty(paragraph.Bookmark?.Name)) {
                pdf.Bookmark(paragraph.Bookmark!.Name!);
            }

            WordTextBox? textBox = GetNativeParagraphTextBox(paragraph, out string? textBoxFallbackText);
            if (textBox != null) {
                RenderNativeTextBox(pdf, textBox, footnoteNumbersById, options, textBoxFallbackText);
                return;
            }

            if (paragraph.Shape != null) {
                RenderNativeShape(pdf, paragraph.Shape);
            }

            if (paragraph.Image != null) {
                RenderNativeImage(pdf, paragraph.Image, MapNativeParagraphAlign(paragraph.ParagraphAlignment, allowJustify: false), options, "body paragraph image");
            }

            WordImage? pictureControlImage = paragraph.PictureControl?.Image;
            if (pictureControlImage != null) {
                RenderNativeImage(pdf, pictureControlImage, MapNativeParagraphAlign(paragraph.ParagraphAlignment, allowJustify: false), options, "body picture control image");
            }

            List<WordParagraph> runs = GetNativeRuns(paragraph);
            if (paragraph.Image == null) {
                RenderNativeRunImages(pdf, runs, MapNativeParagraphAlign(paragraph.ParagraphAlignment, allowJustify: false), options);
            }

            string content = paragraph.IsHyperLink && paragraph.Hyperlink != null ? paragraph.Hyperlink.Text : AppendNativeTextWithEquation(paragraph.Text, paragraph);
            bool hasRenderableRuns = runs.Any(run => !run.IsImage && !string.IsNullOrEmpty(run.Text));
            List<int> paragraphFootnoteNumbers = GetNativeParagraphFootnoteNumbers(paragraph, runs, footnoteNumbers, footnoteNumbersById);
            PdfCore.PdfParagraphStyle style = CreateNativeParagraphStyle(paragraph);
            if (marker == null &&
                paragraphFootnoteNumbers.Count == 0 &&
                IsNativeHorizontalRuleParagraph(paragraph, runs, content) &&
                CreateNativeHorizontalRuleStyle(paragraph, style) is { } horizontalRuleStyle) {
                pdf.HR(style: horizontalRuleStyle);
                return;
            }

            if (!hasRenderableRuns && string.IsNullOrEmpty(content) && marker == null && paragraphFootnoteNumbers.Count == 0 && checkboxControls.Count == 0 && formFieldControls.Count == 0 && repeatingSectionControls.Count == 0) {
                return;
            }

            PdfCore.PdfAlign align = MapNativeParagraphAlign(paragraph.ParagraphAlignment);
            PdfCore.PdfAlign objectAlign = MapNativeParagraphAlign(paragraph.ParagraphAlignment, allowJustify: false);
            PdfCore.PdfColor? defaultColor = ParseNativeColor(paragraph.ColorHex);
            int headingLevel = GetHeadingLevel(paragraph);
            PdfCore.PdfColor? headingColor = GetNativeHeadingColor(headingLevel, defaultColor);
            (string? LinkUri, string? LinkDestinationName, string? LinkContents) headingLink = GetNativeHeadingLink(paragraph);
            bool hasHeadingLinkTarget = headingLink.LinkUri != null || headingLink.LinkDestinationName != null;
            PdfCore.PdfHorizontalRuleStyle? topBorderRuleStyle = marker == null ? CreateNativeTopBorderRuleStyle(paragraph, style) : null;
            PdfCore.PdfParagraphStyle paragraphStyle = topBorderRuleStyle == null ? style : style.Clone();
            if (topBorderRuleStyle != null) {
                paragraphStyle.SpacingBefore = 0;
                pdf.HR(style: topBorderRuleStyle);
            }

            if (headingLevel > 0 && marker == null) {
                if (paragraph._paragraph != null &&
                    string.IsNullOrEmpty(paragraph.Bookmark?.Name) &&
                    headingDestinations.TryGetValue(paragraph._paragraph, out string? generatedDestinationName)) {
                    pdf.Bookmark(generatedDestinationName);
                }

                RenderNativeHeading(pdf, headingLevel, content, objectAlign, headingColor, headingLink.LinkUri, headingLink.LinkDestinationName, headingLink.LinkContents);
                if (CreateNativeBottomBorderRuleStyle(paragraph, paragraphStyle) is { } headingRuleStyle) {
                    pdf.HR(style: headingRuleStyle);
                }

                RenderNativeFormFields(pdf, formFieldControls, objectAlign);
                RenderNativeCheckBoxes(pdf, checkboxControls, objectAlign);
                RenderNativeRepeatingSections(pdf, repeatingSectionControls, align, defaultColor);
                return;
            }

            PdfCore.PanelStyle? panelStyle = CreateNativeParagraphPanelStyle(paragraph, paragraphStyle);
            if (panelStyle != null) {
                pdf.PanelParagraph(builder => {
                    AddNativeParagraphContent(builder, paragraph, marker, runs, hasRenderableRuns, content, paragraphFootnoteNumbers, options);
                }, panelStyle, align, defaultColor);
                RenderNativeFormFields(pdf, formFieldControls, objectAlign);
                RenderNativeCheckBoxes(pdf, checkboxControls, objectAlign);
                RenderNativeRepeatingSections(pdf, repeatingSectionControls, align, defaultColor);
                return;
            }

            PdfCore.PdfHorizontalRuleStyle? bottomBorderRuleStyle = marker == null ? CreateNativeBottomBorderRuleStyle(paragraph, paragraphStyle) : null;
            if (bottomBorderRuleStyle != null && ReferenceEquals(paragraphStyle, style)) {
                paragraphStyle = style.Clone();
            }

            if (bottomBorderRuleStyle != null) {
                paragraphStyle.SpacingAfter = 0;
            }

            if (hasRenderableRuns || !string.IsNullOrEmpty(content) || marker != null || paragraphFootnoteNumbers.Count > 0) {
                pdf.Paragraph(builder => {
                    AddNativeParagraphContent(builder, paragraph, marker, runs, hasRenderableRuns, content, paragraphFootnoteNumbers, options);
                }, align, defaultColor, paragraphStyle);
            }

            if (bottomBorderRuleStyle != null) {
                pdf.HR(style: bottomBorderRuleStyle);
            }

            RenderNativeFormFields(pdf, formFieldControls, objectAlign);
            RenderNativeCheckBoxes(pdf, checkboxControls, objectAlign);
            RenderNativeRepeatingSections(pdf, repeatingSectionControls, align, defaultColor);
        }

        private static void RenderNativeFormFields(INativePdfFlow pdf, IReadOnlyList<W.SdtRun> formFieldControls, PdfCore.PdfAlign align) {
            for (int index = 0; index < formFieldControls.Count; index++) {
                W.SdtRun formField = formFieldControls[index];
                double spacingBefore = index == 0 ? 0D : 2D;
                if (IsNativeDatePickerControl(formField)) {
                    pdf.TextField(
                        GetNativeContentControlFieldName(formField, index, "WordDatePicker"),
                        width: 150D,
                        height: 20D,
                        value: GetNativeDatePickerValue(formField),
                        align: align,
                        fontSize: 10D,
                        spacingBefore: spacingBefore,
                        spacingAfter: 4D);
                    continue;
                }

                IReadOnlyList<string> options = GetNativeChoiceFieldOptions(formField);
                string? value = GetNativeChoiceFieldValue(formField, options);
                if (options.Count == 0 || string.IsNullOrWhiteSpace(value)) {
                    continue;
                }

                string fallbackPrefix = formField.SdtProperties?.Elements<W.SdtContentComboBox>().Any() == true
                    ? "WordComboBox"
                    : "WordDropDownList";
                pdf.ChoiceField(
                    GetNativeContentControlFieldName(formField, index, fallbackPrefix),
                    options,
                    value,
                    width: 150D,
                    height: 20D,
                    align: align,
                    fontSize: 10D,
                    spacingBefore: spacingBefore,
                    spacingAfter: 4D,
                    isComboBox: true);
            }
        }

        private static void RenderNativeCheckBoxes(INativePdfFlow pdf, IReadOnlyList<W.SdtRun> checkboxControls, PdfCore.PdfAlign align) {
            for (int index = 0; index < checkboxControls.Count; index++) {
                W.SdtRun checkbox = checkboxControls[index];
                pdf.CheckBox(
                    GetNativeCheckBoxFieldName(checkbox, index),
                    IsNativeCheckBoxChecked(checkbox),
                    size: 12D,
                    align: align,
                    spacingBefore: index == 0 ? 0D : 2D,
                    spacingAfter: 4D);
            }
        }

        private static void RenderNativeRepeatingSections(INativePdfFlow pdf, IReadOnlyList<W.SdtRun> repeatingSectionControls, PdfCore.PdfAlign align, PdfCore.PdfColor? color) {
            foreach (W.SdtRun repeatingSection in repeatingSectionControls) {
                foreach (string itemText in GetNativeRepeatingSectionItems(repeatingSection)) {
                    if (string.IsNullOrWhiteSpace(itemText)) {
                        continue;
                    }

                    pdf.Paragraph(builder => builder.Text(itemText), align, color);
                }
            }
        }

        private static IReadOnlyList<string> GetNativeRepeatingSectionItems(W.SdtRun repeatingSection) {
            var items = new List<string>();
            IEnumerable<DocumentFormat.OpenXml.OpenXmlElement> itemElements = repeatingSection.SdtContentRun?.ChildElements
                .Where(element => element.LocalName == "repeatingSectionItem") ??
                Enumerable.Empty<DocumentFormat.OpenXml.OpenXmlElement>();

            foreach (DocumentFormat.OpenXml.OpenXmlElement item in itemElements) {
                string text = string.Concat(item.Descendants<W.Text>().Select(value => value.Text));
                if (!string.IsNullOrWhiteSpace(text)) {
                    items.Add(text);
                }
            }

            if (items.Count == 0) {
                string text = GetNativeSdtText(repeatingSection) ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(text)) {
                    items.Add(text);
                }
            }

            return items;
        }

        private static IReadOnlyList<PdfCore.PdfTableCellCheckBox> CreateNativeTableCellCheckBoxes(WordTableCell cell) {
            var checkBoxes = new List<PdfCore.PdfTableCellCheckBox>();
            foreach (WordParagraph paragraph in GetNativeCellParagraphs(cell)) {
                IReadOnlyList<W.SdtRun> controls = GetNativeCheckBoxControls(paragraph);
                for (int index = 0; index < controls.Count; index++) {
                    W.SdtRun checkbox = controls[index];
                    checkBoxes.Add(new PdfCore.PdfTableCellCheckBox(
                        GetNativeCheckBoxFieldName(checkbox, checkBoxes.Count, "WordTableCheckBox"),
                        IsNativeCheckBoxChecked(checkbox),
                        size: 12D));
                }
            }

            return checkBoxes;
        }

        private static IReadOnlyList<PdfCore.PdfTableCellFormField> CreateNativeTableCellFormFields(WordTableCell cell) {
            var formFields = new List<PdfCore.PdfTableCellFormField>();
            foreach (WordParagraph paragraph in GetNativeCellParagraphs(cell)) {
                IReadOnlyList<W.SdtRun> controls = GetNativeFormFieldControls(paragraph);
                for (int index = 0; index < controls.Count; index++) {
                    W.SdtRun formField = controls[index];
                    if (IsNativeDatePickerControl(formField)) {
                        formFields.Add(PdfCore.PdfTableCellFormField.TextField(
                            GetNativeContentControlFieldName(formField, formFields.Count, "WordTableDatePicker"),
                            GetNativeDatePickerValue(formField),
                            width: 150D,
                            height: 20D,
                            fontSize: 10D));
                        continue;
                    }

                    IReadOnlyList<string> options = GetNativeChoiceFieldOptions(formField);
                    string? value = GetNativeChoiceFieldValue(formField, options);
                    if (options.Count == 0 || string.IsNullOrWhiteSpace(value)) {
                        continue;
                    }

                    string fallbackPrefix = formField.SdtProperties?.Elements<W.SdtContentComboBox>().Any() == true
                        ? "WordTableComboBox"
                        : "WordTableDropDownList";
                    formFields.Add(PdfCore.PdfTableCellFormField.ChoiceField(
                        GetNativeContentControlFieldName(formField, formFields.Count, fallbackPrefix),
                        options,
                        value,
                        width: 150D,
                        height: 20D,
                        fontSize: 10D,
                        isComboBox: true));
                }
            }

            return formFields;
        }

        private static IReadOnlyList<PdfCore.PdfTableCellImage> CreateNativeTableCellImages(WordTableCell cell) {
            var images = new List<PdfCore.PdfTableCellImage>();
            foreach (WordParagraph paragraph in GetNativeCellParagraphs(cell)) {
                if (paragraph.Image != null) {
                    AddNativeTableCellImage(images, paragraph.Image);
                }

                foreach (W.SdtRun pictureControl in GetNativePictureControls(paragraph)) {
                    var pictureParagraph = new WordParagraph(paragraph._document, paragraph._paragraph!, pictureControl);
                    WordImage? pictureControlImage = pictureParagraph.PictureControl?.Image;
                    if (pictureControlImage == null) {
                        continue;
                    }

                    AddNativeTableCellImage(images, pictureControlImage);
                }
            }

            return images;
        }

        private static void AddNativeTableCellImage(List<PdfCore.PdfTableCellImage> images, WordImage image) {
            byte[] bytes = ImageEmbedder.GetImageBytes(image);
            if (!IsNativePdfSupportedImageBytes(bytes, out _)) {
                return;
            }

            double width = image.Width.HasValue ? image.Width.Value * 72D / 96D : 144D;
            double height = image.Height.HasValue ? image.Height.Value * 72D / 96D : 144D;
            images.Add(new PdfCore.PdfTableCellImage(bytes, width, height));
        }

        private static void AddNativeParagraphContent(
            PdfCore.PdfParagraphBuilder builder,
            WordParagraph paragraph,
            (int Level, string Marker)? marker,
            IReadOnlyList<WordParagraph> runs,
            bool hasRenderableRuns,
            string content,
            IReadOnlyList<int> paragraphFootnoteNumbers,
            PdfSaveOptions? options) {
                if (marker != null) {
                    builder.Text(new string(' ', Math.Max(0, marker.Value.Level - 1) * 2));
                    builder.Text(marker.Value.Marker);
                    builder.Text(" ");
                }

                IReadOnlyList<WordTabStop> tabStops = paragraph.TabStops;
                int tabIndex = 0;
                if (hasRenderableRuns) {
                    foreach (WordParagraph run in runs) {
                        if (run.IsImage && run.Image != null) {
                            continue;
                        }

                        if (IsNativeTextWrappingBreak(run)) {
                            builder.LineBreak();
                            tabIndex = 0;
                            continue;
                        }

                        AddNativeRun(builder, run, paragraph, tabStops, ref tabIndex, options);
                }
                    string? supplementalText = GetNativeSupplementalTextAfterRuns(content, runs);
                    if (!string.IsNullOrEmpty(supplementalText)) {
                        AddNativeText(builder, supplementalText!, paragraph, tabStops, ref tabIndex);
                    }
            } else if (paragraph.IsHyperLink && paragraph.Hyperlink != null) {
                    ApplyNativeTextStyle(builder, paragraph);
                    AddNativeHyperLinkRun(builder, paragraph.Hyperlink.Text, paragraph.Hyperlink, tabStops, ref tabIndex);
                    ResetNativeTextStyle(builder);
                } else {
                    AddNativeText(builder, content, paragraph, tabStops, ref tabIndex);
                }

                AddNativeFootnoteReferences(builder, paragraphFootnoteNumbers);
        }

        private static void RenderNativeRunImages(INativePdfFlow pdf, IReadOnlyList<WordParagraph> runs, PdfCore.PdfAlign align, PdfSaveOptions? options) {
            foreach (WordParagraph run in runs) {
                if (run.IsImage && run.Image != null) {
                    RenderNativeImage(pdf, run.Image, align, options, "body paragraph image run");
                }
            }
        }

        private static string? GetNativeSupplementalTextAfterRuns(string content, IReadOnlyList<WordParagraph> runs) {
            if (string.IsNullOrEmpty(content)) {
                return null;
            }

            var renderedText = new StringBuilder();
            foreach (WordParagraph run in runs) {
                if (run.IsImage || IsNativeTextWrappingBreak(run) || string.IsNullOrEmpty(run.Text)) {
                    continue;
                }

                renderedText.Append(run.Text);
            }

            if (renderedText.Length == 0) {
                return content;
            }

            string emittedText = renderedText.ToString();
            if (content.Length <= emittedText.Length ||
                !content.StartsWith(emittedText, StringComparison.Ordinal)) {
                return null;
            }

            return content.Substring(emittedText.Length);
        }

        private static void AddNativeFootnoteReferences(PdfCore.PdfParagraphBuilder builder, IReadOnlyList<int> footnoteNumbers) {
            foreach (int footnoteNumber in footnoteNumbers) {
                builder.Baseline(PdfCore.PdfTextBaseline.Superscript);
                builder.Text(footnoteNumber.ToString(CultureInfo.InvariantCulture));
                builder.Baseline(PdfCore.PdfTextBaseline.Normal);
            }
        }

        private static bool IsNativeTextWrappingBreak(WordParagraph run) =>
            run.IsBreak && run.Break?.BreakType != W.BreakValues.Page;

        private static WordTextBox? GetNativeParagraphTextBox(WordParagraph paragraph, out string? fallbackText) {
            fallbackText = GetNativeParagraphTextBoxPlainText(paragraph);
            WordTextBox? textBox = paragraph.TextBox;
            if (textBox != null || paragraph._paragraph == null) {
                return textBox;
            }

            foreach (W.Run run in paragraph._paragraph.Elements<W.Run>()) {
                if (run.Descendants<Wps.TextBoxInfo2>().Any() ||
                    run.Descendants<DocumentFormat.OpenXml.Vml.TextBox>().Any()) {
                    return new WordTextBox(paragraph._document, paragraph._paragraph, run);
                }
            }

            return null;
        }

        private static string? GetNativeParagraphTextBoxPlainText(WordParagraph paragraph) {
            if (paragraph._paragraph == null) {
                return null;
            }

            var parts = new List<string>();
            foreach (Wps.TextBoxInfo2 textBoxInfo in paragraph._paragraph.Descendants<Wps.TextBoxInfo2>()) {
                parts.AddRange(textBoxInfo.Descendants<W.Text>().Select(text => text.Text));
            }

            foreach (DocumentFormat.OpenXml.Vml.TextBox textBox in paragraph._paragraph.Descendants<DocumentFormat.OpenXml.Vml.TextBox>()) {
                parts.AddRange(textBox.Descendants<W.Text>().Select(text => text.Text));
            }

            string textBoxText = string.Concat(parts);
            return string.IsNullOrWhiteSpace(textBoxText) ? null : textBoxText;
        }

        private static void RenderNativeTextBox(INativePdfFlow pdf, WordTextBox textBox, Dictionary<long, int> footnoteNumbersById, PdfSaveOptions? options, string? fallbackText = null) {
            if (!string.IsNullOrWhiteSpace(fallbackText)) {
                PdfCore.PanelStyle fallbackStyle = CreateNativeTextBoxPanelStyle(textBox);
                pdf.PanelParagraph(builder => builder.Text(fallbackText!), fallbackStyle, PdfCore.PdfAlign.Left);
                return;
            }

            IReadOnlyList<WordParagraph> paragraphs = GetNativeTextBoxParagraphs(textBox);
            if (paragraphs.Count == 0) {
                return;
            }

            PdfCore.PanelStyle style = CreateNativeTextBoxPanelStyle(textBox);
            PdfCore.PdfAlign defaultTextAlign = MapNativeTextBoxTextAlign(paragraphs);
            pdf.PanelParagraph(builder => {
                for (int index = 0; index < paragraphs.Count; index++) {
                    WordParagraph paragraph = paragraphs[index];
                    if (index > 0) {
                        builder.LineBreak();
                    }

                    List<WordParagraph> runs = GetNativeRuns(paragraph);
                    string content = paragraph.IsHyperLink && paragraph.Hyperlink != null ? paragraph.Hyperlink.Text : paragraph.Text;
                    bool hasRenderableRuns = runs.Any(run => !run.IsImage && !string.IsNullOrEmpty(run.Text));
                    List<int> paragraphFootnoteNumbers = GetNativeParagraphFootnoteNumbers(paragraph, runs, Array.Empty<int>(), footnoteNumbersById);
                    AddNativeParagraphContent(builder, paragraph, null, runs, hasRenderableRuns, content, paragraphFootnoteNumbers, options);
                }
            }, style, defaultTextAlign);
        }

        private static IReadOnlyList<WordParagraph> GetNativeTextBoxParagraphs(WordTextBox textBox) {
            IReadOnlyList<WordParagraph> directParagraphs = textBox.Paragraphs;
            if (HasNativeRenderableTextBoxText(directParagraphs)) {
                return directParagraphs;
            }

            IReadOnlyList<WordParagraph> elementParagraphs = CollapseNativeParagraphElements(textBox.Elements)
                .OfType<WordParagraph>()
                .ToList();
            return elementParagraphs;
        }

        private static bool HasNativeRenderableTextBoxText(IEnumerable<WordParagraph> paragraphs) {
            foreach (WordParagraph paragraph in paragraphs) {
                if (!string.IsNullOrWhiteSpace(paragraph.Text)) {
                    return true;
                }

                if (GetNativeRuns(paragraph).Any(run => !run.IsImage && !string.IsNullOrWhiteSpace(run.Text))) {
                    return true;
                }
            }

            return false;
        }

        private static PdfCore.PanelStyle CreateNativeTextBoxPanelStyle(WordTextBox textBox) {
            var style = new PdfCore.PanelStyle {
                BorderColor = PdfCore.PdfColor.Black,
                BorderWidth = 0.75D,
                PaddingX = 6D,
                PaddingY = 4D,
                SpacingAfter = 6D,
                Align = MapNativeTextBoxBoxAlign(textBox.HorizontalAlignment)
            };

            double maxWidth = ConvertNativeEmusToPoints(textBox.Width);
            if (maxWidth > 0D) {
                style.MaxWidth = maxWidth;
            }

            return style;
        }

        private static PdfCore.PdfAlign MapNativeTextBoxBoxAlign(WordHorizontalAlignmentValues alignment) {
            switch (alignment) {
                case WordHorizontalAlignmentValues.Center:
                    return PdfCore.PdfAlign.Center;
                case WordHorizontalAlignmentValues.Right:
                case WordHorizontalAlignmentValues.Outside:
                    return PdfCore.PdfAlign.Right;
                default:
                    return PdfCore.PdfAlign.Left;
            }
        }

        private static PdfCore.PdfAlign MapNativeTextBoxTextAlign(IReadOnlyList<WordParagraph> paragraphs) {
            foreach (WordParagraph paragraph in paragraphs) {
                if (!string.IsNullOrEmpty(paragraph.Text)) {
                    return MapNativeParagraphAlign(paragraph.ParagraphAlignment);
                }
            }

            return PdfCore.PdfAlign.Left;
        }

        private static void RenderNativeHeading(INativePdfFlow pdf, int level, string text, PdfCore.PdfAlign align, PdfCore.PdfColor? color, string? linkUri = null, string? linkDestinationName = null, string? linkContents = null) {
            PdfCore.PdfHeadingStyle style = CreateNativeWordHeadingStyle(level);
            pdf.Heading(level, text, align, color, style, linkUri, linkDestinationName, linkContents);
        }

        private static PdfCore.PdfHeadingStyle CreateNativeWordHeadingStyle(int level) {
            double fontSize = level switch {
                1 => 16D,
                2 => 13D,
                _ => 12D
            };

            return new PdfCore.PdfHeadingStyle {
                FontSize = fontSize,
                LineHeight = 1.18D,
                SpacingBefore = level == 1 ? 24D : 10D,
                SpacingAfter = level == 1 ? 5D : 4D,
                Bold = false,
                ApplySpacingBeforeAtTop = true,
                KeepWithNext = true
            };
        }

        private static (string? LinkUri, string? LinkDestinationName, string? LinkContents) GetNativeHeadingLink(WordParagraph paragraph) {
            if (!paragraph.IsHyperLink || paragraph.Hyperlink == null) {
                return (null, null, null);
            }

            string? contents = string.IsNullOrWhiteSpace(paragraph.Hyperlink.Tooltip)
                ? paragraph.Hyperlink.Text
                : paragraph.Hyperlink.Tooltip;
            if (string.IsNullOrWhiteSpace(contents)) {
                contents = null;
            }

            Uri? uri = paragraph.Hyperlink.Uri;
            if (uri != null && uri.IsAbsoluteUri) {
                return (uri.AbsoluteUri, null, contents);
            }

            string? bookmarkName = paragraph.Hyperlink.Anchor;
            if (!string.IsNullOrWhiteSpace(bookmarkName)) {
                return (null, bookmarkName, contents);
            }

            return (null, null, null);
        }

        private static void AddNativeRun(
            PdfCore.PdfParagraphBuilder builder,
            WordParagraph run,
            WordParagraph paragraphStyleFallback,
            IReadOnlyList<WordTabStop> tabStops,
            ref int tabIndex,
            PdfSaveOptions? options) {
            if (string.IsNullOrEmpty(run.Text)) {
                return;
            }

            ApplyNativeTextStyle(builder, run, paragraphStyleFallback);

            if (run.IsHyperLink && run.Hyperlink != null) {
                AddNativeHyperLinkRun(builder, run.Text, run.Hyperlink, tabStops, ref tabIndex);
            } else {
                AddNativeRunText(builder, run.Text, tabStops, ref tabIndex);
            }

            ResetNativeTextStyle(builder);
        }

        private static void AddNativeText(
            PdfCore.PdfParagraphBuilder builder,
            string text,
            WordParagraph paragraph,
            IReadOnlyList<WordTabStop> tabStops,
            ref int tabIndex) {
            ApplyNativeTextStyle(builder, paragraph);
            AddNativeRunText(builder, text, tabStops, ref tabIndex);
            ResetNativeTextStyle(builder);
        }

        private static void ApplyNativeTextStyle(PdfCore.PdfParagraphBuilder builder, WordParagraph paragraph, WordParagraph? fallback = null) {
            builder.Bold(paragraph.Bold);
            builder.Italic(paragraph.Italic);
            builder.Underline(paragraph.Underline != null);
            builder.Strike(paragraph.Strike || paragraph.DoubleStrike);
            builder.Baseline(GetNativeTextBaseline(paragraph));
            if (paragraph.FontSize.HasValue && paragraph.FontSize.Value > 0) {
                builder.FontSize(paragraph.FontSize.Value);
            }

            if (TryMapNativeFontFamily(paragraph.FontFamily, out PdfCore.PdfStandardFont font) ||
                TryMapNativeFontFamily(paragraph.FontFamilyHighAnsi, out font) ||
                TryMapNativeFontFamily(paragraph.FontFamilyEastAsia, out font) ||
                TryMapNativeFontFamily(paragraph.FontFamilyComplexScript, out font) ||
                (fallback != null && (
                    TryMapNativeFontFamily(fallback.FontFamily, out font) ||
                    TryMapNativeFontFamily(fallback.FontFamilyHighAnsi, out font) ||
                    TryMapNativeFontFamily(fallback.FontFamilyEastAsia, out font) ||
                    TryMapNativeFontFamily(fallback.FontFamilyComplexScript, out font)))) {
                builder.Font(font);
            }

            PdfCore.PdfColor? color = ParseNativeColor(paragraph.ColorHex);
            PdfCore.PdfColor? background = MapNativeHighlight(paragraph.Highlight);
            if (color.HasValue) {
                builder.Color(color.Value);
            }

            if (background.HasValue) {
                builder.BackgroundColor(background.Value);
            }
        }

        private static void ResetNativeTextStyle(PdfCore.PdfParagraphBuilder builder) {
            builder.Bold(false)
                .Italic(false)
                .Underline(false)
                .Strike(false)
                .Baseline(PdfCore.PdfTextBaseline.Normal)
                .ResetColor()
                .ResetFontSize()
                .ResetFont()
                .ResetBackgroundColor();
        }

        private static void AddNativeRunText(PdfCore.PdfParagraphBuilder builder, string text, IReadOnlyList<WordTabStop> tabStops, ref int tabIndex) {
            int currentTabIndex = tabIndex;
            AddNativeTextSegments(
                text,
                value => builder.Text(value),
                () => builder.LineBreak(),
                () => {
                    AddNativeTab(builder, tabStops, currentTabIndex);
                    currentTabIndex++;
                },
                () => currentTabIndex = 0);
            tabIndex = currentTabIndex;
        }

        private static void AddNativeTextSegments(string text, Action<string> addText, Action addLineBreak, Action addTab, Action resetTabs) {
            if (string.IsNullOrEmpty(text)) {
                return;
            }

            var buffer = new StringBuilder();
            for (int index = 0; index < text.Length; index++) {
                char ch = text[index];
                if (ch == '\r') {
                    if (index + 1 < text.Length && text[index + 1] == '\n') {
                        continue;
                    }

                    Flush();
                    addLineBreak();
                    resetTabs();
                    continue;
                }

                if (ch == '\n') {
                    Flush();
                    addLineBreak();
                    resetTabs();
                    continue;
                }

                if (ch == '\t') {
                    Flush();
                    addTab();
                    continue;
                }

                buffer.Append(ch);
            }

            Flush();

            void Flush() {
                if (buffer.Length == 0) {
                    return;
                }

                addText(buffer.ToString());
                buffer.Length = 0;
            }
        }

        private static void AddNativeTab(PdfCore.PdfParagraphBuilder builder, IReadOnlyList<WordTabStop> tabStops, int tabIndex) {
            if (tabIndex < tabStops.Count) {
                WordTabStop tabStop = tabStops[tabIndex];
                builder.Tab(MapNativeTabLeader(tabStop.Leader), MapNativeTabAlignment(tabStop.Alignment));
                return;
            }

            builder.Tab();
        }

        private static void AddNativeHyperLinkRun(PdfCore.PdfParagraphBuilder builder, string text, WordHyperLink hyperlink, IReadOnlyList<WordTabStop> tabStops, ref int tabIndex) {
            Uri? uri = hyperlink.Uri;
            string? linkUri = uri != null && uri.IsAbsoluteUri ? uri.AbsoluteUri : null;
            string? bookmarkName = string.IsNullOrWhiteSpace(hyperlink.Anchor) ? null : hyperlink.Anchor;
            if (linkUri == null && bookmarkName == null) {
                AddNativeRunText(builder, text, tabStops, ref tabIndex);
                return;
            }

            string? contents = GetNativeHyperLinkContents(hyperlink);
            int currentTabIndex = tabIndex;
            AddNativeTextSegments(
                text,
                value => {
                    if (linkUri != null) {
                        builder.Link(value, linkUri, contents: contents);
                    } else {
                        builder.LinkToBookmark(value, bookmarkName!, contents: contents);
                    }
                },
                () => builder.LineBreak(),
                () => {
                    AddNativeTab(builder, tabStops, currentTabIndex);
                    currentTabIndex++;
                },
                () => currentTabIndex = 0);
            tabIndex = currentTabIndex;
        }

        private static string? GetNativeHyperLinkContents(WordHyperLink hyperlink) =>
            string.IsNullOrWhiteSpace(hyperlink.Tooltip) ? null : hyperlink.Tooltip;

        private static void AddNativeHyperLinkRun(PdfCore.PdfParagraphBuilder builder, WordHyperLink hyperlink) {
            int tabIndex = 0;
            AddNativeHyperLinkRun(builder, hyperlink.Text, hyperlink, Array.Empty<WordTabStop>(), ref tabIndex);
        }

        private static void RenderNativeHyperLink(INativePdfFlow pdf, WordHyperLink link) {
            if (link == null || string.IsNullOrEmpty(link.Text)) {
                return;
            }

            pdf.Paragraph(builder => AddNativeHyperLinkRun(builder, link));
        }

        private static void RenderNativeShape(INativePdfFlow pdf, WordShape shape) {
            OfficeShape? nativeShape = CreateNativeShape(shape);
            if (nativeShape == null) {
                return;
            }

            pdf.Shape(nativeShape, PdfCore.PdfAlign.Left, spacingAfter: 6);
        }

        private static OfficeShape? CreateNativeShape(WordShape shape) {
            if (shape == null || shape.Hidden == true) {
                return null;
            }

            OfficeShape? nativeShape;
            if (shape.Line != null) {
                (double x1, double y1) = ParseNativeShapePoint(shape.Line.From?.Value ?? "0pt,0pt");
                (double x2, double y2) = ParseNativeShapePoint(shape.Line.To?.Value ?? "0pt,0pt");
                nativeShape = OfficeShape.Line(x1, y1, x2, y2);
            } else if (shape._polygon != null && TryCreateNativePolygonShape(shape._polygon.Points?.Value, out nativeShape)) {
            } else {
                (double Width, double Height)? dimensions = GetNativeShapeDimensions(shape);
                if (!dimensions.HasValue) {
                    return null;
                }

                double width = dimensions.Value.Width;
                double height = dimensions.Value.Height;
                if (shape._ellipse != null) {
                    nativeShape = OfficeShape.Ellipse(width, height);
                } else if (shape._roundRectangle != null) {
                    double arcSize = shape.ArcSize ?? 0.25D;
                    double cornerRadius = Math.Min(width, height) * Math.Max(0D, Math.Min(1D, arcSize)) / 2D;
                    nativeShape = OfficeShape.RoundedRectangle(width, height, cornerRadius);
                } else if (TryGetNativeDrawingPreset(shape, out A.ShapeTypeValues preset)) {
                    nativeShape = CreateNativeDrawingPresetShape(preset, width, height);
                    if (nativeShape == null) {
                        return null;
                    }
                } else {
                    nativeShape = OfficeShape.Rectangle(width, height);
                }
            }

            if (nativeShape == null) {
                return null;
            }

            ApplyNativeShapeStyle(nativeShape, shape);
            return nativeShape;
        }

        private static (double Width, double Height)? GetNativeShapeDimensions(WordShape shape) {
            double width = shape.Width;
            double height = shape.Height;
            if (width > 0 && height > 0) {
                return (width, height);
            }

            A.Extents? extents = shape._wpsShape?
                .GetFirstChild<Wps.ShapeProperties>()?
                .GetFirstChild<A.Transform2D>()?
                .Extents;

            long? cx = extents?.Cx?.Value;
            long? cy = extents?.Cy?.Value;
            if (!cx.HasValue || !cy.HasValue || cx.Value <= 0 || cy.Value <= 0) {
                return null;
            }

            return (ConvertNativeEmusToPoints(cx.Value), ConvertNativeEmusToPoints(cy.Value));
        }

        private static bool TryGetNativeDrawingPreset(WordShape shape, out A.ShapeTypeValues preset) {
            A.PresetGeometry? geometry = shape._wpsShape?
                .GetFirstChild<Wps.ShapeProperties>()?
                .GetFirstChild<A.PresetGeometry>();
            if (geometry?.Preset?.Value is A.ShapeTypeValues value) {
                preset = value;
                return true;
            }

            preset = default;
            return false;
        }

        private static OfficeShape? CreateNativeDrawingPresetShape(A.ShapeTypeValues preset, double width, double height) {
            if (preset == A.ShapeTypeValues.Line) {
                return OfficeShape.Line(0, height / 2D, width, height / 2D);
            }

            if (preset == A.ShapeTypeValues.Ellipse) {
                return OfficeShape.Ellipse(width, height);
            }

            if (preset == A.ShapeTypeValues.RoundRectangle) {
                return OfficeShape.RoundedRectangle(width, height, Math.Min(width, height) / 6D);
            }

            if (preset == A.ShapeTypeValues.Triangle) {
                return OfficeShape.Polygon(
                    new OfficePoint(width / 2D, 0),
                    new OfficePoint(width, height),
                    new OfficePoint(0, height));
            }

            if (preset == A.ShapeTypeValues.Diamond) {
                return OfficeShape.Polygon(
                    new OfficePoint(width / 2D, 0),
                    new OfficePoint(width, height / 2D),
                    new OfficePoint(width / 2D, height),
                    new OfficePoint(0, height / 2D));
            }

            if (preset == A.ShapeTypeValues.Pentagon) {
                return CreateRegularNativePolygon(5, width, height, -90D);
            }

            if (preset == A.ShapeTypeValues.Hexagon) {
                return OfficeShape.Polygon(
                    new OfficePoint(width * 0.25D, 0),
                    new OfficePoint(width * 0.75D, 0),
                    new OfficePoint(width, height / 2D),
                    new OfficePoint(width * 0.75D, height),
                    new OfficePoint(width * 0.25D, height),
                    new OfficePoint(0, height / 2D));
            }

            if (preset == A.ShapeTypeValues.RightArrow) {
                return OfficeShape.Polygon(
                    new OfficePoint(0, height * 0.25D),
                    new OfficePoint(width * 0.6D, height * 0.25D),
                    new OfficePoint(width * 0.6D, 0),
                    new OfficePoint(width, height / 2D),
                    new OfficePoint(width * 0.6D, height),
                    new OfficePoint(width * 0.6D, height * 0.75D),
                    new OfficePoint(0, height * 0.75D));
            }

            if (preset == A.ShapeTypeValues.LeftArrow) {
                return OfficeShape.Polygon(
                    new OfficePoint(width, height * 0.25D),
                    new OfficePoint(width * 0.4D, height * 0.25D),
                    new OfficePoint(width * 0.4D, 0),
                    new OfficePoint(0, height / 2D),
                    new OfficePoint(width * 0.4D, height),
                    new OfficePoint(width * 0.4D, height * 0.75D),
                    new OfficePoint(width, height * 0.75D));
            }

            if (preset == A.ShapeTypeValues.UpArrow) {
                return OfficeShape.Polygon(
                    new OfficePoint(width * 0.25D, height),
                    new OfficePoint(width * 0.25D, height * 0.4D),
                    new OfficePoint(0, height * 0.4D),
                    new OfficePoint(width / 2D, 0),
                    new OfficePoint(width, height * 0.4D),
                    new OfficePoint(width * 0.75D, height * 0.4D),
                    new OfficePoint(width * 0.75D, height));
            }

            if (preset == A.ShapeTypeValues.DownArrow) {
                return OfficeShape.Polygon(
                    new OfficePoint(width * 0.25D, 0),
                    new OfficePoint(width * 0.25D, height * 0.6D),
                    new OfficePoint(0, height * 0.6D),
                    new OfficePoint(width / 2D, height),
                    new OfficePoint(width, height * 0.6D),
                    new OfficePoint(width * 0.75D, height * 0.6D),
                    new OfficePoint(width * 0.75D, 0));
            }

            if (preset == A.ShapeTypeValues.Star5) {
                return CreateNativeStar5(width, height);
            }

            if (preset == A.ShapeTypeValues.Heart) {
                return CreateNativeHeart(width, height);
            }

            if (preset == A.ShapeTypeValues.Cloud) {
                return CreateNativeCloud(width, height);
            }

            if (preset == A.ShapeTypeValues.Donut) {
                return CreateNativeDonut(width, height);
            }

            if (preset == A.ShapeTypeValues.Can) {
                return CreateNativeCan(width, height);
            }

            if (preset == A.ShapeTypeValues.Cube) {
                return CreateNativeCube(width, height);
            }

            if (preset == A.ShapeTypeValues.Rectangle) {
                return OfficeShape.Rectangle(width, height);
            }

            return null;
        }

        private static OfficeShape CreateRegularNativePolygon(int sides, double width, double height, double startAngleDegrees) {
            var points = new OfficePoint[sides];
            double centerX = width / 2D;
            double centerY = height / 2D;
            double radiusX = width / 2D;
            double radiusY = height / 2D;
            for (int i = 0; i < sides; i++) {
                double angle = (startAngleDegrees + (360D * i / sides)) * Math.PI / 180D;
                points[i] = new OfficePoint(centerX + radiusX * Math.Cos(angle), centerY + radiusY * Math.Sin(angle));
            }

            return OfficeShape.Polygon(points);
        }

        private static OfficeShape CreateNativeHeart(double width, double height) =>
            OfficeShape.Path(
                OfficePathCommand.MoveTo(width * 0.5D, height),
                OfficePathCommand.CubicBezierTo(width * 0.18D, height * 0.72D, 0, height * 0.52D, 0, height * 0.28D),
                OfficePathCommand.CubicBezierTo(0, height * 0.08D, width * 0.16D, 0, width * 0.31D, 0),
                OfficePathCommand.CubicBezierTo(width * 0.42D, 0, width * 0.49D, height * 0.07D, width * 0.5D, height * 0.18D),
                OfficePathCommand.CubicBezierTo(width * 0.51D, height * 0.07D, width * 0.58D, 0, width * 0.69D, 0),
                OfficePathCommand.CubicBezierTo(width * 0.84D, 0, width, height * 0.08D, width, height * 0.28D),
                OfficePathCommand.CubicBezierTo(width, height * 0.52D, width * 0.82D, height * 0.72D, width * 0.5D, height),
                OfficePathCommand.Close());

        private static OfficeShape CreateNativeCloud(double width, double height) =>
            OfficeShape.Path(
                OfficePathCommand.MoveTo(width * 0.18D, height * 0.7D),
                OfficePathCommand.CubicBezierTo(width * 0.05D, height * 0.7D, 0, height * 0.58D, width * 0.09D, height * 0.48D),
                OfficePathCommand.CubicBezierTo(width * 0.03D, height * 0.32D, width * 0.19D, height * 0.18D, width * 0.34D, height * 0.26D),
                OfficePathCommand.CubicBezierTo(width * 0.42D, height * 0.04D, width * 0.72D, height * 0.08D, width * 0.75D, height * 0.32D),
                OfficePathCommand.CubicBezierTo(width * 0.94D, height * 0.27D, width, height * 0.46D, width * 0.91D, height * 0.61D),
                OfficePathCommand.CubicBezierTo(width * 0.84D, height * 0.75D, width * 0.63D, height * 0.76D, width * 0.54D, height * 0.68D),
                OfficePathCommand.CubicBezierTo(width * 0.46D, height * 0.82D, width * 0.25D, height * 0.82D, width * 0.18D, height * 0.7D),
                OfficePathCommand.Close());

        private static OfficeShape CreateNativeDonut(double width, double height) {
            List<OfficePathCommand> commands = CreateNativeEllipsePath(width / 2D, height / 2D, width / 2D, height / 2D, clockwise: true);
            commands.AddRange(CreateNativeEllipsePath(width / 2D, height / 2D, width * 0.22D, height * 0.22D, clockwise: false));
            return OfficeShape.Path(commands);
        }

        private static OfficeShape CreateNativeCan(double width, double height) {
            double topY = height * 0.18D;
            double bottomY = height * 0.82D;
            double rx = width / 2D;
            double ry = height * 0.14D;
            double k = 0.5522847498307936D;
            return OfficeShape.Path(
                OfficePathCommand.MoveTo(0, topY),
                OfficePathCommand.CubicBezierTo(0, topY - ry * k, rx - rx * k, topY - ry, rx, topY - ry),
                OfficePathCommand.CubicBezierTo(rx + rx * k, topY - ry, width, topY - ry * k, width, topY),
                OfficePathCommand.LineTo(width, bottomY),
                OfficePathCommand.CubicBezierTo(width, bottomY + ry * k, rx + rx * k, bottomY + ry, rx, bottomY + ry),
                OfficePathCommand.CubicBezierTo(rx - rx * k, bottomY + ry, 0, bottomY + ry * k, 0, bottomY),
                OfficePathCommand.Close());
        }

        private static OfficeShape CreateNativeCube(double width, double height) =>
            OfficeShape.Polygon(
                new OfficePoint(width * 0.32D, 0),
                new OfficePoint(width, height * 0.18D),
                new OfficePoint(width, height * 0.72D),
                new OfficePoint(width * 0.62D, height),
                new OfficePoint(0, height * 0.82D),
                new OfficePoint(0, height * 0.28D));

        private static List<OfficePathCommand> CreateNativeEllipsePath(double centerX, double centerY, double radiusX, double radiusY, bool clockwise) {
            double k = 0.5522847498307936D;
            if (clockwise) {
                return new List<OfficePathCommand> {
                    OfficePathCommand.MoveTo(centerX + radiusX, centerY),
                    OfficePathCommand.CubicBezierTo(centerX + radiusX, centerY + radiusY * k, centerX + radiusX * k, centerY + radiusY, centerX, centerY + radiusY),
                    OfficePathCommand.CubicBezierTo(centerX - radiusX * k, centerY + radiusY, centerX - radiusX, centerY + radiusY * k, centerX - radiusX, centerY),
                    OfficePathCommand.CubicBezierTo(centerX - radiusX, centerY - radiusY * k, centerX - radiusX * k, centerY - radiusY, centerX, centerY - radiusY),
                    OfficePathCommand.CubicBezierTo(centerX + radiusX * k, centerY - radiusY, centerX + radiusX, centerY - radiusY * k, centerX + radiusX, centerY),
                    OfficePathCommand.Close()
                };
            }

            return new List<OfficePathCommand> {
                OfficePathCommand.MoveTo(centerX + radiusX, centerY),
                OfficePathCommand.CubicBezierTo(centerX + radiusX, centerY - radiusY * k, centerX + radiusX * k, centerY - radiusY, centerX, centerY - radiusY),
                OfficePathCommand.CubicBezierTo(centerX - radiusX * k, centerY - radiusY, centerX - radiusX, centerY - radiusY * k, centerX - radiusX, centerY),
                OfficePathCommand.CubicBezierTo(centerX - radiusX, centerY + radiusY * k, centerX - radiusX * k, centerY + radiusY, centerX, centerY + radiusY),
                OfficePathCommand.CubicBezierTo(centerX + radiusX * k, centerY + radiusY, centerX + radiusX, centerY + radiusY * k, centerX + radiusX, centerY),
                OfficePathCommand.Close()
            };
        }

        private static OfficeShape CreateNativeStar5(double width, double height) {
            var points = new OfficePoint[10];
            double centerX = width / 2D;
            double centerY = height / 2D;
            double outerX = width / 2D;
            double outerY = height / 2D;
            double innerX = outerX * 0.45D;
            double innerY = outerY * 0.45D;
            for (int i = 0; i < points.Length; i++) {
                bool outer = i % 2 == 0;
                double angle = (-90D + 36D * i) * Math.PI / 180D;
                double radiusX = outer ? outerX : innerX;
                double radiusY = outer ? outerY : innerY;
                points[i] = new OfficePoint(centerX + radiusX * Math.Cos(angle), centerY + radiusY * Math.Sin(angle));
            }

            return OfficeShape.Polygon(points);
        }

        private static bool TryCreateNativePolygonShape(string? pointsText, out OfficeShape? shape) {
            shape = null;
            if (string.IsNullOrWhiteSpace(pointsText)) {
                return false;
            }

            string text = pointsText!;
            var points = new List<OfficePoint>();
            foreach (string token in text.Split(new[] { ' ', ';' }, StringSplitOptions.RemoveEmptyEntries)) {
                string[] parts = token.Split(',');
                if (parts.Length != 2 ||
                    !double.TryParse(parts[0], NumberStyles.Float, CultureInfo.InvariantCulture, out double x) ||
                    !double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out double y)) {
                    return false;
                }

                points.Add(new OfficePoint(x, y));
            }

            if (points.Count < 3) {
                return false;
            }

            shape = OfficeShape.Polygon(points);
            return true;
        }

        private static void ApplyNativeShapeStyle(OfficeShape nativeShape, WordShape wordShape) {
            if (nativeShape.Kind != OfficeShapeKind.Line) {
                PdfCore.PdfColor? fill = ParseNativeColor(wordShape.FillColorHex);
                if (fill.HasValue) {
                    nativeShape.FillColor = fill.Value.ToOfficeColor();
                }
            }

            bool drawStroke = nativeShape.Kind == OfficeShapeKind.Line ||
                              wordShape.Stroked == true ||
                              (wordShape._wpsShape != null && !string.IsNullOrWhiteSpace(wordShape.StrokeColorHex));
            if (!drawStroke) {
                nativeShape.StrokeColor = null;
                nativeShape.StrokeWidth = 0;
                return;
            }

            PdfCore.PdfColor? stroke = ParseNativeColor(wordShape.StrokeColorHex);
            nativeShape.StrokeColor = (stroke ?? PdfCore.PdfColor.Black).ToOfficeColor();
            nativeShape.StrokeWidth = Math.Max(0D, wordShape.StrokeWeight ?? 1D);
        }

        private static (double X, double Y) ParseNativeShapePoint(string value) {
            string[] parts = value.Split(',');
            if (parts.Length != 2) {
                return (0D, 0D);
            }

            return (ParseNativeShapePointPart(parts[0]), ParseNativeShapePointPart(parts[1]));
        }

        private static double ParseNativeShapePointPart(string value) {
            string normalized = value.Trim().Replace("pt", string.Empty);
            return double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out double result) ? result : 0D;
        }

        private static void RenderNativeTable(INativePdfFlow pdf, WordTable table, Func<WordParagraph, (int Level, string Marker)?> getMarker, Dictionary<long, int> footnoteNumbersById, PdfSaveOptions? options) {
            RecordNativeBodyTableDiagnostics(table, options, "body table");

            TableLayout layout = TableLayoutCache.GetLayout(table);
            var rows = new List<PdfCore.PdfTableCell[]>();
            var cellFills = new Dictionary<(int Row, int Column), PdfCore.PdfColor>();
            var cellBorders = new Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>();
            var cellPaddings = new Dictionary<(int Row, int Column), PdfCore.PdfCellPadding>();
            var cellAlignments = new Dictionary<(int Row, int Column), PdfCore.PdfColumnAlign>();
            var cellVerticalAlignments = new Dictionary<(int Row, int Column), PdfCore.PdfCellVerticalAlign>();
            var horizontalAlignments = CreateNativeTableHorizontalAlignments(layout);
            var verticalAlignments = CreateNativeTableVerticalAlignments(layout);
            for (int rowIndex = 0; rowIndex < layout.Rows.Count; rowIndex++) {
                IReadOnlyList<WordTableCell> row = layout.Rows[rowIndex];
                var nativeCells = new List<PdfCore.PdfTableCell>();
                int logicalColumnIndex = 0;
                for (int columnIndex = 0; columnIndex < row.Count; columnIndex++) {
                    WordTableCell cell = row[columnIndex];
                    if (IsNativeHorizontalMergeContinuation(cell)) {
                        continue;
                    }

                    int columnSpan = GetNativeCellColumnSpan(cell);
                    if (IsNativeVerticalMergeContinuation(cell)) {
                        logicalColumnIndex += columnSpan;
                        continue;
                    }

                    IReadOnlyList<PdfCore.TextRun> cellRuns = CreateNativeCellRuns(cell, footnoteNumbersById);
                    IReadOnlyList<PdfCore.PdfTableCellCheckBox> checkBoxes = CreateNativeTableCellCheckBoxes(cell);
                    IReadOnlyList<PdfCore.PdfTableCellFormField> formFields = CreateNativeTableCellFormFields(cell);
                    IReadOnlyList<PdfCore.PdfTableCellImage> images = CreateNativeTableCellImages(cell);
                    (string? LinkUri, string? LinkContents) link = GetNativeCellLink(cell);
                    int rowSpan = GetNativeCellRowSpan(cell);
                    nativeCells.Add(new PdfCore.PdfTableCell(
                        cellRuns,
                        columnSpan,
                        link.LinkUri,
                        link.LinkContents,
                        rowSpan,
                        checkBoxes.Count == 0 ? null : checkBoxes,
                        formFields.Count == 0 ? null : formFields,
                        images.Count == 0 ? null : images));

                    PdfCore.PdfColor? fill = ParseNativeColor(cell.ShadingFillColorHex);
                    if (fill.HasValue) {
                        cellFills[(rowIndex, logicalColumnIndex)] = fill.Value;
                    }

                    PdfCore.PdfCellBorder? border = CreateNativeTableCellBorder(cell.Borders);
                    if (border != null) {
                        cellBorders[(rowIndex, logicalColumnIndex)] = border;
                    }

                    PdfCore.PdfCellPadding? padding = CreateNativeTableCellPadding(cell);
                    if (padding != null) {
                        cellPaddings[(rowIndex, logicalColumnIndex)] = padding;
                    }

                    PdfCore.PdfColumnAlign cellAlignment = GetNativeCellHorizontalAlignment(cell);
                    if (cellAlignment != PdfCore.PdfColumnAlign.Left) {
                        cellAlignments[(rowIndex, logicalColumnIndex)] = cellAlignment;
                    }

                    PdfCore.PdfCellVerticalAlign cellVerticalAlignment = MapNativeCellVerticalAlign(cell.VerticalAlignment);
                    if (cellVerticalAlignment != PdfCore.PdfCellVerticalAlign.Top) {
                        cellVerticalAlignments[(rowIndex, logicalColumnIndex)] = cellVerticalAlignment;
                    }

                    logicalColumnIndex += columnSpan;
                }

                rows.Add(nativeCells.ToArray());
            }

            if (rows.Count == 0) {
                return;
            }

            PdfCore.PdfTableStyle style = CreateNativeTableStyle(table, rows.Count, options);
            if (cellFills.Count > 0) {
                style.CellFills = cellFills;
            }

            if (cellBorders.Count > 0) {
                style.CellBorders = cellBorders;
            }

            if (cellPaddings.Count > 0) {
                style.CellPaddings = cellPaddings;
            }

            if (cellAlignments.Count > 0) {
                style.CellAlignments = cellAlignments;
            }

            if (cellVerticalAlignments.Count > 0) {
                style.CellVerticalAlignments = cellVerticalAlignments;
            }

            style.ColumnWidthPoints = CreateNativeColumnWidthPoints(layout, style);

            if (horizontalAlignments != null) {
                style.Alignments = horizontalAlignments;
            }

            if (verticalAlignments != null) {
                style.VerticalAlignments = verticalAlignments;
            }

            pdf.Table(rows, MapNativeTableAlignment(table.Alignment), style);
        }

        private static List<double?>? CreateNativeColumnWidthPoints(TableLayout layout, PdfCore.PdfTableStyle style) {
            if (style.AutoFitColumns || layout.ColumnWidths.Length == 0 || !layout.ColumnWidths.All(width => width > 0)) {
                return null;
            }

            var widths = layout.ColumnWidths.Select(width => (double)width).ToList();
            double totalWidth = widths.Sum();
            if (style.MaxWidth.HasValue && totalWidth > style.MaxWidth.Value + 0.001D) {
                double scale = style.MaxWidth.Value / totalWidth;
                for (int i = 0; i < widths.Count; i++) {
                    widths[i] *= scale;
                }
            }

            return widths.Select(width => (double?)width).ToList();
        }

        private static PdfCore.PdfTableStyle CreateNativeTableStyle(WordTable table, int rowCount, PdfSaveOptions? options) {
            PdfCore.PdfTableStyle style = ResolveNativeWordTableStyle(table) ?? new PdfCore.PdfTableStyle {
                RowStripeFill = null
            };
            style.FontSize ??= 10D;
            style.LineHeight ??= 1.15D;

            int repeatedHeaderRowCount = GetNativeTableRepeatedHeaderRowCount(table, rowCount);
            style.HeaderRowCount = GetNativeTableVisualHeaderRowCount(table, rowCount, repeatedHeaderRowCount);
            style.RepeatHeaderRowCount = repeatedHeaderRowCount;
            if (options?.DefaultTableBorders == true && style.BorderColor == null) {
                style.BorderColor = PdfCore.PdfColor.LightGray;
            }

            ApplyNativeTableBorders(table, style);
            ApplyNativeTableDefaultCellMargins(table, style);
            ApplyNativeTableLayoutOptions(table, style);
            ApplyNativeTableRowOptions(table, style);
            return style;
        }

        private static void ApplyNativeTableLayoutOptions(WordTable table, PdfCore.PdfTableStyle style) {
            W.TableProperties? properties = table._tableProperties;
            if (IsNativeTableAutoFitToContents(properties)) {
                style.AutoFitColumns = true;
            }

            double? maxWidth = GetNativeTablePreferredWidth(properties?.TableWidth);
            if (maxWidth.HasValue) {
                style.MaxWidth = maxWidth.Value;
            }

            double? leftIndent = GetNativeTableLeftIndent(properties?.TableIndentation);
            if (leftIndent.HasValue) {
                style.LeftIndent = leftIndent.Value;
            }

            double? cellSpacing = GetNativeTableCellSpacing(properties?.TableCellSpacing);
            if (cellSpacing.HasValue) {
                style.CellSpacing = cellSpacing.Value;
            }
        }

        private static bool IsNativeTableAutoFitToContents(W.TableProperties? properties) =>
            properties?.TableLayout?.Type?.Value == W.TableLayoutValues.Autofit &&
            properties.TableWidth?.Type?.Value == W.TableWidthUnitValues.Auto;

        private static double? GetNativeTablePreferredWidth(W.TableWidth? width) {
            if (width?.Type?.Value != W.TableWidthUnitValues.Dxa) {
                return null;
            }

            return ConvertNativeTwipsToPoints(width.Width?.Value);
        }

        private static double? GetNativeTableLeftIndent(W.TableIndentation? indentation) {
            if (indentation?.Type?.Value != W.TableWidthUnitValues.Dxa || indentation.Width == null) {
                return null;
            }

            return ConvertNativeTwipsToPoints(indentation.Width.Value);
        }

        private static double? GetNativeTableCellSpacing(W.TableCellSpacing? spacing) {
            if (spacing?.Type?.Value != W.TableWidthUnitValues.Dxa) {
                return null;
            }

            return ConvertNativeTwipsToPoints(spacing.Width?.Value);
        }

        private static void ApplyNativeTableBorders(WordTable table, PdfCore.PdfTableStyle style) {
            (PdfCore.PdfColor Color, double Width)? border = GetNativeUniformTableBorder(table._tableProperties?.TableBorders);
            if (border == null) {
                return;
            }

            style.BorderColor = border.Value.Color;
            style.BorderWidth = border.Value.Width;
        }

        private static (PdfCore.PdfColor Color, double Width)? GetNativeUniformTableBorder(W.TableBorders? borders) {
            if (borders == null) {
                return null;
            }

            W.BorderType?[] allBorders = {
                borders.TopBorder,
                borders.BottomBorder,
                borders.LeftBorder,
                borders.RightBorder,
                borders.InsideHorizontalBorder,
                borders.InsideVerticalBorder
            };

            if (allBorders.Any(border => border == null || !HasNativeBorder(border.Val?.Value))) {
                return null;
            }

            W.BorderValues style = allBorders[0]!.Val!.Value;
            if (allBorders.Any(border => border!.Val?.Value != style)) {
                return null;
            }

            uint size = allBorders[0]!.Size?.Value ?? 4U;
            if (allBorders.Any(border => (border!.Size?.Value ?? 4U) != size)) {
                return null;
            }

            string? color = NormalizeNativeBorderColor(allBorders[0]!.Color?.Value);
            if (allBorders.Any(border => !string.Equals(color, NormalizeNativeBorderColor(border!.Color?.Value), StringComparison.OrdinalIgnoreCase))) {
                return null;
            }

            return (ParseNativeColor(color) ?? PdfCore.PdfColor.Black, size / 8D);
        }

        private static void ApplyNativeTableDefaultCellMargins(WordTable table, PdfCore.PdfTableStyle style) {
            W.TableCellMarginDefault? margins = table._tableProperties?.TableCellMarginDefault;
            if (margins == null) {
                style.CellPaddingTop = 3D;
                style.CellPaddingBottom = 3D;
                return;
            }

            double? top = ConvertNativeTwipsToPoints(margins.TopMargin?.Width?.Value);
            double? bottom = ConvertNativeTwipsToPoints(margins.BottomMargin?.Width?.Value);
            double? left = margins.TableCellLeftMargin?.Width == null
                ? null
                : ConvertNativeTwipsToPoints(margins.TableCellLeftMargin.Width.Value);
            double? right = margins.TableCellRightMargin?.Width == null
                ? null
                : ConvertNativeTwipsToPoints(margins.TableCellRightMargin.Width.Value);

            style.CellPaddingTop = top ?? 3D;
            style.CellPaddingBottom = bottom ?? 3D;

            if (left.HasValue) {
                style.CellPaddingLeft = left.Value;
            }

            if (right.HasValue) {
                style.CellPaddingRight = right.Value;
            }
        }

        private static PdfCore.PdfCellPadding? CreateNativeTableCellPadding(WordTableCell cell) {
            double? top = cell.MarginTopWidth.HasValue ? ConvertNativeTwipsToPoints(cell.MarginTopWidth.Value) : null;
            double? bottom = cell.MarginBottomWidth.HasValue ? ConvertNativeTwipsToPoints(cell.MarginBottomWidth.Value) : null;
            double? left = cell.MarginLeftWidth.HasValue ? ConvertNativeTwipsToPoints(cell.MarginLeftWidth.Value) : null;
            double? right = cell.MarginRightWidth.HasValue ? ConvertNativeTwipsToPoints(cell.MarginRightWidth.Value) : null;
            if (!top.HasValue && !bottom.HasValue && !left.HasValue && !right.HasValue) {
                return null;
            }

            return new PdfCore.PdfCellPadding {
                Top = top,
                Bottom = bottom,
                Left = left,
                Right = right
            };
        }

        private static double? ConvertNativeTwipsToPoints(string? value) {
            if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int twips) || twips < 0) {
                return null;
            }

            return twips / 20D;
        }

        private static double? ConvertNativeTwipsToPoints(int twips) {
            return twips < 0 ? null : twips / 20D;
        }

        private static double ConvertNativeEmusToPoints(long emus) {
            return emus <= 0 ? 0D : emus / 12700D;
        }

        private static void ApplyNativeTableRowOptions(WordTable table, PdfCore.PdfTableStyle style) {
            style.AllowRowBreakAcrossPages = table.AllowRowToBreakAcrossPages;
            List<bool?>? rowBreakPolicies = GetNativeTableRowBreakPolicies(table);
            if (rowBreakPolicies != null) {
                style.RowAllowBreakAcrossPages = rowBreakPolicies;
            }

            List<double?>? rowHeights = GetNativeTableRowHeights(table);
            if (rowHeights == null) {
                return;
            }

            double? uniformHeight = GetNativeUniformTableRowHeight(rowHeights);
            if (uniformHeight.HasValue) {
                style.MinRowHeight = uniformHeight.Value;
            } else {
                style.RowMinHeights = rowHeights;
            }
        }

        private static List<bool?>? GetNativeTableRowBreakPolicies(WordTable table) {
            var policies = new List<bool?>(table.Rows.Count);
            bool? firstPolicy = null;
            bool hasMixedPolicies = false;
            foreach (WordTableRow row in table.Rows) {
                bool policy = row.AllowRowToBreakAcrossPages;
                policies.Add(policy);
                if (!firstPolicy.HasValue) {
                    firstPolicy = policy;
                    continue;
                }

                hasMixedPolicies |= firstPolicy.Value != policy;
            }

            return hasMixedPolicies ? policies : null;
        }

        private static List<double?>? GetNativeTableRowHeights(WordTable table) {
            var heights = new List<double?>(table.Rows.Count);
            bool hasHeight = false;
            foreach (WordTableRow row in table.Rows) {
                double? height = row.Height.HasValue && row.Height.Value > 0
                    ? ConvertNativeTwipsToPoints(row.Height.Value)
                    : null;
                heights.Add(height);
                hasHeight |= height.HasValue;
            }

            return hasHeight ? heights : null;
        }

        private static double? GetNativeUniformTableRowHeight(IReadOnlyList<double?> rowHeights) {
            double? height = null;
            foreach (double? rowHeight in rowHeights) {
                if (!rowHeight.HasValue) {
                    return null;
                }

                if (!height.HasValue) {
                    height = rowHeight.Value;
                    continue;
                }

                if (System.Math.Abs(height.Value - rowHeight.Value) > 0.001D) {
                    return null;
                }
            }

            return height;
        }

        private static PdfCore.PdfTableStyle? ResolveNativeWordTableStyle(WordTable table) {
            WordTableStyle? wordStyle = table.Style;
            if (!wordStyle.HasValue) {
                return null;
            }

            return PdfCore.TableStyles.TryFromWordTableStyle(wordStyle.Value.ToString(), out PdfCore.PdfTableStyle? style)
                ? style
                : null;
        }

        private static int GetNativeTableVisualHeaderRowCount(WordTable table, int rowCount, int repeatedHeaderRowCount) {
            if (rowCount == 0) {
                return 0;
            }

            int headerRowCount = repeatedHeaderRowCount;
            if (table.ConditionalFormattingFirstRow == true || headerRowCount > 0) {
                headerRowCount = System.Math.Max(headerRowCount, 1);
            }

            return System.Math.Min(headerRowCount, rowCount);
        }

        private static int GetNativeTableRepeatedHeaderRowCount(WordTable table, int rowCount) {
            if (rowCount == 0 || table.Rows.Count == 0) {
                return 0;
            }

            int repeatedHeaderRowCount = 0;
            foreach (WordTableRow row in table.Rows) {
                if (!row.RepeatHeaderRowAtTheTopOfEachPage) {
                    break;
                }

                repeatedHeaderRowCount++;
                if (repeatedHeaderRowCount == rowCount) {
                    break;
                }
            }

            return repeatedHeaderRowCount;
        }

        private static PdfCore.PdfAlign MapNativeTableAlignment(W.TableRowAlignmentValues? alignment) {
            if (alignment == W.TableRowAlignmentValues.Center) {
                return PdfCore.PdfAlign.Center;
            }

            if (alignment == W.TableRowAlignmentValues.Right) {
                return PdfCore.PdfAlign.Right;
            }

            return PdfCore.PdfAlign.Left;
        }

        private static List<PdfCore.PdfColumnAlign>? CreateNativeTableHorizontalAlignments(TableLayout layout) {
            int columnCount = GetNativeTableColumnCount(layout);
            if (columnCount == 0) {
                return null;
            }

            var alignments = new List<PdfCore.PdfColumnAlign>(columnCount);
            bool hasExplicitAlignment = false;
            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                PdfCore.PdfColumnAlign? columnAlignment = null;
                bool conflict = false;
                foreach ((WordTableCell Cell, int Column, int ColumnSpan) cell in EnumerateNativeTableCells(layout)) {
                    if (columnIndex < cell.Column || columnIndex >= cell.Column + cell.ColumnSpan) {
                        continue;
                    }

                    PdfCore.PdfColumnAlign alignment = GetNativeCellHorizontalAlignment(cell.Cell);
                    if (columnAlignment == null) {
                        columnAlignment = alignment;
                    } else if (columnAlignment.Value != alignment) {
                        conflict = true;
                        break;
                    }
                }

                PdfCore.PdfColumnAlign resolved = conflict ? PdfCore.PdfColumnAlign.Left : columnAlignment ?? PdfCore.PdfColumnAlign.Left;
                if (resolved != PdfCore.PdfColumnAlign.Left) {
                    hasExplicitAlignment = true;
                }

                alignments.Add(resolved);
            }

            return hasExplicitAlignment ? alignments : null;
        }

        private static PdfCore.PdfColumnAlign GetNativeCellHorizontalAlignment(WordTableCell cell) {
            PdfCore.PdfColumnAlign? alignment = null;
            foreach (WordParagraph paragraph in cell.Paragraphs) {
                string text = GetNativeCellParagraphText(paragraph);
                if (string.IsNullOrWhiteSpace(text)) {
                    continue;
                }

                PdfCore.PdfColumnAlign paragraphAlignment = MapNativeColumnAlign(paragraph.ParagraphAlignment);
                if (alignment == null) {
                    alignment = paragraphAlignment;
                } else if (alignment.Value != paragraphAlignment) {
                    return PdfCore.PdfColumnAlign.Left;
                }
            }

            return alignment ?? PdfCore.PdfColumnAlign.Left;
        }

        private static List<PdfCore.PdfCellVerticalAlign>? CreateNativeTableVerticalAlignments(TableLayout layout) {
            int columnCount = GetNativeTableColumnCount(layout);
            if (columnCount == 0) {
                return null;
            }

            var alignments = new List<PdfCore.PdfCellVerticalAlign>(columnCount);
            bool hasExplicitAlignment = false;
            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                PdfCore.PdfCellVerticalAlign? columnAlignment = null;
                bool conflict = false;
                foreach ((WordTableCell Cell, int Column, int ColumnSpan) cell in EnumerateNativeTableCells(layout)) {
                    if (columnIndex < cell.Column || columnIndex >= cell.Column + cell.ColumnSpan) {
                        continue;
                    }

                    PdfCore.PdfCellVerticalAlign alignment = MapNativeCellVerticalAlign(cell.Cell.VerticalAlignment);
                    if (columnAlignment == null) {
                        columnAlignment = alignment;
                    } else if (columnAlignment.Value != alignment) {
                        conflict = true;
                        break;
                    }
                }

                PdfCore.PdfCellVerticalAlign resolved = conflict ? PdfCore.PdfCellVerticalAlign.Top : columnAlignment ?? PdfCore.PdfCellVerticalAlign.Top;
                if (resolved != PdfCore.PdfCellVerticalAlign.Top) {
                    hasExplicitAlignment = true;
                }

                alignments.Add(resolved);
            }

            return hasExplicitAlignment ? alignments : null;
        }

        private static int GetNativeTableColumnCount(TableLayout layout) {
            if (layout.ColumnWidths.Length > 0) {
                return layout.ColumnWidths.Length;
            }

            int columnCount = 0;
            foreach (IReadOnlyList<WordTableCell> row in layout.Rows) {
                int logicalColumn = 0;
                foreach (WordTableCell cell in row) {
                    if (IsNativeHorizontalMergeContinuation(cell)) {
                        continue;
                    }

                    logicalColumn += GetNativeCellColumnSpan(cell);
                }

                if (logicalColumn > columnCount) {
                    columnCount = logicalColumn;
                }
            }

            return columnCount;
        }

        private static IEnumerable<(WordTableCell Cell, int Column, int ColumnSpan)> EnumerateNativeTableCells(TableLayout layout) {
            foreach (IReadOnlyList<WordTableCell> row in layout.Rows) {
                int logicalColumn = 0;
                foreach (WordTableCell cell in row) {
                    if (IsNativeHorizontalMergeContinuation(cell)) {
                        continue;
                    }

                    int columnSpan = GetNativeCellColumnSpan(cell);
                    if (IsNativeVerticalMergeContinuation(cell)) {
                        logicalColumn += columnSpan;
                        continue;
                    }

                    yield return (cell, logicalColumn, columnSpan);
                    logicalColumn += columnSpan;
                }
            }
        }

        private static bool IsNativeHorizontalMergeContinuation(WordTableCell cell) =>
            cell.HorizontalMerge == W.MergedCellValues.Continue;

        private static bool IsNativeVerticalMergeContinuation(WordTableCell cell) =>
            cell.VerticalMerge == W.MergedCellValues.Continue;

        private static int GetNativeCellColumnSpan(WordTableCell cell) =>
            Math.Max(1, cell.ColumnSpan);

        private static int GetNativeCellRowSpan(WordTableCell cell) =>
            Math.Max(1, cell.RowSpan);

        private static (string? LinkUri, string? LinkContents) GetNativeCellLink(WordTableCell cell) {
            string? linkUri = null;
            string? linkContents = null;
            foreach (WordParagraph paragraph in GetNativeCellParagraphs(cell)) {
                if (!TryAddNativeCellLink(paragraph, ref linkUri, ref linkContents)) {
                    return (null, null);
                }

                foreach (WordParagraph run in paragraph.GetRuns()) {
                    if (!TryAddNativeCellLink(run, ref linkUri, ref linkContents)) {
                        return (null, null);
                    }
                }
            }

            return (linkUri, linkContents);
        }

        private static bool TryAddNativeCellLink(WordParagraph paragraph, ref string? linkUri, ref string? linkContents) {
            if (!paragraph.IsHyperLink || paragraph.Hyperlink == null) {
                return true;
            }

            Uri? uri = paragraph.Hyperlink.Uri;
            if (uri == null || !uri.IsAbsoluteUri) {
                return true;
            }

            string candidateUri = uri.AbsoluteUri;
            if (!string.IsNullOrEmpty(linkUri) && !string.Equals(linkUri, candidateUri, StringComparison.Ordinal)) {
                return false;
            }

            linkUri = candidateUri;
            string? contents = string.IsNullOrWhiteSpace(paragraph.Hyperlink.Tooltip)
                ? GetNativeCellParagraphText(paragraph)
                : paragraph.Hyperlink.Tooltip;
            linkContents ??= string.IsNullOrWhiteSpace(contents) ? null : contents;
            return true;
        }

        private static PdfCore.PdfCellBorder? CreateNativeTableCellBorder(WordTableCellBorder borders) {
            bool top = HasNativeBorder(borders.TopStyle);
            bool bottom = HasNativeBorder(borders.BottomStyle);
            bool left = HasNativeBorder(borders.LeftStyle);
            bool right = HasNativeBorder(borders.RightStyle);
            bool diagonalDown = HasNativeBorder(borders.TopLeftToBottomRightStyle);
            bool diagonalUp = HasNativeBorder(borders.TopRightToBottomLeftStyle);
            if (!top && !bottom && !left && !right && !diagonalDown && !diagonalUp) {
                return null;
            }

            if (!diagonalDown && !diagonalUp && TryGetNativeUniformTableCellBorder(borders, out PdfCore.PdfColor uniformColor, out double uniformWidth, out OfficeIMO.Drawing.OfficeStrokeDashStyle uniformDashStyle, out PdfCore.PdfCellBorderLineStyle uniformLineStyle)) {
                return new PdfCore.PdfCellBorder {
                    Color = uniformColor,
                    Width = uniformWidth,
                    DashStyle = uniformDashStyle,
                    LineStyle = uniformLineStyle,
                    Top = top,
                    Bottom = bottom,
                    Left = left,
                    Right = right
                };
            }

            return new PdfCore.PdfCellBorder {
                Color = null,
                Width = 0,
                TopBorder = CreateNativeCellBorderSide(borders.TopStyle, borders.TopColorHex, borders.TopSize),
                BottomBorder = CreateNativeCellBorderSide(borders.BottomStyle, borders.BottomColorHex, borders.BottomSize),
                LeftBorder = CreateNativeCellBorderSide(borders.LeftStyle, borders.LeftColorHex, borders.LeftSize),
                RightBorder = CreateNativeCellBorderSide(borders.RightStyle, borders.RightColorHex, borders.RightSize),
                DiagonalDownBorder = CreateNativeCellBorderSide(borders.TopLeftToBottomRightStyle, borders.TopLeftToBottomRightColorHex, borders.TopLeftToBottomRightSize),
                DiagonalUpBorder = CreateNativeCellBorderSide(borders.TopRightToBottomLeftStyle, borders.TopRightToBottomLeftColorHex, borders.TopRightToBottomLeftSize),
                Top = top,
                Bottom = bottom,
                Left = left,
                Right = right,
                DiagonalDown = diagonalDown,
                DiagonalUp = diagonalUp
            };
        }

        private static bool TryGetNativeUniformTableCellBorder(WordTableCellBorder borders, out PdfCore.PdfColor color, out double width, out OfficeIMO.Drawing.OfficeStrokeDashStyle dashStyle, out PdfCore.PdfCellBorderLineStyle lineStyle) {
            color = PdfCore.PdfColor.Black;
            width = 1D;
            dashStyle = OfficeIMO.Drawing.OfficeStrokeDashStyle.Solid;
            lineStyle = PdfCore.PdfCellBorderLineStyle.Standard;

            string? firstColor = null;
            uint? firstSize = null;
            W.BorderValues? firstStyle = null;
            bool hasFirst = false;
            foreach ((W.BorderValues? BorderStyle, string? Color, DocumentFormat.OpenXml.UInt32Value? Size) side in GetNativeTableCellBorderSides(borders)) {
                if (!HasNativeBorder(side.BorderStyle)) {
                    continue;
                }

                string? sideColor = NormalizeNativeBorderColor(side.Color);
                uint sideSize = side.Size?.Value ?? 4U;
                if (!hasFirst) {
                    firstColor = sideColor;
                    firstSize = sideSize;
                    firstStyle = side.BorderStyle;
                    hasFirst = true;
                    continue;
                }

                if (!string.Equals(firstColor, sideColor, StringComparison.OrdinalIgnoreCase) ||
                    firstSize.GetValueOrDefault() != sideSize ||
                    firstStyle != side.BorderStyle) {
                    return false;
                }
            }

            color = ParseNativeColor(firstColor) ?? PdfCore.PdfColor.Black;
            width = (firstSize ?? 4U) / 8D;
            dashStyle = ToNativeBorderDashStyle(firstStyle);
            lineStyle = ToNativeBorderLineStyle(firstStyle);
            return true;
        }

        private static IEnumerable<(W.BorderValues? BorderStyle, string? Color, DocumentFormat.OpenXml.UInt32Value? Size)> GetNativeTableCellBorderSides(WordTableCellBorder borders) {
            yield return (borders.TopStyle, borders.TopColorHex, borders.TopSize);
            yield return (borders.BottomStyle, borders.BottomColorHex, borders.BottomSize);
            yield return (borders.LeftStyle, borders.LeftColorHex, borders.LeftSize);
            yield return (borders.RightStyle, borders.RightColorHex, borders.RightSize);
        }

        private static PdfCore.PdfCellBorderSide? CreateNativeCellBorderSide(W.BorderValues? borderStyle, string? color, DocumentFormat.OpenXml.UInt32Value? size) {
            if (!HasNativeBorder(borderStyle)) {
                return null;
            }

            return new PdfCore.PdfCellBorderSide {
                Color = ParseNativeColor(NormalizeNativeBorderColor(color)) ?? PdfCore.PdfColor.Black,
                Width = (size?.Value ?? 4U) / 8D,
                DashStyle = ToNativeBorderDashStyle(borderStyle),
                LineStyle = ToNativeBorderLineStyle(borderStyle)
            };
        }

        private static OfficeIMO.Drawing.OfficeStrokeDashStyle ToNativeBorderDashStyle(W.BorderValues? borderStyle) {
            string value = borderStyle?.ToString() ?? string.Empty;
            if (value.IndexOf("dot", StringComparison.OrdinalIgnoreCase) >= 0 &&
                value.IndexOf("dash", StringComparison.OrdinalIgnoreCase) >= 0) {
                return OfficeIMO.Drawing.OfficeStrokeDashStyle.DashDot;
            }

            if (value.IndexOf("dash", StringComparison.OrdinalIgnoreCase) >= 0) {
                return OfficeIMO.Drawing.OfficeStrokeDashStyle.Dash;
            }

            if (value.IndexOf("dot", StringComparison.OrdinalIgnoreCase) >= 0) {
                return OfficeIMO.Drawing.OfficeStrokeDashStyle.Dot;
            }

            return OfficeIMO.Drawing.OfficeStrokeDashStyle.Solid;
        }

        private static PdfCore.PdfCellBorderLineStyle ToNativeBorderLineStyle(W.BorderValues? borderStyle) =>
            borderStyle == W.BorderValues.Double
                ? PdfCore.PdfCellBorderLineStyle.TwoLine
                : PdfCore.PdfCellBorderLineStyle.Standard;

        private static string GetNativeCellText(WordTableCell cell) =>
            GetNativeCellText(cell, null);

        private static IReadOnlyList<PdfCore.TextRun> CreateNativeCellRuns(WordTableCell cell, Dictionary<long, int>? footnoteNumbersById) {
            var runs = new List<PdfCore.TextRun>();
            foreach (WordParagraph paragraph in GetNativeCellParagraphs(cell)) {
                List<PdfCore.TextRun> paragraphRuns = CreateNativeCellParagraphRuns(paragraph, footnoteNumbersById);
                if (paragraphRuns.Count == 0) {
                    continue;
                }

                if (runs.Count > 0) {
                    runs.Add(PdfCore.TextRun.LineBreak());
                }

                runs.AddRange(paragraphRuns);
            }

            return runs;
        }

        private static List<PdfCore.TextRun> CreateNativeCellParagraphRuns(WordParagraph paragraph, Dictionary<long, int>? footnoteNumbersById) {
            var result = new List<PdfCore.TextRun>();
            List<WordParagraph> runs = GetNativeRuns(paragraph);
            string content = paragraph.IsHyperLink && paragraph.Hyperlink != null ? paragraph.Hyperlink.Text : AppendNativeTextWithEquation(paragraph.Text, paragraph);
            bool hasRenderableRuns = runs.Any(run => !run.IsImage && !string.IsNullOrEmpty(run.Text));
            IReadOnlyList<WordTabStop> tabStops = paragraph.TabStops;
            int tabIndex = 0;
            IReadOnlyList<W.SdtRun> repeatingSectionControls = GetNativeRepeatingSectionControls(paragraph);

            if (hasRenderableRuns) {
                foreach (WordParagraph run in runs) {
                    if (run.IsImage && run.Image != null) {
                        continue;
                    }

                    if (IsNativeTextWrappingBreak(run)) {
                        result.Add(PdfCore.TextRun.LineBreak());
                        tabIndex = 0;
                        continue;
                    }

                    AddNativeCellRun(result, run, tabStops, ref tabIndex);
                }

                string? supplementalText = GetNativeSupplementalTextAfterRuns(content, runs);
                if (!string.IsNullOrEmpty(supplementalText)) {
                    AddNativeCellText(result, supplementalText!, paragraph, tabStops, ref tabIndex);
                }
            } else if (paragraph.IsHyperLink && paragraph.Hyperlink != null && !string.IsNullOrEmpty(paragraph.Hyperlink.Text)) {
                AddNativeCellHyperLinkRun(result, paragraph.Hyperlink.Text, paragraph, paragraph.Hyperlink, tabStops, ref tabIndex);
            } else if (!string.IsNullOrEmpty(content)) {
                AddNativeCellText(result, content, paragraph, tabStops, ref tabIndex);
            }

            foreach (W.SdtRun repeatingSection in repeatingSectionControls) {
                foreach (string itemText in GetNativeRepeatingSectionItems(repeatingSection)) {
                    if (string.IsNullOrWhiteSpace(itemText)) {
                        continue;
                    }

                    if (result.Count > 0) {
                        result.Add(PdfCore.TextRun.LineBreak());
                        tabIndex = 0;
                    }

                    AddNativeCellText(result, itemText, paragraph, tabStops, ref tabIndex);
                }
            }

            if (footnoteNumbersById != null) {
                List<int> paragraphFootnoteNumbers = GetNativeParagraphFootnoteNumbers(paragraph, runs, Array.Empty<int>(), footnoteNumbersById);
                AddNativeCellFootnoteReferences(result, paragraphFootnoteNumbers);
            }

            return result;
        }

        private static void AddNativeCellRun(List<PdfCore.TextRun> target, WordParagraph run, IReadOnlyList<WordTabStop> tabStops, ref int tabIndex) {
            if (string.IsNullOrEmpty(run.Text)) {
                return;
            }

            if (run.IsHyperLink && run.Hyperlink != null) {
                AddNativeCellHyperLinkRun(target, run.Text, run, run.Hyperlink, tabStops, ref tabIndex);
                return;
            }

            AddNativeCellTextRuns(target, run.Text, text => CreateNativeCellTextRun(text, run), tabStops, ref tabIndex);
        }

        private static void AddNativeCellText(List<PdfCore.TextRun> target, string text, WordParagraph paragraph, IReadOnlyList<WordTabStop> tabStops, ref int tabIndex) {
            AddNativeCellTextRuns(target, text, value => CreateNativeCellTextRun(value, paragraph), tabStops, ref tabIndex);
        }

        private static void AddNativeCellHyperLinkRun(List<PdfCore.TextRun> target, string text, WordParagraph paragraph, WordHyperLink hyperlink, IReadOnlyList<WordTabStop> tabStops, ref int tabIndex) {
            AddNativeCellTextRuns(target, text, value => CreateNativeCellLinkRun(value, paragraph, hyperlink), tabStops, ref tabIndex);
        }

        private static void AddNativeCellTextRuns(List<PdfCore.TextRun> target, string text, Func<string, PdfCore.TextRun> createRun, IReadOnlyList<WordTabStop> tabStops, ref int tabIndex) {
            int currentTabIndex = tabIndex;
            AddNativeTextSegments(
                text,
                value => AddOrMergeNativeCellTextRun(target, createRun(value)),
                () => target.Add(PdfCore.TextRun.LineBreak()),
                () => {
                    target.Add(CreateNativeCellTabRun(tabStops, currentTabIndex));
                    currentTabIndex++;
                },
                () => currentTabIndex = 0);
            tabIndex = currentTabIndex;
        }

        private static void AddOrMergeNativeCellTextRun(List<PdfCore.TextRun> target, PdfCore.TextRun run) {
            if (target.Count == 0 || !CanMergeNativeCellTextRuns(target[target.Count - 1], run)) {
                target.Add(run);
                return;
            }

            PdfCore.TextRun previous = target[target.Count - 1];
            target[target.Count - 1] = new PdfCore.TextRun(
                previous.Text + run.Text,
                bold: previous.Bold,
                underline: previous.Underline,
                color: previous.Color,
                italic: previous.Italic,
                strike: previous.Strike,
                fontSize: previous.FontSize,
                font: previous.Font,
                baseline: previous.Baseline,
                backgroundColor: previous.BackgroundColor);
        }

        private static bool CanMergeNativeCellTextRuns(PdfCore.TextRun left, PdfCore.TextRun right) =>
            left.LinkUri == null &&
            left.LinkDestinationName == null &&
            right.LinkUri == null &&
            right.LinkDestinationName == null &&
            left.TabLeader == PdfCore.PdfTabLeaderStyle.None &&
            right.TabLeader == PdfCore.PdfTabLeaderStyle.None &&
            left.TabAlignment == PdfCore.PdfTabAlignment.Left &&
            right.TabAlignment == PdfCore.PdfTabAlignment.Left &&
            left.Text != "\n" &&
            left.Text != "\t" &&
            right.Text != "\n" &&
            right.Text != "\t" &&
            left.Bold == right.Bold &&
            left.Underline == right.Underline &&
            left.Italic == right.Italic &&
            left.Strike == right.Strike &&
            NullableDoubleEquals(left.FontSize, right.FontSize) &&
            left.Font == right.Font &&
            left.Baseline == right.Baseline &&
            Equals(left.Color, right.Color) &&
            Equals(left.BackgroundColor, right.BackgroundColor);

        private static PdfCore.TextRun CreateNativeCellTextRun(string text, WordParagraph paragraph) =>
            new PdfCore.TextRun(
                text,
                bold: paragraph.Bold,
                underline: paragraph.Underline != null,
                color: ParseNativeColor(paragraph.ColorHex),
                italic: paragraph.Italic,
                strike: paragraph.Strike || paragraph.DoubleStrike,
                fontSize: paragraph.FontSize.HasValue && paragraph.FontSize.Value > 0 ? paragraph.FontSize.Value : null,
                baseline: GetNativeTextBaseline(paragraph),
                backgroundColor: MapNativeHighlight(paragraph.Highlight));

        private static PdfCore.TextRun CreateNativeCellLinkRun(string text, WordParagraph paragraph, WordHyperLink hyperlink) {
            Uri? uri = hyperlink.Uri;
            string? linkUri = uri != null && uri.IsAbsoluteUri ? uri.AbsoluteUri : null;
            string? destinationName = string.IsNullOrWhiteSpace(hyperlink.Anchor) ? null : hyperlink.Anchor;
            if (linkUri == null && destinationName == null) {
                return CreateNativeCellTextRun(text, paragraph);
            }

            string? contents = string.IsNullOrWhiteSpace(hyperlink.Tooltip) ? null : hyperlink.Tooltip;
            return new PdfCore.TextRun(
                text,
                bold: paragraph.Bold,
                underline: paragraph.Underline != null || linkUri != null || destinationName != null,
                color: ParseNativeColor(paragraph.ColorHex),
                italic: paragraph.Italic,
                strike: paragraph.Strike || paragraph.DoubleStrike,
                fontSize: paragraph.FontSize.HasValue && paragraph.FontSize.Value > 0 ? paragraph.FontSize.Value : null,
                linkUri: linkUri,
                linkContents: contents,
                baseline: GetNativeTextBaseline(paragraph),
                linkDestinationName: destinationName,
                backgroundColor: MapNativeHighlight(paragraph.Highlight));
        }

        private static PdfCore.TextRun CreateNativeCellTabRun(IReadOnlyList<WordTabStop> tabStops, int tabIndex) {
            if (tabIndex < tabStops.Count) {
                WordTabStop tabStop = tabStops[tabIndex];
                return PdfCore.TextRun.Tab(MapNativeTabLeader(tabStop.Leader), MapNativeTabAlignment(tabStop.Alignment));
            }

            return PdfCore.TextRun.Tab();
        }

        private static PdfCore.PdfTextBaseline GetNativeTextBaseline(WordParagraph paragraph) =>
            paragraph.VerticalTextAlignment == W.VerticalPositionValues.Superscript
                ? PdfCore.PdfTextBaseline.Superscript
                : paragraph.VerticalTextAlignment == W.VerticalPositionValues.Subscript
                    ? PdfCore.PdfTextBaseline.Subscript
                    : PdfCore.PdfTextBaseline.Normal;

        private static void AddNativeCellFootnoteReferences(List<PdfCore.TextRun> target, IReadOnlyList<int> footnoteNumbers) {
            foreach (int footnoteNumber in footnoteNumbers) {
                target.Add(PdfCore.TextRun.Superscript(footnoteNumber.ToString(CultureInfo.InvariantCulture)));
            }
        }

        private static string GetNativeCellText(WordTableCell cell, Dictionary<long, int>? footnoteNumbersById) {
            var parts = new List<string>();
            foreach (WordParagraph paragraph in GetNativeCellParagraphs(cell)) {
                string? paragraphText = GetNativeCellParagraphText(paragraph);
                if (!string.IsNullOrEmpty(paragraphText)) {
                    string text = paragraphText;
                    if (footnoteNumbersById != null) {
                        List<int> paragraphFootnoteNumbers = GetNativeParagraphFootnoteNumbers(paragraph, GetNativeRuns(paragraph), Array.Empty<int>(), footnoteNumbersById);
                        if (paragraphFootnoteNumbers.Count > 0) {
                            text += string.Concat(paragraphFootnoteNumbers.Select(number => number.ToString(CultureInfo.InvariantCulture)));
                        }
                    }

                    parts.Add(text);
                }
            }

            return string.Join(Environment.NewLine, parts);
        }

        private static IReadOnlyList<WordParagraph> GetNativeCellParagraphs(WordTableCell cell) =>
            CollapseNativeParagraphElements(cell.Paragraphs.Cast<WordElement>())
                .OfType<WordParagraph>()
                .ToList();

        private static string GetNativeCellParagraphText(WordParagraph paragraph) {
            if (paragraph.IsHyperLink && paragraph.Hyperlink != null && !string.IsNullOrEmpty(paragraph.Hyperlink.Text)) {
                return paragraph.Hyperlink.Text;
            }

            if (!string.IsNullOrEmpty(paragraph.Text)) {
                return AppendNativeTextWithEquation(paragraph.Text, paragraph);
            }

            var parts = new List<string>();
            foreach (WordParagraph run in paragraph.GetRuns()) {
                string runText = run.IsHyperLink && run.Hyperlink != null ? run.Hyperlink.Text : run.Text;
                if (!string.IsNullOrEmpty(runText)) {
                    parts.Add(runText);
                }
            }

            string text = string.Concat(parts);
            return AppendNativeTextWithEquation(text, paragraph);
        }

        private static List<PdfFootnote> CollectNativeFootnotes(IReadOnlyList<WordElement> elements, out Dictionary<long, int> footnoteNumbersById) {
            var footnotes = new List<PdfFootnote>();
            footnoteNumbersById = new Dictionary<long, int>();
            foreach (WordElement element in elements) {
                CollectNativeFootnotes(element, footnotes, footnoteNumbersById);
            }

            return footnotes;
        }

        private static void CollectNativeFootnotes(WordElement element, List<PdfFootnote> footnotes, Dictionary<long, int> footnoteNumbersById) {
            switch (element) {
                case WordFootNote footNote:
                    AddNativeFootnote(footNote, footnotes, footnoteNumbersById);
                    break;
                case WordEndNote endNote:
                    AddNativeEndnote(endNote, footnotes, footnoteNumbersById);
                    break;
                case WordParagraph paragraph:
                    WordFootNote? paragraphFootnote = paragraph.FootNote;
                    if (paragraphFootnote != null) {
                        AddNativeFootnote(paragraphFootnote, footnotes, footnoteNumbersById);
                    }

                    WordEndNote? paragraphEndnote = paragraph.EndNote;
                    if (paragraphEndnote != null) {
                        AddNativeEndnote(paragraphEndnote, footnotes, footnoteNumbersById);
                    }

                    foreach (WordParagraph run in paragraph.GetRuns()) {
                        WordFootNote? runFootnote = run.FootNote;
                        if (runFootnote != null) {
                            AddNativeFootnote(runFootnote, footnotes, footnoteNumbersById);
                        }

                        WordEndNote? runEndnote = run.EndNote;
                        if (runEndnote != null) {
                            AddNativeEndnote(runEndnote, footnotes, footnoteNumbersById);
                        }
                    }

                    break;
                case WordTable table:
                    foreach (WordTableRow row in table.Rows) {
                        foreach (WordTableCell cell in row.Cells) {
                            foreach (WordParagraph paragraph in cell.Paragraphs) {
                                CollectNativeFootnotes(paragraph, footnotes, footnoteNumbersById);
                            }

                            foreach (WordTable nested in cell.NestedTables) {
                                CollectNativeFootnotes(nested, footnotes, footnoteNumbersById);
                            }
                        }
                    }

                    break;
            }
        }

        private static void AddNativeFootnote(WordFootNote footNote, List<PdfFootnote> footnotes, Dictionary<long, int> footnoteNumbersById) {
            long? referenceId = footNote.ReferenceId;
            if (!referenceId.HasValue || referenceId.Value == 0) {
                return;
            }

            long key = GetNativeFootnoteKey(referenceId.Value);
            if (footnoteNumbersById.ContainsKey(key)) {
                return;
            }

            int number = footnotes.Count + 1;
            footnoteNumbersById[key] = number;
            footnotes.Add(new PdfFootnote {
                Number = number,
                Text = GetNativeFootnoteText(footNote)
            });
        }

        private static void AddNativeEndnote(WordEndNote endNote, List<PdfFootnote> footnotes, Dictionary<long, int> footnoteNumbersById) {
            long? referenceId = endNote.ReferenceId;
            if (!referenceId.HasValue || referenceId.Value == 0) {
                return;
            }

            long key = GetNativeEndnoteKey(referenceId.Value);
            if (footnoteNumbersById.ContainsKey(key)) {
                return;
            }

            int number = footnotes.Count + 1;
            footnoteNumbersById[key] = number;
            footnotes.Add(new PdfFootnote {
                Number = number,
                Text = GetNativeEndnoteText(endNote)
            });
        }

        private static string GetNativeFootnoteText(WordFootNote footNote) {
            var parts = new List<string>();
            foreach (WordParagraph paragraph in footNote.Paragraphs ?? Enumerable.Empty<WordParagraph>()) {
                if (!string.IsNullOrWhiteSpace(paragraph.Text)) {
                    parts.Add(paragraph.Text);
                }
            }

            return string.Join(" ", parts);
        }

        private static string GetNativeEndnoteText(WordEndNote endNote) {
            var parts = new List<string>();
            foreach (WordParagraph paragraph in endNote.Paragraphs ?? Enumerable.Empty<WordParagraph>()) {
                if (!string.IsNullOrWhiteSpace(paragraph.Text)) {
                    parts.Add(paragraph.Text);
                }
            }

            return string.Join(" ", parts);
        }

        private static IReadOnlyList<int> GetNativeFootnoteNumbersForElement(IReadOnlyList<WordElement> elements, int index, Dictionary<long, int> footnoteNumbersById) {
            var numbers = new List<int>();
            for (int i = index + 1; i < elements.Count && (elements[i] is WordFootNote || elements[i] is WordEndNote); i++) {
                long? key = GetNativeNoteKey(elements[i]);
                if (key.HasValue && footnoteNumbersById.TryGetValue(key.Value, out int number)) {
                    numbers.Add(number);
                }
            }

            return numbers;
        }

        private static List<int> GetNativeParagraphFootnoteNumbers(WordParagraph paragraph, IReadOnlyList<WordParagraph> runs, IReadOnlyList<int> followingFootnoteNumbers, Dictionary<long, int> footnoteNumbersById) {
            var numbers = new List<int>(followingFootnoteNumbers);
            AddNativeParagraphFootnoteNumber(paragraph, numbers, footnoteNumbersById);
            foreach (WordParagraph run in runs) {
                AddNativeParagraphFootnoteNumber(run, numbers, footnoteNumbersById);
            }

            return numbers.Distinct().ToList();
        }

        private static void AddNativeParagraphFootnoteNumber(WordParagraph paragraph, List<int> numbers, Dictionary<long, int> footnoteNumbersById) {
            WordFootNote? footNote = paragraph.FootNote;
            long? footnoteKey = footNote?.ReferenceId.HasValue == true && footNote.ReferenceId.Value != 0 ? GetNativeFootnoteKey(footNote.ReferenceId.Value) : null;
            if (footnoteKey.HasValue && footnoteNumbersById.TryGetValue(footnoteKey.Value, out int number)) {
                numbers.Add(number);
            }

            WordEndNote? endNote = paragraph.EndNote;
            long? endnoteKey = endNote?.ReferenceId.HasValue == true && endNote.ReferenceId.Value != 0 ? GetNativeEndnoteKey(endNote.ReferenceId.Value) : null;
            if (endnoteKey.HasValue && footnoteNumbersById.TryGetValue(endnoteKey.Value, out number)) {
                numbers.Add(number);
            }
        }

        private static long? GetNativeNoteKey(WordElement element) {
            switch (element) {
                case WordFootNote footNote when footNote.ReferenceId.HasValue && footNote.ReferenceId.Value != 0:
                    return GetNativeFootnoteKey(footNote.ReferenceId.Value);
                case WordEndNote endNote when endNote.ReferenceId.HasValue && endNote.ReferenceId.Value != 0:
                    return GetNativeEndnoteKey(endNote.ReferenceId.Value);
                default:
                    return null;
            }
        }

        private static long GetNativeFootnoteKey(long referenceId) => referenceId;

        private static long GetNativeEndnoteKey(long referenceId) => -referenceId;

        private static void RenderNativeFootnotes(INativePdfFlow pdf, IReadOnlyList<PdfFootnote> footnotes) {
            if (footnotes.Count == 0) {
                return;
            }

            pdf.HR(thickness: 0.5, color: PdfCore.PdfColor.LightGray, spacingBefore: 8, spacingAfter: 4);
            foreach (PdfFootnote footnote in footnotes) {
                pdf.Paragraph(builder => {
                    builder.Baseline(PdfCore.PdfTextBaseline.Superscript);
                    builder.Text(footnote.Number.ToString(CultureInfo.InvariantCulture));
                    builder.Baseline(PdfCore.PdfTextBaseline.Normal);
                    if (!string.IsNullOrWhiteSpace(footnote.Text)) {
                        builder.Text(" ");
                        builder.Text(footnote.Text);
                    }
                });
            }
        }

        private static void RenderNativeImage(INativePdfFlow pdf, WordImage image, PdfCore.PdfAlign align = PdfCore.PdfAlign.Left, PdfSaveOptions? options = null, string source = "body image") {
            if (image == null) {
                return;
            }

            byte[] bytes = ImageEmbedder.GetImageBytes(image);
            if (!IsNativePdfSupportedImageBytes(bytes, out string? unsupportedReason)) {
                if (options != null) {
                    AddNativeExportWarning(
                        options,
                        "NativeBodyImageUnsupported",
                        source,
                        "Word image was not exported because the first-party PDF image writer supports JPEG and simple PNG images only. " + unsupportedReason);
                }

                return;
            }

            double width = image.Width.HasValue ? image.Width.Value * 72D / 96D : 144D;
            double height = image.Height.HasValue ? image.Height.Value * 72D / 96D : 144D;
            pdf.Image(bytes, width, height, align);
        }

        private static bool IsNativePdfSupportedImageBytes(byte[] bytes, out string? unsupportedReason) {
            unsupportedReason = null;
            if (!OfficeImageReader.TryIdentify(bytes, null, out OfficeImageInfo info)) {
                unsupportedReason = "The image format could not be identified.";
                return false;
            }

            if (info.Format == OfficeImageFormat.Jpeg || info.Format == OfficeImageFormat.Png) {
                return true;
            }

            unsupportedReason = "Detected " + info.Format + " (" + info.MimeType + ").";
            return false;
        }

        private static PdfCore.PdfParagraphStyle CreateNativeParagraphStyle(WordParagraph paragraph) {
            var style = new PdfCore.PdfParagraphStyle();
            if (paragraph.LineSpacingBeforePoints.HasValue) {
                style.SpacingBefore = paragraph.LineSpacingBeforePoints.Value;
            }

            if (paragraph.LineSpacingAfterPoints.HasValue) {
                style.SpacingAfter = paragraph.LineSpacingAfterPoints.Value;
            }

            if (paragraph.IndentationBeforePoints.HasValue) {
                style.LeftIndent = paragraph.IndentationBeforePoints.Value;
            }

            if (paragraph.IndentationAfterPoints.HasValue) {
                style.RightIndent = paragraph.IndentationAfterPoints.Value;
            }

            if (paragraph.IndentationFirstLinePoints.HasValue) {
                style.FirstLineIndent = paragraph.IndentationFirstLinePoints.Value;
            }

            double fontSize = paragraph.FontSize.HasValue && paragraph.FontSize.Value > 0 ? paragraph.FontSize.Value : 11D;
            style.LineHeight = ResolveNativeParagraphLineHeight(paragraph, fontSize);

            if (!paragraph.LineSpacingAfterPoints.HasValue) {
                style.SpacingAfter = NativeDefaultParagraphSpacingAfter;
            }

            double? defaultTabStopWidth = GetNativeDefaultTabStopWidth(paragraph);
            if (defaultTabStopWidth.HasValue) {
                style.DefaultTabStopWidth = defaultTabStopWidth.Value;
            }

            style.KeepTogether = paragraph.KeepLinesTogether;
            style.KeepWithNext = paragraph.KeepWithNext;
            style.WidowControl = paragraph.AvoidWidowAndOrphan;
            return style;
        }

        private static double ResolveNativeParagraphLineHeight(WordParagraph paragraph, double fontSize) {
            if (paragraph.LineSpacing.HasValue && paragraph.LineSpacingRule == W.LineSpacingRuleValues.Auto) {
                return Math.Max(0.01D, paragraph.LineSpacing.Value / 240D);
            }

            if (paragraph.LineSpacingPoints.HasValue && fontSize > 0D) {
                return paragraph.LineSpacingPoints.Value / fontSize;
            }

            return NativeDefaultParagraphLineHeight;
        }

        private static double? GetNativeDefaultTabStopWidth(WordParagraph paragraph) {
            int firstTabStop = paragraph.TabStops
                .Where(tabStop => tabStop.Position > 0)
                .Select(tabStop => tabStop.Position)
                .DefaultIfEmpty(0)
                .Min();

            return firstTabStop > 0 ? ConvertNativeTwipsToPoints(firstTabStop) : null;
        }

        private static PdfCore.PdfTabLeaderStyle MapNativeTabLeader(W.TabStopLeaderCharValues leader) {
            if (leader == W.TabStopLeaderCharValues.Dot || leader == W.TabStopLeaderCharValues.MiddleDot || leader == W.TabStopLeaderCharValues.Heavy) {
                return PdfCore.PdfTabLeaderStyle.Dots;
            }

            if (leader == W.TabStopLeaderCharValues.Hyphen) {
                return PdfCore.PdfTabLeaderStyle.Hyphens;
            }

            if (leader == W.TabStopLeaderCharValues.Underscore) {
                return PdfCore.PdfTabLeaderStyle.Underscores;
            }

            return PdfCore.PdfTabLeaderStyle.None;
        }

        private static PdfCore.PdfTabAlignment MapNativeTabAlignment(W.TabStopValues alignment) {
            if (alignment == W.TabStopValues.Center) {
                return PdfCore.PdfTabAlignment.Center;
            }

            if (alignment == W.TabStopValues.Right) {
                return PdfCore.PdfTabAlignment.Right;
            }

            if (alignment == W.TabStopValues.Decimal) {
                return PdfCore.PdfTabAlignment.DecimalSeparator;
            }

            return PdfCore.PdfTabAlignment.Left;
        }

        private static PdfCore.PanelStyle? CreateNativeParagraphPanelStyle(WordParagraph paragraph, PdfCore.PdfParagraphStyle paragraphStyle) {
            PdfCore.PdfColor? background = ParseNativeColor(paragraph.ShadingFillColorHex);
            (PdfCore.PdfColor? Color, double Width)? border = GetNativeUniformParagraphBorder(paragraph.Borders);
            bool renderAsRule = !background.HasValue &&
                (HasNativeOnlyTopParagraphBorder(paragraph.Borders) || HasNativeOnlyBottomParagraphBorder(paragraph.Borders));
            if (renderAsRule) {
                return null;
            }

            bool hasParagraphBorder = HasNativeParagraphBorder(paragraph.Borders);
            if (!background.HasValue && border == null && !hasParagraphBorder) {
                return null;
            }

            var style = new PdfCore.PanelStyle {
                Background = background,
                BorderColor = border?.Color,
                BorderWidth = border?.Width ?? 0D,
                PaddingX = 6,
                PaddingY = 4,
                SpacingBefore = paragraphStyle.SpacingBefore,
                SpacingAfter = paragraphStyle.SpacingAfter ?? 6D,
                Align = MapNativeParagraphAlign(paragraph.ParagraphAlignment, allowJustify: false)
            };

            if (border == null && hasParagraphBorder) {
                style.TopBorder = CreateNativePanelBorder(paragraph.Borders.TopStyle, paragraph.Borders.TopColorHex, paragraph.Borders.TopSize);
                style.RightBorder = CreateNativePanelBorder(paragraph.Borders.RightStyle, paragraph.Borders.RightColorHex, paragraph.Borders.RightSize);
                style.BottomBorder = CreateNativePanelBorder(paragraph.Borders.BottomStyle, paragraph.Borders.BottomColorHex, paragraph.Borders.BottomSize);
                style.LeftBorder = CreateNativePanelBorder(paragraph.Borders.LeftStyle, paragraph.Borders.LeftColorHex, paragraph.Borders.LeftSize);
            }

            return style;
        }

        private static bool IsNativeHorizontalRuleParagraph(WordParagraph paragraph, IReadOnlyList<WordParagraph> runs, string content) {
            if (!string.IsNullOrEmpty(content) ||
                paragraph.Image != null ||
                paragraph.Shape != null ||
                paragraph.TextBox != null ||
                runs.Any(run => run.IsImage || !string.IsNullOrEmpty(run.Text))) {
                return false;
            }

            return HasNativeOnlyBottomParagraphBorder(paragraph.Borders);
        }

        private static PdfCore.PdfHorizontalRuleStyle? CreateNativeHorizontalRuleStyle(WordParagraph paragraph, PdfCore.PdfParagraphStyle paragraphStyle) {
            WordParagraphBorders borders = paragraph.Borders;
            if (!HasNativeBorder(borders.BottomStyle)) {
                return null;
            }

            return new PdfCore.PdfHorizontalRuleStyle {
                Thickness = (borders.BottomSize?.Value ?? 4U) / 8D,
                Color = ParseNativeColor(NormalizeNativeBorderColor(borders.BottomColorHex)) ?? PdfCore.PdfColor.Black,
                SpacingBefore = paragraphStyle.SpacingBefore,
                SpacingAfter = paragraphStyle.SpacingAfter ?? (borders.BottomSpace?.Value ?? 6U),
                KeepWithNext = paragraphStyle.KeepWithNext
            };
        }

        private static PdfCore.PdfHorizontalRuleStyle? CreateNativeBottomBorderRuleStyle(WordParagraph paragraph, PdfCore.PdfParagraphStyle paragraphStyle) {
            WordParagraphBorders borders = paragraph.Borders;
            if (!HasNativeOnlyBottomParagraphBorder(borders)) {
                return null;
            }

            return new PdfCore.PdfHorizontalRuleStyle {
                Thickness = (borders.BottomSize?.Value ?? 4U) / 8D,
                Color = ParseNativeColor(NormalizeNativeBorderColor(borders.BottomColorHex)) ?? PdfCore.PdfColor.Black,
                SpacingBefore = borders.BottomSpace?.Value ?? 0D,
                SpacingAfter = paragraphStyle.SpacingAfter ?? 6D,
                KeepWithNext = paragraphStyle.KeepWithNext
            };
        }

        private static PdfCore.PdfHorizontalRuleStyle? CreateNativeTopBorderRuleStyle(WordParagraph paragraph, PdfCore.PdfParagraphStyle paragraphStyle) {
            WordParagraphBorders borders = paragraph.Borders;
            if (!HasNativeOnlyTopParagraphBorder(borders)) {
                return null;
            }

            return new PdfCore.PdfHorizontalRuleStyle {
                Thickness = (borders.TopSize?.Value ?? 4U) / 8D,
                Color = ParseNativeColor(NormalizeNativeBorderColor(borders.TopColorHex)) ?? PdfCore.PdfColor.Black,
                SpacingBefore = paragraphStyle.SpacingBefore,
                SpacingAfter = borders.TopSpace?.Value ?? 0D,
                KeepWithNext = true
            };
        }

        private static (PdfCore.PdfColor? Color, double Width)? GetNativeUniformParagraphBorder(WordParagraphBorders borders) {
            if (!HasNativeBorder(borders.TopStyle) ||
                !HasNativeBorder(borders.BottomStyle) ||
                !HasNativeBorder(borders.LeftStyle) ||
                !HasNativeBorder(borders.RightStyle)) {
                return null;
            }

            if (borders.TopStyle != borders.BottomStyle ||
                borders.TopStyle != borders.LeftStyle ||
                borders.TopStyle != borders.RightStyle) {
                return null;
            }

            uint topSize = borders.TopSize?.Value ?? 4U;
            if (topSize != (borders.BottomSize?.Value ?? 4U) ||
                topSize != (borders.LeftSize?.Value ?? 4U) ||
                topSize != (borders.RightSize?.Value ?? 4U)) {
                return null;
            }

            string? topColor = NormalizeNativeBorderColor(borders.TopColorHex);
            if (!string.Equals(topColor, NormalizeNativeBorderColor(borders.BottomColorHex), StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(topColor, NormalizeNativeBorderColor(borders.LeftColorHex), StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(topColor, NormalizeNativeBorderColor(borders.RightColorHex), StringComparison.OrdinalIgnoreCase)) {
                return null;
            }

            PdfCore.PdfColor color = ParseNativeColor(topColor) ?? PdfCore.PdfColor.Black;
            return (color, topSize / 8D);
        }

        private static bool HasNativeBorder(W.BorderValues? style) =>
            style != null && style != W.BorderValues.Nil && style != W.BorderValues.None;

        private static bool HasNativeParagraphBorder(WordParagraphBorders borders) =>
            HasNativeBorder(borders.TopStyle) ||
            HasNativeBorder(borders.RightStyle) ||
            HasNativeBorder(borders.BottomStyle) ||
            HasNativeBorder(borders.LeftStyle);

        private static PdfCore.PdfPanelBorder? CreateNativePanelBorder(W.BorderValues? borderStyle, string? color, DocumentFormat.OpenXml.UInt32Value? size) {
            if (!HasNativeBorder(borderStyle)) {
                return null;
            }

            return new PdfCore.PdfPanelBorder {
                Color = ParseNativeColor(NormalizeNativeBorderColor(color)) ?? PdfCore.PdfColor.Black,
                Width = (size?.Value ?? 4U) / 8D
            };
        }

        private static bool HasNativeOnlyBottomParagraphBorder(WordParagraphBorders borders) =>
            HasNativeBorder(borders.BottomStyle) &&
            !HasNativeBorder(borders.TopStyle) &&
            !HasNativeBorder(borders.LeftStyle) &&
            !HasNativeBorder(borders.RightStyle);

        private static bool HasNativeOnlyTopParagraphBorder(WordParagraphBorders borders) =>
            HasNativeBorder(borders.TopStyle) &&
            !HasNativeBorder(borders.BottomStyle) &&
            !HasNativeBorder(borders.LeftStyle) &&
            !HasNativeBorder(borders.RightStyle);

        private static string? NormalizeNativeBorderColor(string? color) =>
            string.IsNullOrWhiteSpace(color) || string.Equals(color, "auto", StringComparison.OrdinalIgnoreCase)
                ? null
                : color;

        private static PdfCore.PdfAlign MapNativeParagraphAlign(W.JustificationValues? alignment, bool allowJustify = true) {
            if (alignment == W.JustificationValues.Center) {
                return PdfCore.PdfAlign.Center;
            }

            if (alignment == W.JustificationValues.Right) {
                return PdfCore.PdfAlign.Right;
            }

            if (allowJustify &&
                (alignment == W.JustificationValues.Both ||
                 alignment == W.JustificationValues.Distribute ||
                 alignment == W.JustificationValues.HighKashida ||
                 alignment == W.JustificationValues.LowKashida ||
                 alignment == W.JustificationValues.MediumKashida ||
                 alignment == W.JustificationValues.ThaiDistribute)) {
                return PdfCore.PdfAlign.Justify;
            }

            return PdfCore.PdfAlign.Left;
        }

        private static PdfCore.PdfColumnAlign MapNativeColumnAlign(W.JustificationValues? alignment) {
            if (alignment == W.JustificationValues.Center) {
                return PdfCore.PdfColumnAlign.Center;
            }

            if (alignment == W.JustificationValues.Right) {
                return PdfCore.PdfColumnAlign.Right;
            }

            return PdfCore.PdfColumnAlign.Left;
        }

        private static PdfCore.PdfCellVerticalAlign MapNativeCellVerticalAlign(W.TableVerticalAlignmentValues? alignment) {
            if (alignment == W.TableVerticalAlignmentValues.Center) {
                return PdfCore.PdfCellVerticalAlign.Middle;
            }

            if (alignment == W.TableVerticalAlignmentValues.Bottom) {
                return PdfCore.PdfCellVerticalAlign.Bottom;
            }

            return PdfCore.PdfCellVerticalAlign.Top;
        }

        private static int GetHeadingLevel(WordParagraph paragraph) {
            if (!paragraph.Style.HasValue) {
                return 0;
            }

            return paragraph.Style.Value switch {
                WordParagraphStyles.Heading1 => 1,
                WordParagraphStyles.Heading2 => 2,
                WordParagraphStyles.Heading3 => 3,
                WordParagraphStyles.Heading4 => 3,
                WordParagraphStyles.Heading5 => 3,
                WordParagraphStyles.Heading6 => 3,
                _ => 0
            };
        }

        private static PdfCore.PdfColor? GetNativeHeadingColor(int headingLevel, PdfCore.PdfColor? explicitColor) {
            if (explicitColor.HasValue || headingLevel <= 0) {
                return explicitColor;
            }

            return PdfCore.PdfColor.FromRgb(47, 84, 150);
        }

        private static PdfCore.PageSize GetNativePageSize(WordSection section, PdfSaveOptions? options) {
            PdfCore.PageSize size;
            if (options?.PageSize != null) {
                size = options.PageSize.Value;
                if (options.Orientation == null) {
                    return size;
                }
            } else if (section.PageSettings.Width?.Value > 0 && section.PageSettings.Height?.Value > 0) {
                size = new PdfCore.PageSize(section.PageSettings.Width.Value / 20D, section.PageSettings.Height.Value / 20D);
            } else if (section.PageSettings.PageSize.HasValue) {
                size = MapNativePageSize(section.PageSettings.PageSize.Value);
            } else if (options?.DefaultPageSize.HasValue == true) {
                size = MapNativePageSize(options.DefaultPageSize.Value);
            } else {
                size = PdfCore.PageSizes.A4;
            }

            PdfPageOrientation orientation;
            if (options?.Orientation != null) {
                orientation = options.Orientation.Value;
            } else if (section.PageSettings.Orientation == W.PageOrientationValues.Landscape) {
                orientation = PdfPageOrientation.Landscape;
            } else if (options?.DefaultOrientation != null) {
                orientation = options.DefaultOrientation == W.PageOrientationValues.Landscape ? PdfPageOrientation.Landscape : PdfPageOrientation.Portrait;
            } else {
                orientation = PdfPageOrientation.Portrait;
            }

            return orientation == PdfPageOrientation.Landscape ? size.Landscape() : size.Portrait();
        }

        private static PdfCore.PageSize MapNativePageSize(WordPageSize pageSize) =>
            pageSize switch {
                WordPageSize.Letter => PdfCore.PageSizes.Letter,
                WordPageSize.Legal => PdfCore.PageSizes.Legal,
                WordPageSize.A3 => new PdfCore.PageSize(842, 1191),
                WordPageSize.A4 => PdfCore.PageSizes.A4,
                WordPageSize.A5 => PdfCore.PageSizes.A5,
                WordPageSize.A6 => new PdfCore.PageSize(298, 420),
                WordPageSize.B5 => new PdfCore.PageSize(499, 709),
                WordPageSize.Executive => new PdfCore.PageSize(522, 756),
                WordPageSize.Statement => new PdfCore.PageSize(396, 612),
                _ => PdfCore.PageSizes.A4
            };

        private static PdfCore.PageMargins GetNativeMargins(WordSection section, PdfSaveOptions? options) {
            if (options?.Margins != null) {
                return options.Margins.Value;
            }

            return new PdfCore.PageMargins(
                (section.Margins.Left?.Value ?? 0) / 20D,
                (section.Margins.Top ?? 0) / 20D,
                (section.Margins.Right?.Value ?? 0) / 20D,
                (section.Margins.Bottom ?? 0) / 20D);
        }

        private static string GetNativePageNumberFormat(PdfSaveOptions? options) {
            string? format = options?.PageNumberFormat;
            if (string.IsNullOrWhiteSpace(format)) {
                return "{page}/{pages}";
            }

            return format!.Replace("{current}", "{page}").Replace("{total}", "{pages}");
        }

        private static string? BuildNativeKeywords(PdfSaveOptions? options, BuiltinDocumentProperties properties) {
            string? keys = options?.Keywords ?? properties.Keywords;
            string? family = options?.FontFamily;
            if (!string.IsNullOrWhiteSpace(family)) {
                keys = string.IsNullOrWhiteSpace(keys) ? family : keys + ";" + family;
            }

            return keys;
        }

        private static PdfCore.PdfColor? ParseNativeColor(string? hex) {
            if (hex == null || string.IsNullOrWhiteSpace(hex) || hex.Equals("auto", StringComparison.OrdinalIgnoreCase)) {
                return null;
            }

            string value = hex.Trim();
            if (value.StartsWith("#", StringComparison.Ordinal)) {
                value = value.Substring(1);
            }

            if (value.Length != 6 ||
                !byte.TryParse(value.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out byte r) ||
                !byte.TryParse(value.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out byte g) ||
                !byte.TryParse(value.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out byte b)) {
                return null;
            }

            return PdfCore.PdfColor.FromRgb(r, g, b);
        }

        private static PdfCore.PdfColor? MapNativeHighlight(W.HighlightColorValues? highlight) {
            if (!highlight.HasValue || highlight.Value == W.HighlightColorValues.None) {
                return null;
            }

            if (highlight.Value == W.HighlightColorValues.Black) return PdfCore.PdfColor.Black;
            if (highlight.Value == W.HighlightColorValues.Blue) return PdfCore.PdfColor.FromRgb(0, 0, 255);
            if (highlight.Value == W.HighlightColorValues.Cyan) return PdfCore.PdfColor.FromRgb(0, 255, 255);
            if (highlight.Value == W.HighlightColorValues.Green) return PdfCore.PdfColor.FromRgb(0, 255, 0);
            if (highlight.Value == W.HighlightColorValues.Magenta) return PdfCore.PdfColor.FromRgb(255, 0, 255);
            if (highlight.Value == W.HighlightColorValues.Red) return PdfCore.PdfColor.FromRgb(255, 0, 0);
            if (highlight.Value == W.HighlightColorValues.Yellow) return PdfCore.PdfColor.FromRgb(255, 255, 0);
            if (highlight.Value == W.HighlightColorValues.White) return PdfCore.PdfColor.White;
            if (highlight.Value == W.HighlightColorValues.DarkBlue) return PdfCore.PdfColor.FromRgb(0, 0, 139);
            if (highlight.Value == W.HighlightColorValues.DarkCyan) return PdfCore.PdfColor.FromRgb(0, 139, 139);
            if (highlight.Value == W.HighlightColorValues.DarkGreen) return PdfCore.PdfColor.FromRgb(0, 100, 0);
            if (highlight.Value == W.HighlightColorValues.DarkMagenta) return PdfCore.PdfColor.FromRgb(139, 0, 139);
            if (highlight.Value == W.HighlightColorValues.DarkRed) return PdfCore.PdfColor.FromRgb(139, 0, 0);
            if (highlight.Value == W.HighlightColorValues.DarkYellow) return PdfCore.PdfColor.FromRgb(184, 134, 11);
            if (highlight.Value == W.HighlightColorValues.LightGray) return PdfCore.PdfColor.FromRgb(211, 211, 211);
            if (highlight.Value == W.HighlightColorValues.DarkGray) return PdfCore.PdfColor.FromRgb(169, 169, 169);

            return null;
        }
    }
}
