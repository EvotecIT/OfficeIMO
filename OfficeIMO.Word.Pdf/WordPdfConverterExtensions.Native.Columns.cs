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
        private static bool TryRenderNativeSectionColumns(
            PdfCore.PdfPageCompose page,
            WordSection section,
            IReadOnlyList<WordElement> elements,
            Dictionary<WordParagraph, (int Level, string Marker)> listMarkers,
            Dictionary<WordParagraph, (int Level, int Index)> listIndices,
            Dictionary<long, int> footnoteNumbersById,
            PdfSaveOptions? options,
            IReadOnlyList<NativeTableOfContentsEntry> tableOfContentsEntries,
            IReadOnlyDictionary<W.Paragraph, string> headingDestinations,
            NativeDocumentDefaults nativeDefaults,
            NativeFontMap nativeFontMap) {
            IReadOnlyList<double> columnWidthPercents = GetNativeSectionColumnWidthPercents(section);
            int columnCount = columnWidthPercents.Count;
            if (columnCount <= 1) {
                return false;
            }

            IReadOnlyList<IReadOnlyList<WordElement>> columns = SplitNativeElementsByColumnBreaks(elements, columnCount);
            double gap = GetNativeSectionColumnGap(section);
            PdfCore.PageSize pageSize = GetNativePageSize(section, options);
            PdfCore.PageMargins margins = GetNativeMargins(section, options);
            double sectionContentWidth = Math.Max(72D, pageSize.Width - margins.Left - margins.Right);
            double availableColumnWidth = Math.Max(1D, sectionContentWidth - (gap * Math.Max(0, columnCount - 1)));
            page.Content(content => content.Row(row => {
                row.Gap(gap);
                if (section.HasColumnSeparator) {
                    row.ColumnSeparator(PdfCore.PdfColor.Black, 0.5D);
                }

                for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                    IReadOnlyList<WordElement> columnElements = columns[columnIndex];
                    row.Column(columnWidthPercents[columnIndex], column => {
                        INativePdfFlow flow = new NativeSpacingCollapseFlow(new NativePdfColumnFlow(page, column));
                        double columnContentWidth = availableColumnWidth * columnWidthPercents[columnIndex] / 100D;
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
                                footnoteNumbersById,
                                nativeDefaults)) {
                                hasContent = true;
                                continue;
                            }

                            RenderNativeElement(
                                flow,
                                element,
                                section,
                                paragraph => listMarkers.TryGetValue(paragraph, out var marker) ? marker : null,
                                GetNativeFootnoteNumbersForElement(columnElements, i, footnoteNumbersById),
                                footnoteNumbersById,
                                options,
                                tableOfContentsEntries,
                                headingDestinations,
                                columnContentWidth,
                                nativeDefaults,
                                nativeFontMap,
                                renderSpacingOnlyEmptyParagraphLineBox: IsPreviousNativeElementTable(columnElements, i),
                                nextElement: GetNextNativeRenderableElement(columnElements, i));
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
            (ShouldKeepNativeParagraphWithFollowingContent(paragraph) || GetHeadingLevel(paragraph) > 0);

        private static bool ShouldKeepNativeParagraphWithFollowingContent(WordParagraph paragraph) =>
            ReadNativeDirectParagraphOnOff<W.KeepNext>(paragraph) ??
            GetNativeParagraphStyleDefaults(paragraph).KeepWithNext ??
            false;

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

    }
}
