using System;
using System.Collections.Generic;
using System.IO;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static IEnumerable<WordElement> EnumerateNativeTableCellElements(WordTableCell cell) =>
            EnumerateNativeStructuredContentElements(CollapseNativeParagraphElements(cell.Elements), 0);

        private static IEnumerable<WordElement> EnumerateNativeStructuredContentElements(IEnumerable<WordElement> elements, int structuredDocumentTagDepth) {
            foreach (WordElement element in CollapseNativeParagraphElements(elements)) {
                if (element is not WordStructuredDocumentTag structuredDocumentTag) {
                    yield return element;
                    continue;
                }

                if (structuredDocumentTagDepth >= MaximumNativeStructuredDocumentTagDepth) {
                    throw new InvalidDataException(
                        $"Structured document tag nesting exceeds the supported limit of {MaximumNativeStructuredDocumentTagDepth} levels.");
                }

                IEnumerable<WordElement> structuredElements = GetNativeStructuredBlockElements(
                    structuredDocumentTag.Document,
                    structuredDocumentTag.SdtBlock);
                foreach (WordElement structuredElement in EnumerateNativeStructuredContentElements(structuredElements, structuredDocumentTagDepth + 1)) {
                    yield return structuredElement;
                }
            }
        }

        private static void AppendNativeNestedTableText(
            WordTable nestedTable,
            Dictionary<long, int>? footnoteNumbersById,
            NativeDocumentDefaults nativeDefaults,
            NativeFontMap? nativeFontMap,
            Func<WordParagraph, (int Level, string Marker)?>? getMarker,
            int tableNestingDepth,
            bool ignoreFallbackTableStyle,
            List<PdfCore.PdfTextRun> runs,
            List<PdfCore.PdfTableCellParagraph> paragraphs) {
            int nestedDepth = tableNestingDepth + 1;
            EnsureNativeTableDepth(nestedDepth);
            TableLayout nestedLayout = TableLayoutCache.GetLayout(nestedTable);
            NativeTableStyleDefaults nestedTableStyleDefaults = GetNativeTableStyleDefaults(
                nestedTable,
                nativeDefaults,
                ignoreFallbackTableStyle);
            int nestedColumnCount = GetNativeTableColumnCount(nestedLayout);
            int nestedHeaderRowCount = GetNativeTableVisualHeaderRowCount(
                nestedTable,
                nestedLayout.Rows.Count,
                GetNativeTableRepeatedHeaderRowCount(nestedTable, nestedLayout.Rows.Count));
            int nestedFooterStartRowIndex = nestedTable.ConditionalFormattingLastRow == true && nestedLayout.Rows.Count > nestedHeaderRowCount
                ? nestedLayout.Rows.Count - 1
                : nestedLayout.Rows.Count;
            for (int rowIndex = 0; rowIndex < nestedTable.Rows.Count; rowIndex++) {
                WordTableRow nestedRow = nestedTable.Rows[rowIndex];
                int logicalColumnIndex = GetNativeTableRowStartColumn(nestedLayout, rowIndex);
                foreach (WordTableCell nestedCell in nestedRow.Cells) {
                    int columnSpan = GetNativeCellColumnSpan(nestedCell);
                    NativeTableStyleDefaults nestedCellStyleDefaults = GetNativeTableCellStyleDefaults(
                        nestedTable,
                        nestedTableStyleDefaults,
                        rowIndex,
                        logicalColumnIndex,
                        columnSpan,
                        nestedColumnCount,
                        nestedHeaderRowCount,
                        nestedFooterStartRowIndex);
                    NativeCellText nestedText = CreateNativeCellText(
                        nestedCell,
                        footnoteNumbersById,
                        nativeDefaults,
                        nestedCellStyleDefaults,
                        nativeFontMap,
                        getMarker,
                        nestedDepth,
                        ignoreFallbackTableStyle);
                    logicalColumnIndex += columnSpan;
                    if (nestedText.Runs.Count == 0) {
                        continue;
                    }

                    if (runs.Count > 0) {
                        runs.Add(PdfCore.PdfTextRun.LineBreak());
                    }
                    runs.AddRange(nestedText.Runs);
                    paragraphs.AddRange(nestedText.Paragraphs);
                }
            }
        }
    }
}
