using System.Collections.Generic;
using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using V = DocumentFormat.OpenXml.Vml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static void ApplyNativeHeaderFooterPageNumberStyle(PdfCore.PdfPageBuilder page, params NativeHeaderFooterText?[] parts) {
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

        private static NativeHeaderFooterText? GetNativeHeaderFooterText(WordHeaderFooter? headerFooter, IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers, NativeFontMap? nativeFontMap = null, WordToPdfOptions? options = null) {
            if (headerFooter == null) {
                return null;
            }

            ValidateNativeHeaderFooterTextBoxNesting(headerFooter);
            var parts = new NativeHeaderFooterText();
            List<WordElement> elements = CollapseNativeParagraphElements(headerFooter.Elements).ToList();
            for (int index = 0; index < elements.Count; index++) {
                WordElement element = elements[index];
                switch (element) {
                    case WordParagraph paragraph:
                        if (TryAddNativeHeaderFooterJoinedParagraphText(parts, elements, ref index, listMarkers, nativeFontMap, options)) break;
                        AddNativeHeaderFooterParagraphText(parts, paragraph, listMarkers, nativeFontMap: nativeFontMap);
                        break;
                    case WordTable table:
                        AddNativeHeaderFooterTableText(parts, table, listMarkers, nativeFontMap, options);
                        break;
                    case WordHyperLink link when !string.IsNullOrWhiteSpace(link.Text):
                        parts.AppendLeft(link.Text);
                        break;
                }
            }

            return parts.HasContent ? parts : null;
        }

        private static void AddNativeHeaderFooterTableText(NativeHeaderFooterText parts, WordTable table, IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers, NativeFontMap? nativeFontMap, WordToPdfOptions? options) {
            foreach (WordTableRow row in table.Rows) {
                IReadOnlyList<WordTableCell> cells = row.Cells;
                if (cells.Count == 1) {
                    AddNativeHeaderFooterTableCellText(parts, cells[0], NativeHeaderFooterZone.Left, listMarkers, nativeFontMap, options);
                    continue;
                }

                for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                    NativeHeaderFooterZone zone = cellIndex == 0
                        ? NativeHeaderFooterZone.Left
                        : cellIndex == cells.Count - 1
                            ? NativeHeaderFooterZone.Right
                            : NativeHeaderFooterZone.Center;

                    AddNativeHeaderFooterTableCellText(parts, cells[cellIndex], zone, listMarkers, nativeFontMap, options);
                }
            }
        }

        private static void AddNativeHeaderFooterTableCellText(NativeHeaderFooterText parts, WordTableCell cell, NativeHeaderFooterZone zone, IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers, NativeFontMap? nativeFontMap, WordToPdfOptions? options) {
            List<WordParagraph> paragraphs = EnumerateNativeTableCellParagraphs(cell).ToList();
            int lastContentIndex = -1;
            for (int index = 0; index < paragraphs.Count; index++) {
                string? text = GetNativeHeaderFooterParagraphText(paragraphs[index], listMarkers, out _);
                if (!string.IsNullOrWhiteSpace(text)) lastContentIndex = index;
            }

            List<WordElement> elements = paragraphs.Take(lastContentIndex + 1).Cast<WordElement>().ToList();
            for (int index = 0; index <= lastContentIndex; index++) {
                if (TryAddNativeHeaderFooterJoinedParagraphText(parts, elements, ref index, listMarkers, nativeFontMap, options, zone)) continue;
                AddNativeHeaderFooterParagraphText(parts, paragraphs[index], listMarkers, zone, nativeFontMap);
            }
        }

    }
}
