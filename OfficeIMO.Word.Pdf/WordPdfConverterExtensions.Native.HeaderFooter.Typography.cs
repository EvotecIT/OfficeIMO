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
        private static PdfCore.PdfStandardFont? ResolveNativeHeaderFooterFont(PdfCore.PdfStandardFont baseFont, NativeFontMap? nativeFontMap, params WordHeaderFooter?[] headerFooters) {
            PdfCore.PdfStandardFont? resolvedFont = null;
            foreach (WordHeaderFooter? headerFooter in headerFooters) {
                foreach (NativeResolvedTextStyle style in EnumerateNativeHeaderFooterTextStyles(headerFooter, nativeFontMap)) {
                    if (!style.Font.HasValue) {
                        continue;
                    }

                    PdfCore.PdfStandardFont fontFamily = PdfCore.PdfStandardFontMapper.GetFontFamily(style.Font.Value);
                    if (resolvedFont.HasValue && resolvedFont.Value != fontFamily) {
                        return null;
                    }

                    resolvedFont = fontFamily;
                }
            }

            NativeHeaderFooterEmphasis emphasis = ResolveNativeHeaderFooterEmphasis(headerFooters);
            bool bold = emphasis.Bold == true;
            bool italic = emphasis.Italic == true;
            if (!resolvedFont.HasValue && !bold && !italic) {
                return null;
            }

            PdfCore.PdfStandardFont resolvedFamily = resolvedFont ?? PdfCore.PdfStandardFontMapper.GetFontFamily(baseFont);
            return PdfCore.PdfStandardFontMapper.GetStyledFont(resolvedFamily, bold, italic);
        }

        private static string? ResolveNativeHeaderFooterFontFamily(NativeFontMap nativeFontMap, params WordHeaderFooter?[] headerFooters) {
            string? resolvedFamily = null;
            foreach (WordHeaderFooter? headerFooter in headerFooters) {
                foreach (NativeResolvedTextStyle style in EnumerateNativeHeaderFooterTextStyles(headerFooter, nativeFontMap)) {
                    string? namedFamily = style.FontFamily;
                    if (string.IsNullOrWhiteSpace(namedFamily)) {
                        continue;
                    }

                    if (resolvedFamily != null &&
                        !string.Equals(resolvedFamily, namedFamily, StringComparison.OrdinalIgnoreCase)) {
                        return null;
                    }

                    resolvedFamily = namedFamily;
                }
            }

            return resolvedFamily;
        }

        private static PdfCore.PdfColor? ResolveNativeHeaderFooterColor(params WordHeaderFooter?[] headerFooters) {
            PdfCore.PdfColor? resolvedColor = null;
            foreach (WordHeaderFooter? headerFooter in headerFooters) {
                foreach (NativeResolvedTextStyle style in EnumerateNativeHeaderFooterTextStyles(headerFooter)) {
                    if (!style.Color.HasValue) {
                        return null;
                    }

                    PdfCore.PdfColor color = style.Color.Value;
                    if (resolvedColor.HasValue && !resolvedColor.Value.Equals(color)) {
                        return null;
                    }

                    resolvedColor = color;
                }
            }

            return resolvedColor;
        }

        private static double? ResolveNativeHeaderFooterFontSize(params WordHeaderFooter?[] headerFooters) {
            double? resolvedFontSize = null;
            foreach (WordHeaderFooter? headerFooter in headerFooters) {
                foreach (double fontSize in EnumerateNativeHeaderFooterFontSizes(headerFooter)) {
                    if (resolvedFontSize.HasValue && !NullableDoubleEquals(resolvedFontSize.Value, fontSize)) {
                        return null;
                    }

                    resolvedFontSize = fontSize;
                }
            }

            return resolvedFontSize;
        }

        private static PdfCore.PdfStandardFont ResolveNativeHeaderFooterBaseFont(WordDocument document, WordToPdfOptions? options, bool isHeader) {
            if (options?.PdfOptions != null) {
                return PdfCore.PdfStandardFontMapper.GetFontFamily(isHeader ? options.PdfOptions.HeaderFont : options.PdfOptions.FooterFont);
            }

            foreach (string? familyName in new[] {
                options?.FontFamily,
                GetNativeDocumentDefaults(document).FontFamily,
                document.Settings.FontFamily,
                document.Settings.FontFamilyHighAnsi,
                document.Settings.FontFamilyEastAsia,
                document.Settings.FontFamilyComplexScript
            }) {
                if (PdfCore.PdfStandardFontMapper.TryMapFontFamily(familyName, out PdfCore.PdfStandardFont mappedFont)) {
                    return PdfCore.PdfStandardFontMapper.GetFontFamily(mappedFont);
                }
            }

            return PdfCore.PdfStandardFont.Helvetica;
        }

        private readonly record struct NativeHeaderFooterEmphasis(bool? Bold, bool? Italic);

        private static NativeHeaderFooterEmphasis ResolveNativeHeaderFooterEmphasis(params WordHeaderFooter?[] headerFooters) {
            bool? bold = null;
            bool? italic = null;
            bool boldConflict = false;
            bool italicConflict = false;
            foreach (WordHeaderFooter? headerFooter in headerFooters) {
                foreach (NativeResolvedTextStyle style in EnumerateNativeHeaderFooterTextStyles(headerFooter)) {
                    MergeNativeHeaderFooterEmphasis(ref bold, ref boldConflict, style.Bold);
                    MergeNativeHeaderFooterEmphasis(ref italic, ref italicConflict, style.Italic);
                }
            }

            return new NativeHeaderFooterEmphasis(
                boldConflict ? null : bold,
                italicConflict ? null : italic);
        }

        private static void MergeNativeHeaderFooterEmphasis(ref bool? current, ref bool hasConflict, bool candidate) {
            if (hasConflict) {
                return;
            }

            if (!current.HasValue) {
                current = candidate;
                return;
            }

            if (current.Value != candidate) {
                current = null;
                hasConflict = true;
            }
        }

        private static IEnumerable<string> EnumerateNativeHeaderFooterFontFamilies(WordHeaderFooter? headerFooter) {
            if (headerFooter == null) {
                yield break;
            }

            foreach (WordElement element in CollapseNativeParagraphElements(headerFooter.Elements)) {
                foreach (string familyName in EnumerateNativeHeaderFooterElementFontFamilies(element)) {
                    yield return familyName;
                }
            }
        }

        private static IEnumerable<double> EnumerateNativeHeaderFooterFontSizes(WordHeaderFooter? headerFooter) {
            if (headerFooter == null) {
                yield break;
            }

            foreach (WordElement element in CollapseNativeParagraphElements(headerFooter.Elements)) {
                foreach (double fontSize in EnumerateNativeHeaderFooterElementFontSizes(element)) {
                    yield return fontSize;
                }
            }
        }

        private static IEnumerable<PdfCore.PdfColor> EnumerateNativeHeaderFooterColors(WordHeaderFooter? headerFooter) {
            if (headerFooter == null) {
                yield break;
            }

            foreach (WordElement element in CollapseNativeParagraphElements(headerFooter.Elements)) {
                foreach (PdfCore.PdfColor color in EnumerateNativeHeaderFooterElementColors(element)) {
                    yield return color;
                }
            }
        }

        private static IEnumerable<NativeResolvedTextStyle> EnumerateNativeHeaderFooterTextStyles(WordHeaderFooter? headerFooter, NativeFontMap? nativeFontMap = null) {
            if (headerFooter == null) {
                yield break;
            }

            foreach (WordElement element in CollapseNativeParagraphElements(headerFooter.Elements)) {
                foreach (NativeResolvedTextStyle style in EnumerateNativeHeaderFooterElementTextStyles(element, nativeFontMap)) {
                    yield return style;
                }
            }
        }

        private static IEnumerable<string> EnumerateNativeHeaderFooterElementFontFamilies(WordElement element) {
            if (element is WordParagraph paragraph) {
                foreach (string familyName in EnumerateNativeParagraphFontFamilies(paragraph)) {
                    yield return familyName;
                }

                yield break;
            }

            if (element is not WordTable table) {
                yield break;
            }

            foreach (WordTable currentTable in EnumerateNativeTableTree(table)) {
                foreach (WordTableRow row in currentTable.Rows) {
                    foreach (WordTableCell cell in row.Cells) {
                        foreach (WordParagraph cellParagraph in cell.Paragraphs) {
                            foreach (string familyName in EnumerateNativeParagraphFontFamilies(cellParagraph)) {
                                yield return familyName;
                            }
                        }
                    }
                }
            }
        }

        private static IEnumerable<NativeResolvedTextStyle> EnumerateNativeHeaderFooterElementTextStyles(WordElement element, NativeFontMap? nativeFontMap) {
            if (element is WordParagraph paragraph) {
                foreach (NativeResolvedTextStyle style in EnumerateNativeParagraphTextStyles(paragraph, nativeFontMap)) {
                    yield return style;
                }

                yield break;
            }

            if (element is not WordTable table) {
                yield break;
            }

            foreach (WordTable currentTable in EnumerateNativeTableTree(table)) {
                foreach (WordTableRow row in currentTable.Rows) {
                    foreach (WordTableCell cell in row.Cells) {
                        foreach (WordParagraph cellParagraph in cell.Paragraphs) {
                            foreach (NativeResolvedTextStyle style in EnumerateNativeParagraphTextStyles(cellParagraph, nativeFontMap)) {
                                yield return style;
                            }
                        }
                    }
                }
            }
        }

        private static IEnumerable<double> EnumerateNativeHeaderFooterElementFontSizes(WordElement element) {
            if (element is WordParagraph paragraph) {
                foreach (double fontSize in EnumerateNativeHeaderFooterParagraphFontSizes(paragraph)) yield return fontSize;

                yield break;
            }

            if (element is not WordTable table) {
                yield break;
            }

            foreach (WordTable currentTable in EnumerateNativeTableTree(table)) {
                foreach (WordTableRow row in currentTable.Rows) {
                    foreach (WordTableCell cell in row.Cells) {
                        foreach (WordParagraph cellParagraph in cell.Paragraphs) {
                            foreach (double fontSize in EnumerateNativeHeaderFooterParagraphFontSizes(cellParagraph)) yield return fontSize;
                        }
                    }
                }
            }
        }

        private static IEnumerable<double> EnumerateNativeHeaderFooterParagraphFontSizes(WordParagraph paragraph) {
            double defaultSize = GetNativeDocumentDefaults(paragraph._document).FontSize;
            foreach (NativeResolvedTextStyle style in EnumerateNativeParagraphTextStyles(paragraph)) {
                // A missing run/style size inherits the document's size, as it
                // does in body text; it must not select the PDF zone fallback.
                double fontSize = style.FontSize ?? defaultSize;
                if (fontSize > 0D) yield return fontSize;
            }
        }

        private static IEnumerable<PdfCore.PdfColor> EnumerateNativeHeaderFooterElementColors(WordElement element) {
            if (element is WordParagraph paragraph) {
                foreach (PdfCore.PdfColor color in EnumerateNativeParagraphColors(paragraph)) {
                    yield return color;
                }

                yield break;
            }

            if (element is not WordTable table) {
                yield break;
            }

            foreach (WordTable currentTable in EnumerateNativeTableTree(table)) {
                foreach (WordTableRow row in currentTable.Rows) {
                    foreach (WordTableCell cell in row.Cells) {
                        foreach (WordParagraph cellParagraph in cell.Paragraphs) {
                            foreach (PdfCore.PdfColor color in EnumerateNativeParagraphColors(cellParagraph)) {
                                yield return color;
                            }
                        }
                    }
                }
            }
        }

        private static IEnumerable<NativeResolvedTextStyle> EnumerateNativeParagraphTextStyles(WordParagraph paragraph, NativeFontMap? nativeFontMap = null) {
            List<WordParagraph> runs = GetNativeRuns(paragraph);
            bool emittedRun = false;
            foreach (WordParagraph run in runs) {
                if (run.IsImage || string.IsNullOrWhiteSpace(run.Text)) {
                    continue;
                }

                emittedRun = true;
                yield return ResolveNativeTextRunStyle(run, paragraph, nativeFontMap: nativeFontMap);
            }

            if (!emittedRun && !string.IsNullOrWhiteSpace(paragraph.Text)) {
                yield return ResolveNativeTextRunStyle(paragraph, nativeFontMap: nativeFontMap);
            }
        }

        private static IEnumerable<string> EnumerateNativeParagraphFontFamilies(WordParagraph paragraph) {
            foreach (string familyName in EnumerateNativeParagraphOwnFontFamilies(paragraph)) {
                yield return familyName;
            }

            string? styleFamily = GetNativeParagraphStyleDefaults(paragraph).FontFamily;
            if (!string.IsNullOrWhiteSpace(styleFamily)) {
                yield return styleFamily!;
            }

            string? characterStyleFamily = GetNativeCharacterStyleDefaults(paragraph._document, GetNativeRunProperties(paragraph)).FontFamily;
            if (!string.IsNullOrWhiteSpace(characterStyleFamily)) {
                yield return characterStyleFamily!;
            }

            foreach (WordParagraph run in GetNativeRuns(paragraph)) {
                if (run.IsImage || string.IsNullOrWhiteSpace(run.Text)) {
                    continue;
                }

                foreach (string familyName in EnumerateNativeParagraphOwnFontFamilies(run)) {
                    yield return familyName;
                }

                string? runCharacterStyleFamily = GetNativeCharacterStyleDefaults(run._document, GetNativeRunProperties(run)).FontFamily;
                if (!string.IsNullOrWhiteSpace(runCharacterStyleFamily)) {
                    yield return runCharacterStyleFamily!;
                }
            }
        }

        private static IEnumerable<PdfCore.PdfColor> EnumerateNativeParagraphColors(WordParagraph paragraph) {
            PdfCore.PdfColor? paragraphColor = ParseNativeColor(paragraph.ColorHex);
            if (paragraphColor.HasValue) {
                yield return paragraphColor.Value;
            }

            PdfCore.PdfColor? styleColor = ParseNativeColor(GetNativeParagraphStyleDefaults(paragraph).ColorHex);
            if (styleColor.HasValue) {
                yield return styleColor.Value;
            }

            PdfCore.PdfColor? characterStyleColor = ParseNativeColor(GetNativeCharacterStyleDefaults(paragraph._document, GetNativeRunProperties(paragraph)).ColorHex);
            if (characterStyleColor.HasValue) {
                yield return characterStyleColor.Value;
            }

            foreach (WordParagraph run in GetNativeRuns(paragraph)) {
                if (run.IsImage || string.IsNullOrWhiteSpace(run.Text)) {
                    continue;
                }

                PdfCore.PdfColor? runColor = ParseNativeColor(run.ColorHex);
                if (runColor.HasValue) {
                    yield return runColor.Value;
                }

                PdfCore.PdfColor? runCharacterStyleColor = ParseNativeColor(GetNativeCharacterStyleDefaults(run._document, GetNativeRunProperties(run)).ColorHex);
                if (runCharacterStyleColor.HasValue) {
                    yield return runCharacterStyleColor.Value;
                }
            }
        }

        private static IEnumerable<string> EnumerateNativeParagraphOwnFontFamilies(WordParagraph paragraph) {
            foreach (string? familyName in EnumerateNativeLatinFontFamilies(
                paragraph._document, GetNativeRunProperties(paragraph)?.GetFirstChild<W.RunFonts>()).Concat(new[] {
                paragraph.FontFamilyEastAsia,
                paragraph.FontFamilyComplexScript
            })) {
                if (!string.IsNullOrWhiteSpace(familyName)) {
                    yield return familyName!;
                }
            }
        }

    }
}
