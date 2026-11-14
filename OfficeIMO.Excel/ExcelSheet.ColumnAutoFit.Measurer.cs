using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>Measures native AutoFit text using finite style metadata and a bounded text cache.</summary>
        internal sealed partial class AutoFitWidthMeasurer {
            private readonly ExcelTextMeasurer textMeasurer;
            private readonly float defaultMdw;
            private readonly AutoFitTextContext textContext;
            private readonly List<CellFormat>? cellFormats;
            private readonly List<DocumentFormat.OpenXml.Spreadsheet.Font>? fonts;
            private readonly Dictionary<uint, ExcelTextMeasurer.Style> styleCache = new();
            private readonly Dictionary<(uint styleIndex, string text), float> textWidthCache = new();
            private readonly Dictionary<uint, Dictionary<char, float>> charWidthCache = new();

            internal AutoFitWidthMeasurer(Stylesheet? stylesheet, OfficeFontInfo? defaultFont = null) {
                var font = stylesheet?.Fonts?.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().FirstOrDefault();
                textMeasurer = ExcelTextMeasurer.Create(defaultFont ?? (font == null ? null : CreateFontInfoFromOpenXml(font, 11F)));
                defaultMdw = textMeasurer.DefaultStyle.MaximumDigitWidth;
                cellFormats = stylesheet?.CellFormats?.Elements<CellFormat>().ToList();
                fonts = stylesheet?.Fonts?.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().ToList();
                textContext = new AutoFitTextContext(null, cellFormats, BuildCustomNumberFormatCache(stylesheet));
            }

            /// <summary>Measures one serialized streaming cell with its resolved XF and canonical date sampling.</summary>
            internal double MeasureCell(Cell cell, ExcelDateSystem dateSystem, bool tableHeader, bool filterButton) {
                uint style = GetCellDisplayStyleIndex(cell);
                string text = TryGetDateAutoFitSample(cell, textContext, out string sample)
                    ? sample : GetCellAutoFitText(cell, textContext, dateSystem);
                double width = Measure(text, style, GetCellAutoFitRichTextRuns(cell, textContext));
                if (tableHeader) {
                    string header = GetCellAutoFitText(cell, textContext, dateSystem);
                    width = Math.Max(width, MeasureHeader(header, style, bold: true, filterButton));
                }
                return width;
            }

            /// <summary>Converts canonical display text or rich runs to native column units.</summary>
            internal double Measure(string text, uint styleIndex, IReadOnlyList<AutoFitTextRun>? richTextRuns = null) {
                if (defaultMdw <= 0.0001F || string.IsNullOrWhiteSpace(text)) return 0D;
                var style = ResolveStyleInfo(styleIndex);
                float pixels = richTextRuns != null && richTextRuns.Count > 0
                    ? MeasureRichTextWidth(richTextRuns, style) : MeasureTextWidth(text, styleIndex, style);
                return ToColumnWidth(ApplyTextGeometry(text, pixels, styleIndex, style));
            }

            /// <summary>Includes table styling and the native filter-button allowance in a header measurement.</summary>
            internal double MeasureHeader(string text, uint styleIndex, bool bold, bool filterButton) {
                if (defaultMdw <= 0.0001F || string.IsNullOrWhiteSpace(text)) return 0D;
                var style = ResolveStyleInfo(styleIndex);
                if (bold && !style.FontInfo.IsBold)
                    style = textMeasurer.CreateStyle(style.FontInfo.WithStyle(style.FontInfo.Style | OfficeFontStyle.Bold));
                float pixels = 0F;
                foreach (string line in text.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None)) {
                    if (line.Length > 0) pixels = Math.Max(pixels, textMeasurer.MeasureWidthOrDefault(line, style, 0));
                }
                return ToColumnWidth(ApplyTextGeometry(text, pixels, styleIndex, style), filterButton ? AutoFitFilterButtonPixels * style.Dpi / 96D : 0D);
            }

            private float ApplyTextGeometry(string text, float pixels, uint styleIndex, ExcelTextMeasurer.Style style) {
                Alignment? alignment = cellFormats != null && styleIndex < cellFormats.Count
                    ? cellFormats[(int)styleIndex].Alignment : null;
                uint rotation = alignment?.TextRotation?.Value ?? 0U;
                if (rotation == 0U && (alignment?.Indent?.Value ?? 0U) == 0U) return pixels;
                double lineHeight = Math.Ceiling(textMeasurer.MeasureHeightOrDefault("Xg", style, 0) + 2D);
                int lines = rotation == 255U ? 1 : text.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None).Length;
                var box = ProjectAutoFitTextBox(text, pixels, lineHeight * lines, rotation, textMeasurer, style);
                return (float)(box.Width + GetAutoFitIndentPixels(alignment, textMeasurer, style));
            }

            private double ToColumnWidth(float pixels, double buttonPixels = 0D) =>
                Math.Truncate(((pixels * AutoFitTextWidthSafetyFactor) + 5D + buttonPixels) / defaultMdw * 256D) / 256D;

            private ExcelTextMeasurer.Style ResolveStyleInfo(uint styleIndex) {
                if (styleCache.TryGetValue(styleIndex, out var cached)) {
                    return cached;
                }

                var info = textMeasurer.DefaultStyle;
                if (cellFormats != null && fonts != null) {
                    var cellFormat = styleIndex < cellFormats.Count ? cellFormats[(int)styleIndex] : null;
                    if (cellFormat?.FontId != null) {
                        uint fontId = cellFormat.FontId.Value;
                        if (fontId < fonts.Count) {
                            var fontInfo = CreateFontInfoFromOpenXml(fonts[(int)fontId], textMeasurer.DefaultFontSize);
                            info = textMeasurer.CreateStyle(fontInfo);
                        }
                    }
                }

                styleCache[styleIndex] = info;
                return info;
            }

            private float MeasureTextWidth(string text, uint styleIndex, ExcelTextMeasurer.Style styleInfo) {
                if (textWidthCache.TryGetValue((styleIndex, text), out float cached)) {
                    return cached;
                }

                float measured;
                if (text.Contains('\n') || text.Contains('\r')) {
                    measured = 0;
                    string[] lines = text.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
                    foreach (string line in lines) {
                        if (string.IsNullOrEmpty(line)) {
                            continue;
                        }

                        float lineWidth = textMeasurer.MeasureWidthOrDefault(line, styleInfo, 0);
                        if (lineWidth > measured) {
                            measured = lineWidth;
                        }
                    }
                } else if (TryMeasureSimpleAutoFitTextWidth(text, styleIndex, styleInfo, textMeasurer, charWidthCache, out float fastMeasured)) {
                    measured = fastMeasured;
                } else {
                    measured = textMeasurer.MeasureWidthOrDefault(text, styleInfo, 0);
                }

                if (textWidthCache.Count < 1024) textWidthCache[(styleIndex, text)] = measured;
                return measured;
            }

            private float MeasureRichTextWidth(IReadOnlyList<AutoFitTextRun> runs, ExcelTextMeasurer.Style baseStyle) {
                float maxWidth = 0;
                float currentWidth = 0;

                foreach (var run in runs) {
                    var fontInfo = run.CreateFontInfo(baseStyle.FontInfo);
                    var runStyle = textMeasurer.CreateStyle(fontInfo);
                    string[] parts = run.Text.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);

                    for (int i = 0; i < parts.Length; i++) {
                        if (i > 0) {
                            if (currentWidth > maxWidth) {
                                maxWidth = currentWidth;
                            }
                            currentWidth = 0;
                        }

                        if (parts[i].Length > 0) {
                            currentWidth += textMeasurer.MeasureWidthOrDefault(parts[i], runStyle, 0);
                        }
                    }
                }

                return Math.Max(maxWidth, currentWidth);
            }

        }
    }
}
