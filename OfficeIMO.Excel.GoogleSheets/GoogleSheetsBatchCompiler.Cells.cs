using OfficeIMO.GoogleWorkspace;
using System.Globalization;
using System.IO;

namespace OfficeIMO.Excel.GoogleSheets {
    internal static partial class GoogleSheetsBatchCompiler {
        private static GoogleSheetsCellValue BuildCellValue(
            ExcelCellSnapshot cell,
            GoogleSheetsFormulaOptions formulaOptions) {
            if (!string.IsNullOrWhiteSpace(cell.Formula)) {
                GoogleSheetsFormulaTranslation translation = GoogleSheetsFormulaCatalog.Translate(cell.Formula!, formulaOptions);
                if (!translation.IsSupported
                    && formulaOptions.UnsupportedFormulaMode == GoogleSheetsUnsupportedFormulaMode.UseCachedValue) {
                    return BuildTypedValue(cell.Value);
                }
                return GoogleSheetsCellValue.Formula(translation.Formula);
            }

            return BuildTypedValue(cell.Value);
        }

        private static GoogleSheetsCellValue BuildTypedValue(object? typedValue) {
            if (typedValue == null) {
                return GoogleSheetsCellValue.Blank();
            }

            if (typedValue is bool booleanValue) {
                return GoogleSheetsCellValue.Boolean(booleanValue);
            }

            if (typedValue is DateTime dateTimeValue) {
                return GoogleSheetsCellValue.DateTime(dateTimeValue);
            }

            if (typedValue is DateTimeOffset dateTimeOffsetValue) {
                return GoogleSheetsCellValue.DateTime(dateTimeOffsetValue.LocalDateTime);
            }

            if (typedValue is byte || typedValue is sbyte || typedValue is short || typedValue is ushort
                || typedValue is int || typedValue is uint || typedValue is long || typedValue is ulong
                || typedValue is float || typedValue is double || typedValue is decimal) {
                return GoogleSheetsCellValue.Number(Convert.ToDouble(typedValue, System.Globalization.CultureInfo.InvariantCulture));
            }

            return GoogleSheetsCellValue.String(Convert.ToString(typedValue, System.Globalization.CultureInfo.InvariantCulture));
        }

        private static string? GetNumberFormatHint(object? typedValue, ExcelCellStyleSnapshot? style) {
            if (style?.IsDateLike == true || typedValue is DateTime || typedValue is DateTimeOffset) {
                return "DateTime";
            }

            if (!string.IsNullOrWhiteSpace(style?.NumberFormatCode)) {
                return style!.NumberFormatCode;
            }

            return null;
        }

        private static GoogleSheetsCellStyle? BuildCellStyle(ExcelCellStyleSnapshot? style) {
            if (style == null) {
                return null;
            }

            return new GoogleSheetsCellStyle {
                SourceStyleIndex = style.StyleIndex,
                NumberFormatId = style.NumberFormatId,
                NumberFormatCode = style.NumberFormatCode,
                IsDateLike = style.IsDateLike,
                Bold = style.Bold,
                Italic = style.Italic,
                Underline = style.Underline,
                Strikethrough = style.Strikethrough,
                FontName = style.FontName,
                FontSize = style.FontSize,
                FontColorArgb = style.FontColorArgb,
                FillColorArgb = style.FillColorArgb,
                Borders = BuildBorders(style.Border),
                HorizontalAlignment = style.HorizontalAlignment,
                VerticalAlignment = style.VerticalAlignment,
                WrapText = style.WrapText,
                TextRotation = style.TextRotation,
                TextIndent = style.TextIndent,
            };
        }

        private static GoogleSheetsCellBorders? BuildBorders(ExcelCellBorderSnapshot? border) {
            if (border == null) {
                return null;
            }

            var left = BuildBorderSide(border.Left);
            var right = BuildBorderSide(border.Right);
            var top = BuildBorderSide(border.Top);
            var bottom = BuildBorderSide(border.Bottom);

            if (left == null && right == null && top == null && bottom == null) {
                return null;
            }

            return new GoogleSheetsCellBorders {
                Left = left,
                Right = right,
                Top = top,
                Bottom = bottom,
            };
        }

        private static GoogleSheetsBorderSide? BuildBorderSide(ExcelBorderSideSnapshot? side) {
            if (side == null) {
                return null;
            }

            if (string.IsNullOrWhiteSpace(side.Style) && string.IsNullOrWhiteSpace(side.ColorArgb)) {
                return null;
            }

            return new GoogleSheetsBorderSide {
                Style = side.Style,
                ColorArgb = side.ColorArgb,
            };
        }

        private static GoogleSheetsHyperlink? BuildHyperlink(ExcelHyperlinkSnapshot? hyperlink) {
            if (hyperlink == null) {
                return null;
            }

            return new GoogleSheetsHyperlink {
                IsExternal = hyperlink.IsExternal,
                Target = hyperlink.Target,
            };
        }

        private static GoogleSheetsComment? BuildComment(ExcelCommentSnapshot? comment) {
            if (comment == null || string.IsNullOrWhiteSpace(comment.Text)) {
                return null;
            }

            return new GoogleSheetsComment {
                Author = string.IsNullOrWhiteSpace(comment.Author) ? null : comment.Author,
                Text = comment.Text,
            };
        }

        private static IReadOnlyList<GoogleSheetsTextFormatRun> BuildTextFormatRuns(IReadOnlyList<ExcelRichTextRun> runs) {
            if (runs == null || runs.Count == 0) return Array.Empty<GoogleSheetsTextFormatRun>();
            var result = new List<GoogleSheetsTextFormatRun>(runs.Count);
            int startIndex = 0;
            foreach (ExcelRichTextRun run in runs) {
                result.Add(new GoogleSheetsTextFormatRun {
                    StartIndex = startIndex,
                    Format = new GoogleSheetsCellStyle {
                        Bold = run.Bold,
                        Italic = run.Italic,
                        Underline = run.Underline,
                        Strikethrough = run.Strikethrough,
                        FontName = run.FontName,
                        FontSize = run.FontSize,
                        FontColorArgb = run.FontColor,
                    },
                });
                startIndex += (run.Text ?? string.Empty).Length;
            }
            return result;
        }
    }
}
