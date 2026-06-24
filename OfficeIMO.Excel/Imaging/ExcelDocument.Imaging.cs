using System.IO;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Exports workbook sheets as PNG or SVG images.
        /// </summary>
        public IReadOnlyList<OfficeImageExportResult> ExportImages(OfficeImageExportFormat format, ExcelWorkbookImageExportOptions? options = null) {
            ExcelWorkbookImageExportOptions resolved = NormalizeWorkbookOptions(options);
            HashSet<string>? selected = resolved.SheetNames == null
                ? null
                : new HashSet<string>(resolved.SheetNames, StringComparer.OrdinalIgnoreCase);
            var results = new List<OfficeImageExportResult>();
            foreach (ExcelSheet sheet in Sheets) {
                if (selected != null && !selected.Contains(sheet.Name)) {
                    continue;
                }

                var sheetOptions = new ExcelWorksheetImageExportOptions {
                    Scale = resolved.Scale,
                    BackgroundColor = resolved.BackgroundColor,
                    GridlineColor = resolved.GridlineColor,
                    ShowGridlines = resolved.ShowGridlines,
                    IncludeHidden = resolved.IncludeHidden,
                    IncludeImages = resolved.IncludeImages,
                    IncludeCharts = resolved.IncludeCharts,
                    IncludeDrawingObjects = resolved.IncludeDrawingObjects,
                    IncludeConditionalFormatting = resolved.IncludeConditionalFormatting,
                    ConditionalFormattingDate = resolved.ConditionalFormattingDate,
                    ShowHyperlinkHints = resolved.ShowHyperlinkHints,
                    ShowCommentBodies = resolved.ShowCommentBodies,
                    DefaultColumnWidthPixels = resolved.DefaultColumnWidthPixels,
                    DefaultRowHeightPixels = resolved.DefaultRowHeightPixels,
                    HeaderFooterDateTime = resolved.HeaderFooterDateTime,
                    UsePrintArea = resolved.UseWorksheetPrintAreas,
                    SplitByManualPageBreaks = resolved.SplitWorksheetsByManualPageBreaks
                };
                results.AddRange(sheet.ExportImages(format, sheetOptions));
            }

            return results.AsReadOnly();
        }

        /// <summary>
        /// Saves workbook sheets as PNG files in a folder.
        /// </summary>
        public IReadOnlyList<OfficeImageExportResult> SaveAsImages(string folderPath, ExcelWorkbookImageExportOptions? options = null) =>
            SaveAsImages(folderPath, OfficeImageExportFormat.Png, options);

        /// <summary>
        /// Saves workbook sheets as image files in a folder.
        /// </summary>
        public IReadOnlyList<OfficeImageExportResult> SaveAsImages(string folderPath, OfficeImageExportFormat format, ExcelWorkbookImageExportOptions? options = null) {
            if (string.IsNullOrWhiteSpace(folderPath)) {
                throw new ArgumentException("Output folder cannot be null or whitespace.", nameof(folderPath));
            }

            string fullFolder = Path.GetFullPath(folderPath);
            Directory.CreateDirectory(fullFolder);
            IReadOnlyList<OfficeImageExportResult> results = ExportImages(format, options);
            string extension = format == OfficeImageExportFormat.Svg ? ".svg" : ".png";
            var usedNames = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < results.Count; i++) {
                OfficeImageExportResult result = results[i];
                string name = string.IsNullOrWhiteSpace(result.Name) ? "sheet-" + (i + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) : result.Name!;
                string path = Path.Combine(fullFolder, GetUniqueFileName(SanitizeFileName(name), extension, usedNames));
                File.WriteAllBytes(path, result.Bytes);
            }

            return results;
        }

        private static ExcelWorkbookImageExportOptions NormalizeWorkbookOptions(ExcelWorkbookImageExportOptions? options) {
            ExcelWorkbookImageExportOptions source = options ?? new ExcelWorkbookImageExportOptions();
            ExcelWorkbookImageExportOptions resolved = new ExcelWorkbookImageExportOptions {
                Scale = source.Scale,
                BackgroundColor = source.BackgroundColor,
                GridlineColor = source.GridlineColor,
                ShowGridlines = source.ShowGridlines,
                IncludeHidden = source.IncludeHidden,
                IncludeImages = source.IncludeImages,
                IncludeCharts = source.IncludeCharts,
                IncludeDrawingObjects = source.IncludeDrawingObjects,
                IncludeConditionalFormatting = source.IncludeConditionalFormatting,
                ConditionalFormattingDate = source.ConditionalFormattingDate ?? DateTime.Today,
                ShowHyperlinkHints = source.ShowHyperlinkHints,
                ShowCommentBodies = source.ShowCommentBodies,
                DefaultColumnWidthPixels = source.DefaultColumnWidthPixels,
                DefaultRowHeightPixels = source.DefaultRowHeightPixels,
                SheetNames = source.SheetNames,
                HeaderFooterDateTime = source.HeaderFooterDateTime ?? DateTime.Now,
                UseWorksheetPrintAreas = source.UseWorksheetPrintAreas,
                SplitWorksheetsByManualPageBreaks = source.SplitWorksheetsByManualPageBreaks
            };
            if (resolved.Scale <= 0D || double.IsNaN(resolved.Scale) || double.IsInfinity(resolved.Scale)) {
                throw new ArgumentOutOfRangeException(nameof(options), "Scale must be a finite positive number.");
            }

            return resolved;
        }

        private static string SanitizeFileName(string name) {
            char[] invalid = Path.GetInvalidFileNameChars();
            var chars = name.ToCharArray();
            for (int i = 0; i < chars.Length; i++) {
                if (Array.IndexOf(invalid, chars[i]) >= 0) {
                    chars[i] = '_';
                }
            }

            return new string(chars).Trim();
        }

        private static string GetUniqueFileName(string baseName, string extension, Dictionary<string, int> usedNames) {
            if (string.IsNullOrWhiteSpace(baseName)) {
                baseName = "sheet";
            }

            if (!usedNames.TryGetValue(baseName, out int count)) {
                usedNames[baseName] = 1;
                return baseName + extension;
            }

            count++;
            usedNames[baseName] = count;
            return baseName + "-" + count.ToString(System.Globalization.CultureInfo.InvariantCulture) + extension;
        }
    }
}
