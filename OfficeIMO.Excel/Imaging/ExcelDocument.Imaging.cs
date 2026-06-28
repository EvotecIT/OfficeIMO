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

                if (selected == null && sheet.Hidden && !resolved.IncludeHiddenSheets) {
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
            new ExcelWorkbookImageExportBuilder(this, options).AsPng().Save(folderPath);

        /// <summary>
        /// Saves workbook sheets as image files in a folder.
        /// </summary>
        public IReadOnlyList<OfficeImageExportResult> SaveAsImages(string folderPath, OfficeImageExportFormat format, ExcelWorkbookImageExportOptions? options = null) {
            ExcelWorkbookImageExportBuilder builder = new ExcelWorkbookImageExportBuilder(this, options);
            if (format == OfficeImageExportFormat.Svg) {
                builder.AsSvg();
            } else {
                builder.AsPng();
            }

            return builder.Save(folderPath);
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
                IncludeHiddenSheets = source.IncludeHiddenSheets,
                HeaderFooterDateTime = source.HeaderFooterDateTime ?? DateTime.Now,
                UseWorksheetPrintAreas = source.UseWorksheetPrintAreas,
                SplitWorksheetsByManualPageBreaks = source.SplitWorksheetsByManualPageBreaks
            };
            OfficeImageExportOptions.ValidateScale(resolved.Scale, nameof(options));

            return resolved;
        }
    }
}
