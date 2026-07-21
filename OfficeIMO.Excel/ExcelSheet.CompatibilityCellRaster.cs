using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel;

public partial class ExcelSheet {
    internal void ApplyCompatibilityCellRaster(
        IReadOnlyList<string> argbPixels,
        int width,
        int height) {
        if (argbPixels == null) throw new ArgumentNullException(nameof(argbPixels));
        if (width <= 0 || width > 256) throw new ArgumentOutOfRangeException(nameof(width));
        if (height <= 0 || height > 65536) throw new ArgumentOutOfRangeException(nameof(height));
        if (argbPixels.Count != checked(width * height)) {
            throw new ArgumentException("Cell-raster pixel count does not match its dimensions.", nameof(argbPixels));
        }

        WriteLock(() => {
            WorkbookStylesPart stylesPart = _excelDocument.WorkbookPartRoot.WorkbookStylesPart
                ?? _excelDocument.WorkbookPartRoot.AddNewPart<WorkbookStylesPart>();
            Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            EnsureDefaultStylePrimitives(stylesheet);
            var styleByColor = new Dictionary<string, uint>(StringComparer.OrdinalIgnoreCase);
            foreach (string color in argbPixels.Distinct(StringComparer.OrdinalIgnoreCase)) {
                var fill = new Fill(new PatternFill {
                    PatternType = PatternValues.Solid,
                    ForegroundColor = new ForegroundColor { Rgb = color },
                    BackgroundColor = new BackgroundColor { Rgb = color }
                });
                uint fillId = GetOrCreateFill(stylesheet, fill);
                var format = GetBaseCellFormat(stylesheet, 0U);
                format.FillId = fillId;
                format.ApplyFill = true;
                styleByColor[color] = AppendOrReuseCellFormat(stylesheet, format);
            }

            Worksheet worksheet = WorksheetRoot;
            worksheet.GetFirstChild<Columns>()?.Remove();
            var columns = new Columns(new Column {
                Min = 1U,
                Max = (uint)width,
                Width = 0.42D,
                CustomWidth = true
            });
            SheetData sheetData = worksheet.GetFirstChild<SheetData>() ?? new SheetData();
            if (sheetData.Parent == null) worksheet.Append(sheetData);
            sheetData.RemoveAllChildren<Row>();
            worksheet.InsertBefore(columns, sheetData);

            for (int rowIndex = 0; rowIndex < height; rowIndex++) {
                var row = new Row {
                    RowIndex = (uint)(rowIndex + 1),
                    Height = 2.25D,
                    CustomHeight = true
                };
                int rowOffset = rowIndex * width;
                for (int columnIndex = 0; columnIndex < width; columnIndex++) {
                    string color = argbPixels[rowOffset + columnIndex];
                    row.Append(new Cell {
                        CellReference = A1.ColumnIndexToLetters(columnIndex + 1) + (rowIndex + 1).ToString(System.Globalization.CultureInfo.InvariantCulture),
                        StyleIndex = styleByColor[color]
                    });
                }
                sheetData.Append(row);
            }

            SheetView? view = worksheet.GetFirstChild<SheetViews>()?.Elements<SheetView>().FirstOrDefault();
            if (view != null) view.ShowGridLines = false;
            stylesheet.Save();
            worksheet.Save();
        });
    }
}
