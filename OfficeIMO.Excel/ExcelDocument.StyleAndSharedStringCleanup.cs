using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        internal void CleanupStyleAndSharedStringArtifacts(bool save = true) {
            var workbookPart = _workBookPart;

            NormalizeStylesheet(workbookPart);
            CleanupSharedStringArtifacts(workbookPart);

            if (save) {
                workbookPart.Workbook.Save();
            }
        }

        private void NormalizeStylesheet(WorkbookPart workbookPart) {
            var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
            var stylesheet = stylesPart.Stylesheet ??= CreateDefaultStylesheet();

            EnsureStylesheetPrimitives(stylesheet);

            var fonts = stylesheet.Fonts!.Elements<Font>().ToList();
            var fills = stylesheet.Fills!.Elements<Fill>().ToList();
            var borders = stylesheet.Borders!.Elements<Border>().ToList();
            var cellStyleFormats = stylesheet.CellStyleFormats!.Elements<CellFormat>().ToList();
            var cellFormats = stylesheet.CellFormats!.Elements<CellFormat>().ToList();

            uint maxFontId = (uint)Math.Max(0, fonts.Count - 1);
            uint maxFillId = (uint)Math.Max(0, fills.Count - 1);
            uint maxBorderId = (uint)Math.Max(0, borders.Count - 1);
            uint maxCellStyleFormatId = (uint)Math.Max(0, cellStyleFormats.Count - 1);

            bool stylesChanged = false;

            foreach (var cellFormat in cellFormats) {
                if (cellFormat.FontId?.Value > maxFontId) {
                    cellFormat.FontId = 0U;
                    stylesChanged = true;
                }
                if (cellFormat.FillId?.Value > maxFillId) {
                    cellFormat.FillId = 0U;
                    stylesChanged = true;
                }
                if (cellFormat.BorderId?.Value > maxBorderId) {
                    cellFormat.BorderId = 0U;
                    stylesChanged = true;
                }
                if (cellFormat.FormatId?.Value > maxCellStyleFormatId) {
                    cellFormat.FormatId = 0U;
                    stylesChanged = true;
                }
            }

            int cellFormatCount = cellFormats.Count;
            foreach (var worksheetPart in workbookPart.WorksheetParts) {
                bool worksheetChanged = false;
                foreach (var cell in worksheetPart.Worksheet.Descendants<Cell>()) {
                    if (cell.StyleIndex?.Value >= cellFormatCount) {
                        cell.StyleIndex = 0U;
                        worksheetChanged = true;
                    }
                }

                if (worksheetChanged) {
                    worksheetPart.Worksheet.Save();
                }
            }

            if (stylesChanged) {
                stylesPart.Stylesheet.Save();
            }
        }

        private void CleanupSharedStringArtifacts(WorkbookPart workbookPart) {
            SharedStringTablePart? sharedStringTablePart = workbookPart.SharedStringTablePart;
            bool sawSharedStringCell = false;
            int sharedStringCellCount = 0;
            int maxSharedStringIndex = -1;
            bool sharedStringTableChanged = false;

            foreach (var worksheetPart in workbookPart.WorksheetParts) {
                bool worksheetChanged = false;
                foreach (var cell in worksheetPart.Worksheet.Descendants<Cell>()) {
                    if (cell.DataType?.Value != CellValues.SharedString) {
                        continue;
                    }

                    string rawValue = cell.CellValue?.Text ?? cell.InnerText ?? string.Empty;
                    if (!int.TryParse(rawValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int sharedStringIndex) || sharedStringIndex < 0) {
                        cell.DataType = CellValues.InlineString;
                        cell.CellValue = null;
                        cell.InlineString = new InlineString(new Text(rawValue));
                        worksheetChanged = true;
                        continue;
                    }

                    sawSharedStringCell = true;
                    sharedStringCellCount++;
                    if (sharedStringIndex > maxSharedStringIndex) {
                        maxSharedStringIndex = sharedStringIndex;
                    }
                }

                if (worksheetChanged) {
                    worksheetPart.Worksheet.Save();
                }
            }

            if (sharedStringTablePart == null && !sawSharedStringCell) {
                return;
            }

            if (sharedStringTablePart == null) {
                sharedStringTablePart = workbookPart.AddNewPart<SharedStringTablePart>();
                sharedStringTablePart.SharedStringTable = new SharedStringTable();
                _sharedStringTablePart = sharedStringTablePart;
                sharedStringTableChanged = true;
            }

            var sharedStringTable = sharedStringTablePart.SharedStringTable ??= new SharedStringTable();
            int itemCount = sharedStringTable.Elements<SharedStringItem>().Count();
            while (itemCount <= maxSharedStringIndex) {
                sharedStringTable.AppendChild(new SharedStringItem(new Text(string.Empty)));
                itemCount++;
                sharedStringTableChanged = true;
            }

            uint expectedCount = (uint)Math.Max(sharedStringCellCount, itemCount);
            uint expectedUniqueCount = (uint)itemCount;
            if (sharedStringTable.Count?.Value != expectedCount) {
                sharedStringTable.Count = expectedCount;
                sharedStringTableChanged = true;
            }
            if (sharedStringTable.UniqueCount?.Value != expectedUniqueCount) {
                sharedStringTable.UniqueCount = expectedUniqueCount;
                sharedStringTableChanged = true;
            }

            if (sharedStringTableChanged) {
                sharedStringTable.Save();
                _sharedStringCache.Clear();
            }
        }

        private static void EnsureStylesheetPrimitives(Stylesheet stylesheet) {
            if (stylesheet.Fonts == null || !stylesheet.Fonts.Elements<Font>().Any()) {
                stylesheet.Fonts = new Fonts(new Font());
            }
            stylesheet.Fonts.Count = (uint)stylesheet.Fonts.Count();

            if (stylesheet.Fills == null) {
                stylesheet.Fills = new Fills();
            }

            bool hasNoneFill = stylesheet.Fills.Elements<Fill>().Any(fill => fill.PatternFill?.PatternType?.Value == PatternValues.None);
            bool hasGray125Fill = stylesheet.Fills.Elements<Fill>().Any(fill => fill.PatternFill?.PatternType?.Value == PatternValues.Gray125);
            if (!hasNoneFill) {
                stylesheet.Fills.Append(new Fill(new PatternFill { PatternType = PatternValues.None }));
            }
            if (!hasGray125Fill) {
                stylesheet.Fills.Append(new Fill(new PatternFill { PatternType = PatternValues.Gray125 }));
            }
            stylesheet.Fills.Count = (uint)stylesheet.Fills.Count();

            if (stylesheet.Borders == null || !stylesheet.Borders.Elements<Border>().Any()) {
                stylesheet.Borders = new Borders(new Border());
            }
            stylesheet.Borders.Count = (uint)stylesheet.Borders.Count();

            if (stylesheet.CellStyleFormats == null || !stylesheet.CellStyleFormats.Elements<CellFormat>().Any()) {
                stylesheet.CellStyleFormats = new CellStyleFormats(new CellFormat());
            }
            stylesheet.CellStyleFormats.Count = (uint)stylesheet.CellStyleFormats.Count();

            if (stylesheet.CellFormats == null || !stylesheet.CellFormats.Elements<CellFormat>().Any()) {
                stylesheet.CellFormats = new CellFormats(new CellFormat());
            }
            stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();

            if (stylesheet.NumberingFormats != null) {
                stylesheet.NumberingFormats.Count = (uint)stylesheet.NumberingFormats.Count();
            }
        }
    }
}
