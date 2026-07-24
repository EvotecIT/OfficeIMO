using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        internal void CleanupStyleAndSharedStringArtifacts(bool save = true) {
            var workbookPart = WorkbookPartRoot;

            NormalizeStylesheet(workbookPart);
            CleanupSharedStringArtifacts(workbookPart);

            if (save) {
                WorkbookRoot.Save();
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
                var worksheet = worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet is missing.");
                bool worksheetChanged = false;
                foreach (var cell in worksheet.Descendants<Cell>()) {
                    if (cell.StyleIndex?.Value >= cellFormatCount) {
                        cell.StyleIndex = 0U;
                        worksheetChanged = true;
                    }
                }

                if (worksheetChanged) {
                    worksheet.Save();
                }
            }

            if (stylesChanged) {
                stylesPart.Stylesheet.Save();
            }
        }

        private void CleanupSharedStringArtifacts(WorkbookPart workbookPart) {
            SharedStringTablePart? sharedStringTablePart = workbookPart.SharedStringTablePart;
            var sharedStringTable = sharedStringTablePart?.SharedStringTable;
            int itemCount = sharedStringTable?.Elements<SharedStringItem>().Count() ?? 0;
            int sharedStringCellCount = 0;
            bool sharedStringTableChanged = false;

            foreach (var worksheetPart in workbookPart.WorksheetParts) {
                var worksheet = worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet is missing.");
                bool worksheetChanged = false;
                var sheetData = worksheet.GetFirstChild<SheetData>();
                if (sheetData != null) {
                    foreach (var row in sheetData.Elements<Row>()) {
                        foreach (var cell in row.Elements<Cell>()) {
                            if (!IsSharedStringCell(cell)) {
                                continue;
                            }

                            string rawValue = cell.CellValue?.Text ?? cell.InnerText ?? string.Empty;
                            if (!TryParseSharedStringIndex(rawValue, out int sharedStringIndex)
                                || sharedStringIndex >= itemCount) {
                                cell.DataType = CellValues.InlineString;
                                cell.RemoveAllChildren<CellValue>();
                                cell.RemoveAllChildren<InlineString>();
                                cell.AppendChild(new InlineString(new Text(rawValue)));
                                worksheetChanged = true;
                                continue;
                            }

                            sharedStringCellCount++;
                        }
                    }
                }

                if (worksheetChanged) {
                    worksheet.Save();
                }
            }

            if (sharedStringTablePart == null) {
                return;
            }

            sharedStringTable ??= sharedStringTablePart.SharedStringTable = new SharedStringTable();

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
                _sharedStringTableCount = itemCount;
            }
        }

        private static bool TryParseSharedStringIndex(string text, out int index) {
            index = 0;
            if (string.IsNullOrEmpty(text)) {
                return false;
            }

            int parsed = 0;
            for (int i = 0; i < text.Length; i++) {
                int digit = text[i] - '0';
                if ((uint)digit > 9U || parsed > (int.MaxValue - digit) / 10) {
                    return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out index)
                        && index >= 0;
                }

                parsed = (parsed * 10) + digit;
            }

            index = parsed;
            return true;
        }

        private static bool IsSharedStringCell(Cell cell) {
            var dataType = cell.DataType;
            return dataType?.Value == CellValues.SharedString
                || string.Equals(dataType?.InnerText, "s", StringComparison.Ordinal);
        }

        private static void EnsureStylesheetPrimitives(Stylesheet stylesheet) {
            if (stylesheet.Fonts == null || !stylesheet.Fonts.Elements<Font>().Any()) {
                stylesheet.Fonts = new Fonts(new Font(new FontSize { Val = 11D }, new FontName { Val = "Calibri" }));
            } else {
                var defaultFont = stylesheet.Fonts.Elements<Font>().First();
                defaultFont.FontSize ??= new FontSize { Val = 11D };
                defaultFont.FontName ??= new FontName { Val = "Calibri" };
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
                stylesheet.CellStyleFormats = new CellStyleFormats(new CellFormat {
                    NumberFormatId = 0U,
                    FontId = 0U,
                    FillId = 0U,
                    BorderId = 0U
                });
            }
            stylesheet.CellStyleFormats.Count = (uint)stylesheet.CellStyleFormats.Count();

            if (stylesheet.CellFormats == null || !stylesheet.CellFormats.Elements<CellFormat>().Any()) {
                stylesheet.CellFormats = new CellFormats(new CellFormat {
                    NumberFormatId = 0U,
                    FontId = 0U,
                    FillId = 0U,
                    BorderId = 0U,
                    FormatId = 0U
                });
            }
            stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();

            if (stylesheet.CellStyles == null || !stylesheet.CellStyles.Elements<CellStyle>().Any()) {
                stylesheet.CellStyles = new CellStyles(new CellStyle {
                    Name = "Normal",
                    FormatId = 0U,
                    BuiltinId = 0U
                });
            }
            stylesheet.CellStyles.Count = (uint)stylesheet.CellStyles.Count();

            stylesheet.DifferentialFormats ??= new DifferentialFormats();
            stylesheet.DifferentialFormats.Count = (uint)stylesheet.DifferentialFormats.Count();

            stylesheet.TableStyles ??= new TableStyles {
                DefaultTableStyle = "TableStyleMedium2",
                DefaultPivotStyle = "PivotStyleLight16"
            };
            stylesheet.TableStyles.Count = (uint)stylesheet.TableStyles.Count();

            if (stylesheet.NumberingFormats != null) {
                stylesheet.NumberingFormats.Count = (uint)stylesheet.NumberingFormats.Count();
            }
        }
    }
}
