using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private ExcelSheet CopyWorkSheetFromValues(ExcelDocument sourceDocument, string sourceSheetName, string newSheetName, SheetNameValidationMode validationMode) {
            ExcelSheet sourceSheet = sourceDocument.GetSheet(sourceSheetName);
            string usedRange = sourceSheet.GetUsedRangeA1();
            var (startRow, startColumn, _, _) = A1.ParseRange(usedRange);
            object?[,] values;
            using (var reader = sourceDocument.CreateReader()) {
                values = reader.GetSheet(sourceSheet.Name).ReadRange(usedRange);
            }

            ExcelSheet targetSheet = AddWorkSheet(newSheetName, validationMode);
            for (int rowOffset = 0; rowOffset < values.GetLength(0); rowOffset++) {
                for (int columnOffset = 0; columnOffset < values.GetLength(1); columnOffset++) {
                    object? value = values[rowOffset, columnOffset];
                    if (value == null) continue;
                    targetSheet.CellValue(startRow + rowOffset, startColumn + columnOffset, value);
                }
            }

            CopyWorksheetTables(sourceSheet.WorksheetPart, targetSheet.WorksheetPart);
            targetSheet.WorksheetPart.Worksheet!.Save();
            return targetSheet;
        }

        private ExcelSheet CopyWorkSheetFromPackage(ExcelDocument sourceDocument, string sourceSheetName, string newSheetName, SheetNameValidationMode validationMode) {
            ExcelSheet sourceSheet = sourceDocument.GetSheet(sourceSheetName);

            return Locking.ExecuteWrite(EnsureLock(), () => {
                string validatedName = ValidateOrSanitizeSheetName(newSheetName, validationMode, currentSheetName: null);
                WorksheetPart sourcePart = sourceSheet.WorksheetPart;
                WorksheetPart copiedPart = WorkbookPartRoot.AddNewPart<WorksheetPart>();
                copiedPart.Worksheet = (Worksheet)sourcePart.Worksheet!.CloneNode(true);
                RewriteSharedStringCellsToInlineStrings(copiedPart.Worksheet, sourceDocument.WorkbookPartRoot.SharedStringTablePart);
                RemapCopiedWorksheetStyles(sourceDocument.WorkbookPartRoot.WorkbookStylesPart?.Stylesheet, WorkbookPartRoot, copiedPart.Worksheet);
                RemoveRelationshipBackedWorksheetFeatures(copiedPart.Worksheet);
                CopyWorksheetTables(sourcePart, copiedPart);
                copiedPart.Worksheet.Save();

                Sheet sheet = AppendWorksheetElement(copiedPart, validatedName);
                MarkSheetCacheDirty();
                WorkbookRoot.Save();
                return new ExcelSheet(this, _spreadSheetDocument, sheet);
            });
        }

        private static void RewriteSharedStringCellsToInlineStrings(Worksheet worksheet, SharedStringTablePart? sharedStringTablePart) {
            SharedStringTable? table = sharedStringTablePart?.SharedStringTable;
            if (table == null) {
                return;
            }

            var sharedItems = table.Elements<SharedStringItem>().ToList();
            foreach (Cell cell in worksheet.Descendants<Cell>()) {
                if (cell.DataType?.Value != CellValues.SharedString || cell.CellValue?.Text == null) {
                    continue;
                }

                if (!int.TryParse(cell.CellValue.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index) ||
                    index < 0 ||
                    index >= sharedItems.Count) {
                    continue;
                }

                SharedStringItem item = sharedItems[index];
                var inline = new InlineString();
                foreach (OpenXmlElement child in item.ChildElements) {
                    inline.Append(child.CloneNode(true));
                }

                cell.CellValue = null;
                cell.InlineString = inline;
                cell.DataType = CellValues.InlineString;
            }
        }

        private static void RemapCopiedWorksheetStyles(Stylesheet? sourceStylesheet, WorkbookPart targetWorkbookPart, Worksheet worksheet) {
            if (sourceStylesheet?.CellFormats == null) {
                return;
            }

            WorkbookStylesPart targetStylesPart = targetWorkbookPart.WorkbookStylesPart ?? targetWorkbookPart.AddNewPart<WorkbookStylesPart>();
            Stylesheet targetStylesheet = targetStylesPart.Stylesheet ??= CreateDefaultStylesheet();
            EnsureStylesheetPrimitives(targetStylesheet);

            Dictionary<uint, uint> numberingMap = AppendNumberingFormats(sourceStylesheet, targetStylesheet);
            Dictionary<uint, uint> fontMap = AppendStyleElements<Fonts, Font>(sourceStylesheet.Fonts, targetStylesheet.Fonts!, (container, count) => container.Count = count);
            Dictionary<uint, uint> fillMap = AppendStyleElements<Fills, Fill>(sourceStylesheet.Fills, targetStylesheet.Fills!, (container, count) => container.Count = count);
            Dictionary<uint, uint> borderMap = AppendStyleElements<Borders, Border>(sourceStylesheet.Borders, targetStylesheet.Borders!, (container, count) => container.Count = count);
            Dictionary<uint, uint> styleFormatMap = AppendCellStyleFormats(sourceStylesheet.CellStyleFormats, targetStylesheet.CellStyleFormats!, numberingMap, fontMap, fillMap, borderMap);
            Dictionary<uint, uint> cellFormatMap = AppendCellFormats(sourceStylesheet.CellFormats, targetStylesheet.CellFormats!, numberingMap, fontMap, fillMap, borderMap, styleFormatMap);

            foreach (Cell cell in worksheet.Descendants<Cell>()) {
                uint? oldStyleIndex = cell.StyleIndex?.Value;
                if (oldStyleIndex.HasValue && cellFormatMap.TryGetValue(oldStyleIndex.Value, out uint newStyleIndex)) {
                    cell.StyleIndex = newStyleIndex;
                }
            }

            EnsureStylesheetPrimitives(targetStylesheet);
            targetStylesheet.Save();
        }

        private static Dictionary<uint, uint> AppendNumberingFormats(Stylesheet sourceStylesheet, Stylesheet targetStylesheet) {
            var map = new Dictionary<uint, uint>();
            if (sourceStylesheet.NumberingFormats == null) {
                return map;
            }

            NumberingFormats targetFormats = targetStylesheet.NumberingFormats ?? InsertNumberingFormats(targetStylesheet);
            uint nextCustomId = Math.Max(164U, targetFormats.Elements<NumberingFormat>()
                .Select(format => format.NumberFormatId?.Value ?? 0U)
                .DefaultIfEmpty(163U)
                .Max() + 1U);

            foreach (NumberingFormat sourceFormat in sourceStylesheet.NumberingFormats.Elements<NumberingFormat>()) {
                uint oldId = sourceFormat.NumberFormatId?.Value ?? 0U;
                if (oldId < 164U) {
                    map[oldId] = oldId;
                    continue;
                }

                string? sourceCode = sourceFormat.FormatCode?.Value;
                NumberingFormat? existing = targetFormats.Elements<NumberingFormat>()
                    .FirstOrDefault(format => string.Equals(format.FormatCode?.Value, sourceCode, StringComparison.Ordinal));
                if (existing?.NumberFormatId?.Value is uint existingId) {
                    map[oldId] = existingId;
                    continue;
                }

                var copied = (NumberingFormat)sourceFormat.CloneNode(true);
                copied.NumberFormatId = nextCustomId++;
                targetFormats.Append(copied);
                map[oldId] = copied.NumberFormatId!.Value;
            }

            targetFormats.Count = (uint)targetFormats.Elements<NumberingFormat>().Count();
            return map;
        }

        private static NumberingFormats InsertNumberingFormats(Stylesheet stylesheet) {
            var formats = new NumberingFormats();
            if (stylesheet.Fonts != null) {
                stylesheet.InsertBefore(formats, stylesheet.Fonts);
            } else {
                stylesheet.PrependChild(formats);
            }

            return formats;
        }

        private static Dictionary<uint, uint> AppendStyleElements<TContainer, TElement>(TContainer? source, TContainer target, Action<TContainer, uint> setCount)
            where TContainer : OpenXmlCompositeElement
            where TElement : OpenXmlElement {
            var map = new Dictionary<uint, uint>();
            uint offset = (uint)target.Elements<TElement>().Count();
            if (source == null) {
                return map;
            }

            uint index = 0;
            foreach (TElement element in source.Elements<TElement>()) {
                uint newIndex = offset + index;
                target.Append(element.CloneNode(true));
                map[index] = newIndex;
                index++;
            }

            setCount(target, (uint)target.Elements<TElement>().Count());
            return map;
        }

        private static Dictionary<uint, uint> AppendCellFormats(
            CellFormats? source,
            CellFormats target,
            IReadOnlyDictionary<uint, uint> numberingMap,
            IReadOnlyDictionary<uint, uint> fontMap,
            IReadOnlyDictionary<uint, uint> fillMap,
            IReadOnlyDictionary<uint, uint> borderMap,
            IReadOnlyDictionary<uint, uint>? formatMap) {
            var map = new Dictionary<uint, uint>();
            if (source == null) {
                return map;
            }

            uint offset = (uint)target.Elements<CellFormat>().Count();
            uint index = 0;
            foreach (CellFormat sourceFormat in source.Elements<CellFormat>()) {
                uint newIndex = offset + index;
                var copied = (CellFormat)sourceFormat.CloneNode(true);
                copied.NumberFormatId = MapValue(copied.NumberFormatId?.Value, numberingMap);
                copied.FontId = MapValue(copied.FontId?.Value, fontMap);
                copied.FillId = MapValue(copied.FillId?.Value, fillMap);
                copied.BorderId = MapValue(copied.BorderId?.Value, borderMap);
                if (formatMap != null) {
                    copied.FormatId = MapValue(copied.FormatId?.Value, formatMap);
                }

                target.Append(copied);
                map[index] = newIndex;
                index++;
            }

            target.Count = (uint)target.Elements<CellFormat>().Count();
            return map;
        }

        private static Dictionary<uint, uint> AppendCellStyleFormats(
            CellStyleFormats? source,
            CellStyleFormats target,
            IReadOnlyDictionary<uint, uint> numberingMap,
            IReadOnlyDictionary<uint, uint> fontMap,
            IReadOnlyDictionary<uint, uint> fillMap,
            IReadOnlyDictionary<uint, uint> borderMap) {
            var map = new Dictionary<uint, uint>();
            if (source == null) {
                return map;
            }

            uint offset = (uint)target.Elements<CellFormat>().Count();
            uint index = 0;
            foreach (CellFormat sourceFormat in source.Elements<CellFormat>()) {
                uint newIndex = offset + index;
                var copied = (CellFormat)sourceFormat.CloneNode(true);
                copied.NumberFormatId = MapValue(copied.NumberFormatId?.Value, numberingMap);
                copied.FontId = MapValue(copied.FontId?.Value, fontMap);
                copied.FillId = MapValue(copied.FillId?.Value, fillMap);
                copied.BorderId = MapValue(copied.BorderId?.Value, borderMap);

                target.Append(copied);
                map[index] = newIndex;
                index++;
            }

            target.Count = (uint)target.Elements<CellFormat>().Count();
            return map;
        }

        private static uint MapValue(uint? value, IReadOnlyDictionary<uint, uint> map) {
            if (!value.HasValue) {
                return 0U;
            }

            return map.TryGetValue(value.Value, out uint mapped) ? mapped : value.Value;
        }
    }
}
