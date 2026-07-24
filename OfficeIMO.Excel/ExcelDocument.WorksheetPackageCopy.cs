using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private ExcelSheet CopyWorksheetFromValues(ExcelDocument sourceDocument, string sourceSheetName, string newSheetName, SheetNameValidationMode validationMode) {
            ExcelSheet sourceSheet = sourceDocument.GetSheet(sourceSheetName);
            string usedRange = sourceSheet.GetUsedRangeA1();
            var (startRow, startColumn, _, _) = A1.ParseRange(usedRange);
            object?[,] values;
            using (var reader = sourceDocument.CreateReader()) {
                values = reader.GetSheet(sourceSheet.Name).ReadRange(usedRange);
            }

            ExcelSheet targetSheet = AddWorksheet(newSheetName, validationMode);
            for (int rowOffset = 0; rowOffset < values.GetLength(0); rowOffset++) {
                for (int columnOffset = 0; columnOffset < values.GetLength(1); columnOffset++) {
                    object? value = values[rowOffset, columnOffset];
                    if (value == null) continue;
                    targetSheet.CellValue(startRow + rowOffset, startColumn + columnOffset, value);
                }
            }

            CopyWorksheetTables(
                sourceSheet.WorksheetPart,
                targetSheet.WorksheetPart,
                rewriteCopiedTableReferences: true,
                preserveTableFormulas: false);
            targetSheet.WorksheetPart.Worksheet!.Save();
            return targetSheet;
        }

        private ExcelSheet CopyWorksheetFromPackage(
            ExcelDocument sourceDocument,
            string sourceSheetName,
            string newSheetName,
            SheetNameValidationMode validationMode,
            ExcelWorksheetCopyOptions options,
            DefinedNameCopyBudget definedNameBudget) {
            return CopyWorksheetFromPackage(
                sourceDocument,
                sourceSheetName,
                newSheetName,
                validationMode,
                rewriteCopiedReferences: true,
                copyReferencedDefinedNames: true,
                options.CopyExternalWorkbookReferences,
                definedNameBudget).Sheet;
        }

        private WorksheetPackageCopyResult CopyWorksheetFromPackage(
            ExcelDocument sourceDocument,
            string sourceSheetName,
            string newSheetName,
            SheetNameValidationMode validationMode,
            bool rewriteCopiedReferences,
            bool copyReferencedDefinedNames,
            bool copyExternalWorkbookReferences,
            DefinedNameCopyBudget definedNameBudget) {
            MaterializeDeferredDataSetImport();
            sourceDocument.MaterializeDeferredDataSetImport();
            ExcelSheet sourceSheet = sourceDocument.GetSheet(sourceSheetName);
            if (!copyExternalWorkbookReferences
                && sourceDocument.WorkbookRoot.ExternalReferences?.Elements<ExternalReference>().Any() == true) {
                throw new InvalidOperationException(
                    "Package-mode worksheet copy refuses external-workbook references unless CopyExternalWorkbookReferences is enabled explicitly.");
            }

            return Locking.ExecuteWrite(EnsureLock(), () => {
                string validatedName = ValidateOrSanitizeSheetName(newSheetName, validationMode, currentSheetName: null);
                WorksheetPart sourcePart = sourceSheet.WorksheetPart;
                if (TryCopyWorksheetPartGraphToEmptyWorkbook(sourceDocument, sourceSheet, validatedName, out WorksheetPackageCopyResult? directCopy)) {
                    return directCopy!;
                }

                bool adoptedSourceIndexes = CanAdoptSourceWorkbookIndexParts(sourceDocument);
                if (adoptedSourceIndexes) {
                    AdoptSourceWorkbookIndexParts(sourceDocument);
                }

                WorksheetPart copiedPart = WorkbookPartRoot.AddNewPart<WorksheetPart>();
                copiedPart.Worksheet = (Worksheet)sourcePart.Worksheet!.CloneNode(true);
                RewriteSharedStringCellsToInlineStrings(copiedPart.Worksheet, sourceDocument.WorkbookPartRoot.SharedStringTablePart);
                if (!adoptedSourceIndexes) {
                    WorksheetStyleCopyMap styleMap = RemapCopiedWorksheetStyles(sourceDocument.WorkbookPartRoot, WorkbookPartRoot, copiedPart.Worksheet);
                    RemapCopiedWorksheetConditionalFormats(copiedPart.Worksheet, styleMap.DifferentialFormats);
                    ConvertCopiedWorksheetDateSerials(copiedPart.Worksheet, sourceDocument.DateSystem, DateSystem);
                }

                StripCopiedCellMetadataReferences(copiedPart.Worksheet);
                RemoveRelationshipBackedWorksheetFeatures(copiedPart.Worksheet);
                IReadOnlyDictionary<int, int> externalReferenceMap = copyExternalWorkbookReferences
                    ? CopyExternalWorkbookReferencesFromSource(sourceDocument)
                    : new Dictionary<int, int>();
                IReadOnlyDictionary<string, string> tableNameMap = CopyWorksheetTables(sourcePart, copiedPart);
                bool copiedFormulas = copiedPart.Worksheet.Descendants<CellFormula>()
                    .Any(formula => !string.IsNullOrWhiteSpace(formula.Text));

                Sheet sheet = AppendWorksheetElement(copiedPart, validatedName);
                var targetSheet = new ExcelSheet(this, _spreadSheetDocument, sheet);
                MarkSheetCacheDirty();
                var sheetNameMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
                    [sourceSheet.Name] = targetSheet.Name
                };
                if (rewriteCopiedReferences) {
                    RewriteCopiedWorksheetReferences(copiedPart, sheetNameMap, tableNameMap);
                }

                if (externalReferenceMap.Count > 0) {
                    RewriteCopiedWorksheetExternalReferences(copiedPart, externalReferenceMap);
                }

                if (copyReferencedDefinedNames) {
                    CopyReferencedDefinedNamesFromSource(
                        sourceDocument,
                        targetSheet,
                        sheetNameMap,
                        tableNameMap,
                        externalReferenceMap,
                        definedNameBudget);
                }

                copiedPart.Worksheet.Save();
                if (copiedFormulas) {
                    MarkRequiresSavePreflight();
                } else {
                    MarkPackageDirty();
                }
                WorkbookRoot.Save();
                return new WorksheetPackageCopyResult(targetSheet, tableNameMap, externalReferenceMap);
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

        private IReadOnlyDictionary<int, int> CopyExternalWorkbookReferencesFromSource(ExcelDocument sourceDocument) {
            if (ReferenceEquals(sourceDocument.WorkbookPartRoot, WorkbookPartRoot)) {
                return new Dictionary<int, int>();
            }

            ExternalReferences? sourceReferences = sourceDocument.WorkbookPartRoot.Workbook?.ExternalReferences;
            if (sourceReferences == null) {
                return new Dictionary<int, int>();
            }

            ExternalReferences targetReferences = WorkbookRoot.ExternalReferences ??= new ExternalReferences();
            int targetIndex = targetReferences.Elements<ExternalReference>().Count() + 1;
            int sourceIndex = 1;
            var referenceMap = new Dictionary<int, int>();

            foreach (ExternalReference sourceReference in sourceReferences.Elements<ExternalReference>()) {
                string? sourceRelationshipId = sourceReference.Id?.Value;
                if (string.IsNullOrWhiteSpace(sourceRelationshipId)) {
                    sourceIndex++;
                    continue;
                }

                OpenXmlPart sourcePart;
                try {
                    sourcePart = sourceDocument.WorkbookPartRoot.GetPartById(sourceRelationshipId!);
                } catch (ArgumentOutOfRangeException) {
                    sourceIndex++;
                    continue;
                }

                OpenXmlPart copiedPart = WorkbookPartRoot.AddPart(sourcePart);
                string targetRelationshipId = WorkbookPartRoot.GetIdOfPart(copiedPart);
                targetReferences.Append(new ExternalReference { Id = targetRelationshipId });
                referenceMap[sourceIndex] = targetIndex;
                targetIndex++;
                sourceIndex++;
            }

            if (referenceMap.Count > 0) {
                WorkbookRoot.Save();
            }

            return referenceMap;
        }

        private void ConvertCopiedWorksheetDateSerials(Worksheet worksheet, ExcelDateSystem sourceDateSystem, ExcelDateSystem targetDateSystem) {
            if (sourceDateSystem == targetDateSystem || _spreadSheetDocument == null) {
                return;
            }

            StylesCache styles = StylesCache.Build(_spreadSheetDocument);
            if (!styles.HasDateStyles) {
                return;
            }

            double offset = targetDateSystem == ExcelDateSystem.NineteenFour
                ? -ExcelDateSystemConverter.Date1904OffsetDays
                : ExcelDateSystemConverter.Date1904OffsetDays;
            Dictionary<int, uint> rowStyleIndexes = BuildRowStyleIndexMap(worksheet);
            (int Min, int Max, uint StyleIndex)[] columnStyleIndexes = BuildColumnStyleIndexRanges(worksheet);

            foreach (Cell cell in worksheet.Descendants<Cell>()) {
                if (cell.CellFormula != null || !IsDateSystemShiftCell(cell, styles, GetEffectiveStyleIndex(cell, rowStyleIndexes, columnStyleIndexes))) {
                    continue;
                }

                string? text = cell.CellValue?.Text;
                if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double serial)) {
                    continue;
                }

                cell.CellValue = new CellValue(InvariantNumberText.Get(serial + offset));
            }
        }

        private static Dictionary<int, uint> BuildRowStyleIndexMap(Worksheet worksheet) {
            var rowStyles = new Dictionary<int, uint>();
            foreach (Row row in worksheet.Descendants<Row>()) {
                IList<OpenXmlAttribute> attributes = row.GetAttributes();
                string rowIndexText = GetUnqualifiedAttributeValue(attributes, "r");
                string styleIndexText = GetUnqualifiedAttributeValue(attributes, "s");
                if (!int.TryParse(rowIndexText, NumberStyles.None, CultureInfo.InvariantCulture, out int rowIndex)
                    || rowIndex < 1
                    || !uint.TryParse(styleIndexText, NumberStyles.None, CultureInfo.InvariantCulture, out uint styleIndex)
                    || rowStyles.ContainsKey(rowIndex)) {
                    continue;
                }

                rowStyles.Add(rowIndex, styleIndex);
            }

            return rowStyles;
        }

        private static string GetUnqualifiedAttributeValue(IList<OpenXmlAttribute> attributes, string localName) {
            for (int i = 0; i < attributes.Count; i++) {
                OpenXmlAttribute attribute = attributes[i];
                if (attribute.LocalName == localName && string.IsNullOrEmpty(attribute.NamespaceUri)) {
                    return attribute.Value ?? string.Empty;
                }
            }

            return string.Empty;
        }

        private static (int Min, int Max, uint StyleIndex)[] BuildColumnStyleIndexRanges(Worksheet worksheet) {
            return worksheet.Elements<Columns>()
                .SelectMany(columns => columns.Elements<Column>())
                .Where(column => column.Style != null)
                .Select(column => (
                    Min: (int)(column.Min?.Value ?? 1U),
                    Max: (int)(column.Max?.Value ?? column.Min?.Value ?? 1U),
                    StyleIndex: column.Style!.Value))
                .ToArray();
        }

        private static uint? GetEffectiveStyleIndex(
            Cell cell,
            IReadOnlyDictionary<int, uint> rowStyleIndexes,
            IReadOnlyList<(int Min, int Max, uint StyleIndex)> columnStyleIndexes) {
            if (cell.StyleIndex?.Value is uint cellStyleIndex) {
                return cellStyleIndex;
            }

            string? reference = cell.CellReference?.Value;
            if (string.IsNullOrWhiteSpace(reference)) {
                return null;
            }

            (int row, int column) = A1.ParseCellRef(reference!);
            if (rowStyleIndexes.TryGetValue(row, out uint rowStyleIndex)) {
                return rowStyleIndex;
            }

            foreach (var columnStyle in columnStyleIndexes) {
                if (column >= columnStyle.Min && column <= columnStyle.Max) {
                    return columnStyle.StyleIndex;
                }
            }

            return null;
        }

        private static bool IsDateSystemShiftCell(Cell cell, StylesCache styles, uint? styleIndex) {
            if (!styleIndex.HasValue || !styles.IsDateSystemShiftStyle(styleIndex.Value)) {
                return false;
            }

            CellValues? dataType = cell.DataType?.Value;
            return dataType == null || dataType == CellValues.Number;
        }

        private static void StripCopiedCellMetadataReferences(Worksheet worksheet) {
            foreach (Cell cell in worksheet.Descendants<Cell>()) {
                cell.RemoveAttribute("cm", string.Empty);
                cell.RemoveAttribute("vm", string.Empty);
            }
        }

        private static WorksheetStyleCopyMap RemapCopiedWorksheetStyles(WorkbookPart sourceWorkbookPart, WorkbookPart targetWorkbookPart, Worksheet worksheet) {
            Stylesheet? sourceStylesheet = sourceWorkbookPart.WorkbookStylesPart?.Stylesheet;
            if (sourceStylesheet?.CellFormats == null) {
                return WorksheetStyleCopyMap.Empty;
            }

            WorkbookStylesPart targetStylesPart = targetWorkbookPart.WorkbookStylesPart ?? targetWorkbookPart.AddNewPart<WorkbookStylesPart>();
            Stylesheet targetStylesheet = targetStylesPart.Stylesheet ??= CreateDefaultStylesheet();
            EnsureStylesheetPrimitives(targetStylesheet);

            var colorResolver = WorkbookStyleColorResolver.Create(sourceWorkbookPart, sourceStylesheet);
            Dictionary<uint, uint> numberingMap = AppendNumberingFormats(sourceStylesheet, targetStylesheet);
            Dictionary<uint, uint> fontMap = AppendStyleElements<Fonts, Font>(sourceStylesheet.Fonts, targetStylesheet.Fonts!, (container, count) => container.Count = count, colorResolver);
            Dictionary<uint, uint> fillMap = AppendStyleElements<Fills, Fill>(sourceStylesheet.Fills, targetStylesheet.Fills!, (container, count) => container.Count = count, colorResolver);
            Dictionary<uint, uint> borderMap = AppendStyleElements<Borders, Border>(sourceStylesheet.Borders, targetStylesheet.Borders!, (container, count) => container.Count = count, colorResolver);
            Dictionary<uint, uint> styleFormatMap = AppendCellStyleFormats(sourceStylesheet.CellStyleFormats, targetStylesheet.CellStyleFormats!, numberingMap, fontMap, fillMap, borderMap);
            Dictionary<uint, uint> cellFormatMap = AppendCellFormats(sourceStylesheet.CellFormats, targetStylesheet.CellFormats!, numberingMap, fontMap, fillMap, borderMap, styleFormatMap);
            Dictionary<uint, uint> differentialFormatMap = AppendDifferentialFormats(sourceStylesheet.DifferentialFormats, targetStylesheet.DifferentialFormats!, colorResolver);
            var inheritedStyleCells = BuildInheritedStyleCellSet(worksheet);

            foreach (Cell cell in worksheet.Descendants<Cell>()) {
                uint? oldStyleIndex = cell.StyleIndex?.Value;
                if (oldStyleIndex.HasValue && cellFormatMap.TryGetValue(oldStyleIndex.Value, out uint newStyleIndex)) {
                    cell.StyleIndex = newStyleIndex;
                } else if (!oldStyleIndex.HasValue
                    && !CellHasInheritedRowOrColumnStyle(cell, inheritedStyleCells)
                    && cellFormatMap.TryGetValue(0U, out uint newDefaultStyleIndex)) {
                    cell.StyleIndex = newDefaultStyleIndex;
                }
            }

            foreach (Row row in worksheet.Descendants<Row>()) {
                uint? oldStyleIndex = row.StyleIndex?.Value;
                if (oldStyleIndex.HasValue && cellFormatMap.TryGetValue(oldStyleIndex.Value, out uint newStyleIndex)) {
                    row.StyleIndex = newStyleIndex;
                }
            }

            foreach (Column column in worksheet.Descendants<Column>()) {
                uint? oldStyleIndex = column.Style?.Value;
                if (oldStyleIndex.HasValue && cellFormatMap.TryGetValue(oldStyleIndex.Value, out uint newStyleIndex)) {
                    column.Style = newStyleIndex;
                }
            }

            EnsureStylesheetPrimitives(targetStylesheet);
            targetStylesheet.Save();
            return new WorksheetStyleCopyMap(cellFormatMap, differentialFormatMap);
        }

        private static HashSet<string> BuildInheritedStyleCellSet(Worksheet worksheet) {
            string[] cells = worksheet.Descendants<Cell>()
                .Select(cell => cell.CellReference?.Value)
                .Where(reference => !string.IsNullOrEmpty(reference))
                .Select(reference => reference!)
                .ToArray();
            if (cells.Length == 0) {
                return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            }

            var styledRows = new HashSet<int>(worksheet.Descendants<Row>()
                .Where(row => row.StyleIndex != null && row.RowIndex != null)
                .Select(row => (int)row.RowIndex!.Value));
            var styledColumns = worksheet.Elements<Columns>()
                .SelectMany(columns => columns.Elements<Column>())
                .Where(column => column.Style != null)
                .Select(column => (
                    Min: (int)(column.Min?.Value ?? 1U),
                    Max: (int)(column.Max?.Value ?? column.Min?.Value ?? 1U)))
                .ToArray();
            if (styledRows.Count == 0 && styledColumns.Length == 0) {
                return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            }

            var inherited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string reference in cells) {
                (int row, int column) = A1.ParseCellRef(reference);
                if (styledRows.Contains(row) || styledColumns.Any(range => column >= range.Min && column <= range.Max)) {
                    inherited.Add(reference);
                }
            }

            return inherited;
        }

        private static bool CellHasInheritedRowOrColumnStyle(Cell cell, HashSet<string> inheritedStyleCells) {
            string? reference = cell.CellReference?.Value;
            return !string.IsNullOrEmpty(reference) && inheritedStyleCells.Contains(reference!);
        }

        private static void RemapCopiedWorksheetConditionalFormats(Worksheet worksheet, IReadOnlyDictionary<uint, uint> differentialFormatMap) {
            if (differentialFormatMap.Count == 0) {
                return;
            }

            foreach (ConditionalFormattingRule rule in worksheet.Descendants<ConditionalFormattingRule>()) {
                uint? oldDxfId = rule.FormatId?.Value;
                if (oldDxfId.HasValue && differentialFormatMap.TryGetValue(oldDxfId.Value, out uint newDxfId)) {
                    rule.FormatId = newDxfId;
                }
            }

            foreach (OpenXmlElement element in worksheet.Descendants<OpenXmlElement>()) {
                foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                    if (!string.Equals(attribute.LocalName, "dxfId", StringComparison.Ordinal)
                        || !uint.TryParse(attribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out uint oldDxfId)
                        || !differentialFormatMap.TryGetValue(oldDxfId, out uint newDxfId)) {
                        continue;
                    }

                    element.SetAttribute(new OpenXmlAttribute(
                        attribute.Prefix,
                        attribute.LocalName,
                        attribute.NamespaceUri,
                        newDxfId.ToString(CultureInfo.InvariantCulture)));
                }
            }
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

        private static Dictionary<uint, uint> AppendStyleElements<TContainer, TElement>(
            TContainer? source,
            TContainer target,
            Action<TContainer, uint> setCount,
            WorkbookStyleColorResolver colorResolver)
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
                var copied = element.CloneNode(true);
                colorResolver.NormalizeThemeAndIndexedColors(copied);
                target.Append(copied);
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

        private static Dictionary<uint, uint> AppendDifferentialFormats(DifferentialFormats? source, DifferentialFormats target, WorkbookStyleColorResolver colorResolver) {
            var map = new Dictionary<uint, uint>();
            if (source == null) {
                return map;
            }

            uint offset = (uint)target.Elements<DifferentialFormat>().Count();
            uint index = 0;
            foreach (DifferentialFormat sourceFormat in source.Elements<DifferentialFormat>()) {
                uint newIndex = offset + index;
                var copied = sourceFormat.CloneNode(true);
                colorResolver.NormalizeThemeAndIndexedColors(copied);
                target.Append(copied);
                map[index] = newIndex;
                index++;
            }

            target.Count = (uint)target.Elements<DifferentialFormat>().Count();
            return map;
        }

        private static uint MapValue(uint? value, IReadOnlyDictionary<uint, uint> map) {
            if (!value.HasValue) {
                return 0U;
            }

            return map.TryGetValue(value.Value, out uint mapped) ? mapped : value.Value;
        }

        private sealed class WorksheetStyleCopyMap {
            internal static readonly WorksheetStyleCopyMap Empty = new WorksheetStyleCopyMap(
                new Dictionary<uint, uint>(),
                new Dictionary<uint, uint>());

            internal WorksheetStyleCopyMap(
                IReadOnlyDictionary<uint, uint> cellFormats,
                IReadOnlyDictionary<uint, uint> differentialFormats) {
                CellFormats = cellFormats;
                DifferentialFormats = differentialFormats;
            }

            internal IReadOnlyDictionary<uint, uint> CellFormats { get; }
            internal IReadOnlyDictionary<uint, uint> DifferentialFormats { get; }
        }

        private sealed class WorksheetPackageCopyResult {
            internal WorksheetPackageCopyResult(
                ExcelSheet sheet,
                IReadOnlyDictionary<string, string> tableNameMap,
                IReadOnlyDictionary<int, int>? externalReferenceMap = null) {
                Sheet = sheet;
                TableNameMap = tableNameMap;
                ExternalReferenceMap = externalReferenceMap ?? new Dictionary<int, int>();
            }

            internal ExcelSheet Sheet { get; }

            internal IReadOnlyDictionary<string, string> TableNameMap { get; }

            internal IReadOnlyDictionary<int, int> ExternalReferenceMap { get; }
        }

        private sealed class WorkbookStyleColorResolver {
            private readonly WorkbookPart _workbookPart;
            private readonly Dictionary<uint, string> _indexedColors;

            private WorkbookStyleColorResolver(WorkbookPart workbookPart, Dictionary<uint, string> indexedColors) {
                _workbookPart = workbookPart;
                _indexedColors = indexedColors;
            }

            internal static WorkbookStyleColorResolver Create(WorkbookPart workbookPart, Stylesheet sourceStylesheet) {
                return new WorkbookStyleColorResolver(workbookPart, ReadIndexedColors(sourceStylesheet));
            }

            internal void NormalizeThemeAndIndexedColors(OpenXmlElement element) {
                foreach (ColorType color in EnumerateColorTypes(element)) {
                    NormalizeColor(color);
                }
            }

            private void NormalizeColor(ColorType color) {
                string? argb = Utilities.ExcelThemeColorResolver.Resolve(color, _workbookPart, _indexedColors);
                if (argb == null) {
                    return;
                }

                color.Rgb = argb;
                color.Theme = null;
                color.Indexed = null;
                color.Tint = null;
            }

            private static IEnumerable<ColorType> EnumerateColorTypes(OpenXmlElement element) {
                if (element is ColorType color) {
                    yield return color;
                }

                foreach (ColorType child in element.Descendants<ColorType>()) {
                    yield return child;
                }
            }

            private static Dictionary<uint, string> ReadIndexedColors(Stylesheet sourceStylesheet) {
                var colors = new Dictionary<uint, string>();
                IndexedColors? indexedColors = sourceStylesheet.Colors?.IndexedColors;
                if (indexedColors == null) {
                    return colors;
                }

                uint customIndex = 0;
                foreach (RgbColor color in indexedColors.Elements<RgbColor>()) {
                    string? rgb = color.Rgb?.Value;
                    string? argb = Utilities.ExcelThemeColorResolver.NormalizeArgb(rgb);
                    if (argb != null) {
                        colors[customIndex] = argb;
                    }

                    customIndex++;
                }

                return colors;
            }
        }
    }
}
