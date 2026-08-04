using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.PowerPoint {
    internal static class PowerPointChartWorkbookEditor {
        internal static bool IsSafelyEditable(SpreadsheetDocument workbook) {
            WorkbookPart? workbookPart = workbook.WorkbookPart;
            if (workbook.DocumentType != SpreadsheetDocumentType.Workbook
                || workbookPart?.Workbook == null
                || workbookPart.VbaProjectPart != null
                || workbook.Parts.Count(pair =>
                    pair.OpenXmlPart is WorkbookPart) != 1
                || workbook.Parts.Any(pair =>
                    pair.OpenXmlPart is not WorkbookPart
                    && pair.OpenXmlPart is not CoreFilePropertiesPart
                    && pair.OpenXmlPart is not ExtendedFilePropertiesPart
                    && pair.OpenXmlPart is not CustomFilePropertiesPart)
                || workbook.ExternalRelationships.Any()
                || HasUnsupportedRelationships(workbookPart)) {
                return false;
            }

            WorksheetPart[] worksheets = workbookPart
                .GetPartsOfType<WorksheetPart>().ToArray();
            SharedStringTablePart[] sharedStrings = workbookPart
                .GetPartsOfType<SharedStringTablePart>().ToArray();
            WorkbookStylesPart[] styles = workbookPart
                .GetPartsOfType<WorkbookStylesPart>().ToArray();
            ThemePart[] themes = workbookPart.GetPartsOfType<ThemePart>()
                .ToArray();
            if (worksheets.Length != 1 || sharedStrings.Length != 1
                || styles.Length > 1 || themes.Length > 1
                || workbookPart.Parts.Any(pair =>
                    pair.OpenXmlPart is not WorksheetPart
                    && pair.OpenXmlPart is not SharedStringTablePart
                    && pair.OpenXmlPart is not WorkbookStylesPart
                    && pair.OpenXmlPart is not ThemePart)) {
                return false;
            }

            S.Sheet? sheet = workbookPart.Workbook.Sheets?
                .Elements<S.Sheet>().SingleOrDefault();
            WorksheetPart worksheetPart = worksheets[0];
            if (sheet?.Id?.Value == null
                || !string.Equals(sheet.Id.Value,
                    workbookPart.GetIdOfPart(worksheetPart),
                    StringComparison.Ordinal)
                || !string.Equals(sheet.Name?.Value, "Sheet1",
                    StringComparison.Ordinal)
                || workbookPart.Workbook.DefinedNames != null
                || workbookPart.Workbook.ExternalReferences != null) {
                return false;
            }

            S.Worksheet? worksheet = worksheetPart.Worksheet;
            TableDefinitionPart[] tables = worksheetPart
                .GetPartsOfType<TableDefinitionPart>().ToArray();
            if (worksheet == null || tables.Length > 1
                || worksheetPart.Parts.Any(pair =>
                    pair.OpenXmlPart is not TableDefinitionPart)
                || sharedStrings[0].Parts.Any()
                || styles.Any(part => part.Parts.Any())
                || themes.Any(part => part.Parts.Any())
                || tables.Any(part => part.Parts.Any())
                || HasUnsupportedRelationships(worksheetPart)
                || HasUnsupportedRelationships(sharedStrings[0])
                || styles.Any(HasUnsupportedRelationships)
                || themes.Any(HasUnsupportedRelationships)
                || tables.Any(HasUnsupportedRelationships)
                || HasMeaningChangingWorksheetContent(worksheet)) {
                return false;
            }

            return tables.Length == 0 || IsSimpleChartDataTable(tables[0]);
        }

        internal static byte[] Update(byte[] source,
            OfficeChartData data) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (data == null) throw new ArgumentNullException(nameof(data));
            using var stream = new MemoryStream();
            stream.Write(source, 0, source.Length);
            stream.Position = 0;
            using (SpreadsheetDocument workbook = SpreadsheetDocument.Open(
                       stream, true,
                       PowerPointChartWorkbookSecurity.CreateOpenSettings())) {
                if (!IsSafelyEditable(workbook)) {
                    throw new NotSupportedException(
                        "The imported workbook contains sheets, formulas, macros, relationships, or range-dependent content that cannot be updated without changing its meaning.");
                }
                WorkbookPart workbookPart = workbook.WorkbookPart!;
                WorksheetPart worksheetPart = workbookPart
                    .GetPartsOfType<WorksheetPart>().Single();
                SharedStringTablePart sharedStringPart = workbookPart
                    .GetPartsOfType<SharedStringTablePart>().Single();
                S.Worksheet worksheet = worksheetPart.Worksheet!;
                S.SheetData sheetData = worksheet.GetFirstChild<S.SheetData>()
                    ?? worksheet.AppendChild(new S.SheetData());
                TableDefinitionPart? tablePart = worksheetPart
                    .GetPartsOfType<TableDefinitionPart>().SingleOrDefault();
                if (tablePart != null) {
                    int columnCount = tablePart.Table?.TableColumns?
                        .Elements<S.TableColumn>().Count() ?? 0;
                    if (columnCount != data.Series.Count + 1) {
                        throw new NotSupportedException(
                            "Changing the series count of a producer-authored chart table can alter its meaning.");
                    }
                }

                S.SharedStringTable sharedStrings =
                    sharedStringPart.SharedStringTable
                    ?? (sharedStringPart.SharedStringTable =
                        new S.SharedStringTable());
                var indices = new Dictionary<string, int>(
                    StringComparer.Ordinal);
                int sharedStringIndex = 0;
                foreach (S.SharedStringItem item in sharedStrings
                             .Elements<S.SharedStringItem>()) {
                    string value = item.InnerText;
                    if (!indices.ContainsKey(value))
                        indices.Add(value, sharedStringIndex);
                    sharedStringIndex++;
                }

                int GetSharedStringIndex(string? value) {
                    string text = value ?? string.Empty;
                    if (indices.TryGetValue(text, out int existing))
                        return existing;
                    int created = sharedStringIndex++;
                    sharedStrings.AppendChild(new S.SharedStringItem(
                        new S.Text(text) {
                            Space = text.Length != text.Trim().Length
                                ? SpaceProcessingModeValues.Preserve : null
                        }));
                    indices.Add(text, created);
                    return created;
                }

                Dictionary<string, S.Cell> existingCells = sheetData
                    .Elements<S.Row>().SelectMany(row => row.Elements<S.Cell>())
                    .Where(cell => !string.IsNullOrWhiteSpace(
                        cell.CellReference?.Value))
                    .ToDictionary(cell => cell.CellReference!.Value!,
                        cell => cell, StringComparer.OrdinalIgnoreCase);
                S.Row[] existingRows = sheetData.Elements<S.Row>().ToArray();

                S.Row CreateRow(uint rowIndex, bool headerRow) {
                    S.Row? template = existingRows.FirstOrDefault(row =>
                        row.RowIndex?.Value == rowIndex)
                        ?? (headerRow
                            ? existingRows.FirstOrDefault(row =>
                                row.RowIndex?.Value == 1U)
                            : existingRows.FirstOrDefault(row =>
                                row.RowIndex?.Value >= 2U));
                    S.Row row = template == null
                        ? new S.Row()
                        : (S.Row)template.CloneNode(false);
                    row.RowIndex = rowIndex;
                    return row;
                }

                S.Cell CreateCell(string reference, string column,
                    bool headerCell) {
                    S.Cell? template;
                    if (!existingCells.TryGetValue(reference,
                            out template)) {
                        string preferred = column + (headerCell ? "1" : "2");
                        existingCells.TryGetValue(preferred, out template);
                    }
                    S.Cell cell = template == null
                        ? new S.Cell()
                        : (S.Cell)template.CloneNode(false);
                    cell.CellReference = reference;
                    cell.CellValue = null;
                    cell.InlineString = null;
                    return cell;
                }

                sheetData.RemoveAllChildren<S.Row>();
                S.Row header = CreateRow(1U, headerRow: true);
                header.Append(CreateSharedStringCell(
                    CreateCell("A1", "A", headerCell: true),
                    GetSharedStringIndex(" ")));
                for (int seriesIndex = 0;
                     seriesIndex < data.Series.Count; seriesIndex++) {
                    string column = GetExcelColumn(seriesIndex + 2);
                    header.Append(CreateSharedStringCell(
                        CreateCell(column + "1", column,
                            headerCell: true),
                        GetSharedStringIndex(data.Series[seriesIndex].Name)));
                }
                sheetData.Append(header);

                for (int rowIndex = 0;
                     rowIndex < data.Categories.Count; rowIndex++) {
                    uint excelRow = checked((uint)rowIndex + 2U);
                    S.Row row = CreateRow(excelRow, headerRow: false);
                    row.Append(CreateSharedStringCell(
                        CreateCell("A" + excelRow, "A",
                            headerCell: false),
                        GetSharedStringIndex(data.Categories[rowIndex])));
                    for (int seriesIndex = 0;
                         seriesIndex < data.Series.Count; seriesIndex++) {
                        string column = GetExcelColumn(seriesIndex + 2);
                        S.Cell cell = CreateCell(column + excelRow, column,
                            headerCell: false);
                        cell.DataType = null;
                        cell.CellValue = new S.CellValue(
                            data.Series[seriesIndex].Values[rowIndex]);
                        row.Append(cell);
                    }
                    sheetData.Append(row);
                }

                string lastColumn = GetExcelColumn(data.Series.Count + 1);
                uint lastRow = checked((uint)data.Categories.Count + 1U);
                string reference = $"A1:{lastColumn}{lastRow}";
                S.SheetDimension? dimension = worksheet
                    .GetFirstChild<S.SheetDimension>();
                if (dimension == null) {
                    worksheet.InsertAt(new S.SheetDimension {
                        Reference = reference
                    }, 0);
                } else {
                    dimension.Reference = reference;
                }
                if (tablePart?.Table is S.Table table) {
                    table.Reference = reference;
                    if (table.AutoFilter != null)
                        table.AutoFilter.Reference = reference;
                    S.TableColumn[] columns = table.TableColumns!
                        .Elements<S.TableColumn>().ToArray();
                    columns[0].Name = " ";
                    for (int index = 0; index < data.Series.Count; index++)
                        columns[index + 1].Name = data.Series[index].Name;
                    table.Save();
                }
                sharedStrings.Count = checked((uint)(1
                    + data.Series.Count + data.Categories.Count));
                sharedStrings.UniqueCount = checked((uint)
                    sharedStrings.Elements<S.SharedStringItem>().Count());
                sharedStrings.Save();
                worksheet.Save();
                workbookPart.Workbook!.Save();
            }
            return stream.ToArray();
        }

        private static S.Cell CreateSharedStringCell(S.Cell cell,
            int index) {
            cell.DataType = S.CellValues.SharedString;
            cell.CellValue = new S.CellValue(index);
            return cell;
        }

        private static bool HasMeaningChangingWorksheetContent(
            S.Worksheet worksheet) =>
            worksheet.Descendants<S.CellFormula>().Any()
            || worksheet.Descendants<S.Hyperlinks>().Any()
            || worksheet.Descendants<S.OleObjects>().Any()
            || worksheet.Descendants<S.Controls>().Any()
            || worksheet.Descendants<S.MergeCells>().Any()
            || worksheet.Descendants<S.ConditionalFormatting>().Any()
            || worksheet.Descendants<S.DataValidations>().Any()
            || worksheet.Descendants<S.SheetProtection>().Any()
            || worksheet.Descendants<S.ProtectedRanges>().Any()
            || worksheet.Descendants<S.Scenarios>().Any()
            || worksheet.Descendants<S.CustomSheetViews>().Any()
            || worksheet.Descendants<S.PivotTableDefinition>().Any()
            || worksheet.Descendants<S.Drawing>().Any()
            || worksheet.Descendants<S.LegacyDrawing>().Any();

        private static bool IsSimpleChartDataTable(
            TableDefinitionPart part) {
            S.Table? table = part.Table;
            S.TableColumn[] columns = table?.TableColumns?
                .Elements<S.TableColumn>().ToArray()
                ?? Array.Empty<S.TableColumn>();
            return table?.Reference?.Value != null
                && columns.Length > 1
                && columns.All(column =>
                    column.CalculatedColumnFormula == null
                    && column.TotalsRowFormula == null)
                && table.Descendants<S.SortState>().Count() == 0;
        }

        private static bool HasUnsupportedRelationships(OpenXmlPart part) =>
            part.ExternalRelationships.Any()
            || part.HyperlinkRelationships.Any()
            || part.DataPartReferenceRelationships.Any();

        private static string GetExcelColumn(int index) {
            string result = string.Empty;
            while (index > 0) {
                index--;
                result = (char)('A' + index % 26) + result;
                index /= 26;
            }
            return result;
        }
    }
}
