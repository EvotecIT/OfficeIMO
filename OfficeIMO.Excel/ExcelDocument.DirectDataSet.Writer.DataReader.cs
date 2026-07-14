using System.Data;
using System.IO.Compression;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static partial class DirectDataSetWorkbookWriter {
            internal static int WriteDataReader(
                Stream stream,
                DirectDataSetWorkbookModel model,
                IDataReader reader,
                CancellationToken ct) {
                DirectStylePlan stylePlan = DirectStylePlan.Create(model);
                DirectColumnWritePlan[] columnWritePlans = CreateColumnWritePlans(model, stylePlan, ct);
                using var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);
                WriteContentTypes(archive, model.Sheets, includeSharedStrings: false);
                WriteTextEntry(archive, "_rels/.rels",
                    "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                    "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>" +
                    "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>" +
                    "</Relationships>");
                WriteCoreProperties(archive);
                WriteAppProperties(archive);
                WriteWorkbook(archive, model);
                WriteWorkbookRelationships(archive, model.Sheets.Count, includeSharedStrings: false);
                WriteStyles(archive, stylePlan);
                return WriteDataReaderWorksheet(
                    archive,
                    model.Sheets[0],
                    reader,
                    model.DateTimeOffsetWriteStrategy,
                    model.DateSystem,
                    columnWritePlans[0],
                    ct);
            }

            private static int WriteDataReaderWorksheet(
                ZipArchive archive,
                DirectDataSetSheetModel sheet,
                IDataReader reader,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                ExcelDateSystem dateSystem,
                DirectColumnWritePlan columnWritePlan,
                CancellationToken ct) {
                var entry = archive.CreateEntry("xl/worksheets/sheet1.xml", CompressionLevel.Fastest);
                using var stream = entry.Open();
                using var writer = new StreamWriter(stream, Utf8NoBom, XmlWriterBufferSize);

                writer.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                writer.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheetData>");

                int columnCount = sheet.Table.ColumnCount;
                string[] cellReferencePrefixes = CreateCellReferencePrefixes(columnCount);
                if (sheet.IncludeHeaders) {
                    const string headerRowReference = "1";
                    writer.Write("<row r=\"1\">");
                    for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                        WriteCell(
                            writer,
                            headerRowReference,
                            cellReferencePrefixes[columnIndex],
                            sheet.Table.GetColumnName(columnIndex),
                            styleAttribute: null,
                            dateTimeOffsetWriteStrategy,
                            dateSystem,
                            sharedStrings: null);
                    }
                    writer.Write("</row>");
                }

                int rowCount = 0;
                int maximumDataRows = A1.MaxRows - (sheet.IncludeHeaders ? 1 : 0);
                bool canCancel = ct.CanBeCanceled;
                bool useBulkRead = ExcelSheet.CanUseBulkDataReaderValues(reader);
                object?[] values = new object?[columnCount];
                while (reader.Read()) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }
                    if (rowCount >= maximumDataRows) {
                        throw new InvalidOperationException("Data reader export exceeds the maximum worksheet row count.");
                    }

                    ExcelSheet.FillDataReaderValues(reader, values, columnCount, ref useBulkRead);
                    if (sheet.IncludeCellReferences) {
                        WriteReferencedDataReaderRow(
                            writer,
                            values,
                            rowCount + (sheet.IncludeHeaders ? 2 : 1),
                            cellReferencePrefixes,
                            columnWritePlan,
                            sheet.UseCellValueNumberFormats,
                            dateTimeOffsetWriteStrategy,
                            dateSystem);
                    } else {
                        WriteCompactDataReaderRow(
                            writer,
                            values,
                            columnWritePlan,
                            sheet.UseCellValueNumberFormats,
                            dateTimeOffsetWriteStrategy,
                            dateSystem);
                    }
                    rowCount++;
                }

                writer.Write("</sheetData></worksheet>");
                return rowCount;
            }

            private static void WriteCompactDataReaderRow(
                TextWriter writer,
                object?[] values,
                DirectColumnWritePlan columnWritePlan,
                bool useCellValueNumberFormats,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                ExcelDateSystem dateSystem) {
                string?[]? styleAttributes = columnWritePlan.StyleAttributes;
                DirectCellValueKind[] cellValueKinds = columnWritePlan.CellValueKinds;
                bool[]? valueStyleColumns = columnWritePlan.ValueStyleColumns;
                if (valueStyleColumns == null &&
                    IsIntStringStringDateTimeDoubleIntBooleanStringPlan(cellValueKinds)) {
                    WriteCompactIntStringStringDateTimeDoubleIntBooleanStringRow(
                        writer,
                        values,
                        offset: 0,
                        styleAttributes,
                        dateTimeOffsetWriteStrategy,
                        dateSystem,
                        sharedStrings: null);
                    return;
                }

                writer.Write("<row>");
                for (int columnIndex = 0; columnIndex < cellValueKinds.Length; columnIndex++) {
                    object? value = values[columnIndex];
                    writer.Write("<c");
                    bool useValueStyle = valueStyleColumns?[columnIndex] ?? false;
                    string? styleAttribute = styleAttributes?[columnIndex]
                        ?? (useValueStyle ? CreateStyleAttributeForValue(value, useCellValueNumberFormats) : null);
                    if (styleAttribute != null) {
                        writer.Write(styleAttribute);
                    }

                    if (useValueStyle) {
                        WriteCellValue(writer, value, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings: null);
                    } else {
                        WriteCellValue(writer, value, cellValueKinds[columnIndex], dateTimeOffsetWriteStrategy, dateSystem, sharedStrings: null);
                    }
                }
                writer.Write("</row>");
            }

            private static void WriteReferencedDataReaderRow(
                TextWriter writer,
                object?[] values,
                int rowIndex,
                string[] cellReferencePrefixes,
                DirectColumnWritePlan columnWritePlan,
                bool useCellValueNumberFormats,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                ExcelDateSystem dateSystem) {
                string rowReference = InvariantNumberText.Get(rowIndex);
                writer.Write("<row r=\"");
                writer.Write(rowReference);
                writer.Write("\">");

                string?[]? styleAttributes = columnWritePlan.StyleAttributes;
                DirectCellValueKind[] cellValueKinds = columnWritePlan.CellValueKinds;
                bool[]? valueStyleColumns = columnWritePlan.ValueStyleColumns;
                for (int columnIndex = 0; columnIndex < cellValueKinds.Length; columnIndex++) {
                    WriteCell(
                        writer,
                        rowReference,
                        cellReferencePrefixes[columnIndex],
                        values[columnIndex],
                        styleAttributes?[columnIndex],
                        valueStyleColumns?[columnIndex] ?? false,
                        cellValueKinds[columnIndex],
                        useCellValueNumberFormats,
                        dateTimeOffsetWriteStrategy,
                        dateSystem,
                        sharedStrings: null);
                }
                writer.Write("</row>");
            }
        }
    }
}
