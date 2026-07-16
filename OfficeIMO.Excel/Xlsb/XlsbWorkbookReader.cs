using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb.Package;
using OfficeIMO.Excel.Xlsb.Styles;
using System.IO.Compression;

namespace OfficeIMO.Excel.Xlsb {
    /// <summary>Reads workbook metadata and worksheet values from BIFF12 package parts.</summary>
    internal static class XlsbWorkbookReader {
        private const int BrtRowHdr = 0;
        private const int BrtCellBlank = 1;
        private const int BrtCellRk = 2;
        private const int BrtCellError = 3;
        private const int BrtCellBool = 4;
        private const int BrtCellReal = 5;
        private const int BrtCellSt = 6;
        private const int BrtCellIsst = 7;
        private const int BrtFmlaString = 8;
        private const int BrtFmlaNum = 9;
        private const int BrtFmlaBool = 10;
        private const int BrtFmlaError = 11;
        private const int BrtCellRString = 62;
        private const int BrtBeginSheet = 129;
        private const int BrtEndSheet = 130;
        private const int BrtBeginBook = 131;
        private const int BrtEndBook = 132;
        private const int BrtBeginSheetData = 145;
        private const int BrtEndSheetData = 146;
        private const int BrtWbProp = 153;
        private const int BrtBeginBundleShs = 143;
        private const int BrtEndBundleShs = 144;
        private const int BrtBundleSh = 156;
        private const int BrtBeginSst = 159;
        private const int BrtEndSst = 160;
        private const int BrtSstItem = 19;

        private const string WorksheetRelationshipSuffix = "/worksheet";
        private const string SharedStringsRelationshipSuffix = "/sharedStrings";
        private const string StylesRelationshipSuffix = "/styles";

        internal static XlsbWorkbook Load(byte[] packageBytes, XlsbImportOptions? options = null) {
            if (packageBytes == null) throw new ArgumentNullException(nameof(packageBytes));
            XlsbImportOptions resolved = options ?? new XlsbImportOptions();
            resolved.Validate();

            if (!XlsbPackageDetector.TryFindWorkbookPart(packageBytes, out string? workbookPartName)
                || string.IsNullOrWhiteSpace(workbookPartName)) {
                throw new InvalidDataException("The package does not contain a canonical XLSB workbook part.");
            }

            var workbook = new XlsbWorkbook((byte[])packageBytes.Clone());
            using var packageStream = new MemoryStream(packageBytes, writable: false);
            using var archive = new ZipArchive(packageStream, ZipArchiveMode.Read, leaveOpen: false);
            var parts = new XlsbPackagePartReader(archive, resolved);
            IReadOnlyDictionary<string, XlsbPackageRelationship> relationships = parts.ReadRelationships(workbookPartName!);
            IReadOnlyList<string> sharedStrings = ReadSharedStrings(parts, workbookPartName!, relationships, resolved, workbook);
            workbook.Stylesheet = ReadStyles(parts, workbookPartName!, relationships, resolved, workbook);
            ParseWorkbookPart(parts.ReadPart(workbookPartName!), workbookPartName!, resolved, workbook);

            if (workbook.Worksheets.Count == 0) {
                throw new InvalidDataException("The XLSB workbook contains no worksheet bundle records.");
            }

            if (workbook.Worksheets.Count > resolved.MaxWorksheets) {
                throw new InvalidDataException($"The XLSB workbook contains {workbook.Worksheets.Count} worksheets, exceeding the configured limit of {resolved.MaxWorksheets}.");
            }

            int totalCells = 0;
            foreach (XlsbWorksheet worksheet in workbook.Worksheets) {
                if (!relationships.TryGetValue(worksheet.RelationshipId, out XlsbPackageRelationship? relationship)
                    || relationship.IsExternal
                    || !relationship.Type.EndsWith(WorksheetRelationshipSuffix, StringComparison.Ordinal)) {
                    throw new InvalidDataException($"The XLSB worksheet '{worksheet.Name}' refers to missing or invalid relationship '{worksheet.RelationshipId}'.");
                }

                string sheetPartName = XlsbPackagePartReader.ResolveTarget(workbookPartName!, relationship.Target);
                worksheet.PartName = sheetPartName;
                ParseWorksheetPart(
                    parts.ReadPart(sheetPartName),
                    sheetPartName,
                    worksheet,
                    sharedStrings,
                    resolved,
                    workbook,
                    ref totalCells);
            }

            workbook.SharedStringCount = sharedStrings.Count;
            if (workbook.PreservedRecords.Count > 0) {
                workbook.AddDiagnostic(new XlsbImportDiagnostic(
                    "XLSB-RECORDS-PRESERVED",
                    XlsbImportDiagnosticSeverity.Information,
                    $"Preserved {workbook.PreservedRecords.Count} BIFF12 records that are not yet projected into the normal workbook model."));
            }

            return workbook;
        }

        private static void ParseWorkbookPart(
            byte[] bytes,
            string partName,
            XlsbImportOptions options,
            XlsbWorkbook workbook) {
            IReadOnlyList<XlsbRecord> records = ReadRecords(bytes, options);
            if (records.Count < 2 || records[0].Type != BrtBeginBook || records[records.Count - 1].Type != BrtEndBook) {
                throw new InvalidDataException($"The XLSB workbook part '{partName}' is missing its BrtBeginBook/BrtEndBook boundaries.");
            }

            foreach (XlsbRecord record in records) {
                switch (record.Type) {
                    case BrtBeginBook:
                    case BrtEndBook:
                    case BrtBeginBundleShs:
                    case BrtEndBundleShs:
                        break;
                    case BrtBundleSh:
                        var cursor = new XlsbBinaryCursor(record.Data);
                        uint state = cursor.ReadUInt32();
                        uint tabId = cursor.ReadUInt32();
                        string relationshipId = cursor.ReadWideString(options.MaxStringCharacters);
                        string name = cursor.ReadWideString(options.MaxStringCharacters);
                        if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(relationshipId)) {
                            throw new InvalidDataException($"The BrtBundleSh record at offset {record.Offset} does not contain a worksheet name and relationship id.");
                        }

                        workbook.AddWorksheet(new XlsbWorksheet(name, relationshipId, tabId, state));
                        break;
                    case BrtWbProp:
                        if (record.Data.Length < 4) {
                            throw new InvalidDataException($"The BrtWbProp record at offset {record.Offset} is truncated.");
                        }

                        var propertiesCursor = new XlsbBinaryCursor(record.Data);
                        workbook.Uses1904DateSystem = (propertiesCursor.ReadUInt32() & 0x01U) != 0;
                        PreserveRecord(options, workbook, partName, record);
                        break;
                    default:
                        PreserveRecord(options, workbook, partName, record);
                        break;
                }
            }
        }

        private static XlsbStylesheet? ReadStyles(
            XlsbPackagePartReader parts,
            string workbookPartName,
            IReadOnlyDictionary<string, XlsbPackageRelationship> relationships,
            XlsbImportOptions options,
            XlsbWorkbook workbook) {
            XlsbPackageRelationship? relationship = relationships.Values.FirstOrDefault(item =>
                !item.IsExternal && item.Type.EndsWith(StylesRelationshipSuffix, StringComparison.Ordinal));
            if (relationship == null) return null;

            string partName = XlsbPackagePartReader.ResolveTarget(workbookPartName, relationship.Target);
            return XlsbStylesheetReader.Read(parts.ReadPart(partName), partName, options, workbook);
        }

        private static IReadOnlyList<string> ReadSharedStrings(
            XlsbPackagePartReader parts,
            string workbookPartName,
            IReadOnlyDictionary<string, XlsbPackageRelationship> relationships,
            XlsbImportOptions options,
            XlsbWorkbook workbook) {
            XlsbPackageRelationship? relationship = relationships.Values.FirstOrDefault(item =>
                !item.IsExternal && item.Type.EndsWith(SharedStringsRelationshipSuffix, StringComparison.Ordinal));
            if (relationship == null) return Array.Empty<string>();

            string partName = XlsbPackagePartReader.ResolveTarget(workbookPartName, relationship.Target);
            IReadOnlyList<XlsbRecord> records = ReadRecords(parts.ReadPart(partName), options);
            var values = new List<string>();
            bool hasBegin = false;
            bool hasEnd = false;
            foreach (XlsbRecord record in records) {
                switch (record.Type) {
                    case BrtBeginSst:
                        hasBegin = true;
                        if (record.Size < 8) {
                            throw new InvalidDataException($"The BrtBeginSst record in '{partName}' is truncated.");
                        }
                        break;
                    case BrtSstItem:
                        if (values.Count >= options.MaxSharedStrings) {
                            throw new InvalidDataException($"The XLSB shared-string table exceeds the configured limit of {options.MaxSharedStrings} items.");
                        }

                        var cursor = new XlsbBinaryCursor(record.Data);
                        byte flags = cursor.ReadByte();
                        values.Add(cursor.ReadWideString(options.MaxStringCharacters));
                        if ((flags & 0x03) != 0 || cursor.Remaining > 0) {
                            PreserveRecord(options, workbook, partName, record);
                        }
                        break;
                    case BrtEndSst:
                        hasEnd = true;
                        break;
                    default:
                        PreserveRecord(options, workbook, partName, record);
                        break;
                }
            }

            if (!hasBegin || !hasEnd) {
                throw new InvalidDataException($"The XLSB shared-string part '{partName}' is missing its boundary records.");
            }

            return values.AsReadOnly();
        }

        private static void ParseWorksheetPart(
            byte[] bytes,
            string partName,
            XlsbWorksheet worksheet,
            IReadOnlyList<string> sharedStrings,
            XlsbImportOptions options,
            XlsbWorkbook workbook,
            ref int totalCells) {
            IReadOnlyList<XlsbRecord> records = ReadRecords(bytes, options);
            if (records.Count < 2 || records[0].Type != BrtBeginSheet || records[records.Count - 1].Type != BrtEndSheet) {
                throw new InvalidDataException($"The XLSB worksheet part '{partName}' is missing its BrtBeginSheet/BrtEndSheet boundaries.");
            }

            int currentRow = -1;
            foreach (XlsbRecord record in records) {
                switch (record.Type) {
                    case BrtBeginSheet:
                    case BrtEndSheet:
                    case BrtBeginSheetData:
                    case BrtEndSheetData:
                        break;
                    case BrtRowHdr:
                        var rowCursor = new XlsbBinaryCursor(record.Data);
                        currentRow = rowCursor.ReadInt32();
                        if (currentRow < 0 || currentRow >= 1_048_576) {
                            throw new InvalidDataException($"The XLSB row record at offset {record.Offset} contains invalid row index {currentRow}.");
                        }
                        break;
                    case BrtCellBlank:
                    case BrtCellRk:
                    case BrtCellError:
                    case BrtCellBool:
                    case BrtCellReal:
                    case BrtCellSt:
                    case BrtCellIsst:
                    case BrtFmlaString:
                    case BrtFmlaNum:
                    case BrtFmlaBool:
                    case BrtFmlaError:
                    case BrtCellRString:
                        if (currentRow < 0) {
                            throw new InvalidDataException($"The XLSB cell record at offset {record.Offset} appears before a row header.");
                        }

                        totalCells = checked(totalCells + 1);
                        if (totalCells > options.MaxCells) {
                            throw new InvalidDataException($"The XLSB workbook exceeds the configured limit of {options.MaxCells} populated cells.");
                        }

                        worksheet.AddCell(ParseCell(record, currentRow, sharedStrings, options, workbook, partName));
                        break;
                    default:
                        PreserveRecord(options, workbook, partName, record);
                        break;
                }
            }
        }

        private static XlsbCell ParseCell(
            XlsbRecord record,
            int zeroBasedRow,
            IReadOnlyList<string> sharedStrings,
            XlsbImportOptions options,
            XlsbWorkbook workbook,
            string partName) {
            var cursor = new XlsbBinaryCursor(record.Data);
            int zeroBasedColumn = cursor.ReadInt32();
            uint styleIndex = cursor.ReadUInt32() & 0x00FFFFFFU;
            if (zeroBasedColumn < 0 || zeroBasedColumn >= 16_384) {
                throw new InvalidDataException($"The XLSB cell record at offset {record.Offset} contains invalid column index {zeroBasedColumn}.");
            }

            int row = zeroBasedRow + 1;
            int column = zeroBasedColumn + 1;
            XlsbCell cell;
            switch (record.Type) {
                case BrtCellBlank:
                    cell = new XlsbCell(row, column, XlsbCellValueKind.Blank, null, styleIndex);
                    break;
                case BrtCellRk:
                    cell = new XlsbCell(row, column, XlsbCellValueKind.Number, BiffRkNumberReader.ReadRkNumber(cursor.ReadUInt32()), styleIndex);
                    break;
                case BrtCellError:
                    cell = new XlsbCell(row, column, XlsbCellValueKind.Error, BiffErrorValue.ToText(cursor.ReadByte()), styleIndex);
                    break;
                case BrtCellBool:
                    cell = new XlsbCell(row, column, XlsbCellValueKind.Boolean, cursor.ReadByte() != 0, styleIndex);
                    break;
                case BrtCellReal:
                    cell = new XlsbCell(row, column, XlsbCellValueKind.Number, cursor.ReadDouble(), styleIndex);
                    break;
                case BrtCellSt:
                    cell = new XlsbCell(row, column, XlsbCellValueKind.Text, cursor.ReadWideString(options.MaxStringCharacters), styleIndex);
                    break;
                case BrtCellIsst:
                    uint sharedStringIndex = cursor.ReadUInt32();
                    if (sharedStringIndex >= sharedStrings.Count) {
                        throw new InvalidDataException($"The XLSB cell at row {row}, column {column} refers to missing shared string {sharedStringIndex}.");
                    }

                    cell = new XlsbCell(row, column, XlsbCellValueKind.Text, sharedStrings[checked((int)sharedStringIndex)], styleIndex);
                    break;
                case BrtCellRString:
                    byte flags = cursor.ReadByte();
                    string richText = cursor.ReadWideString(options.MaxStringCharacters);
                    if ((flags & 0x03) != 0 || cursor.Remaining > 0) {
                        PreserveRecord(options, workbook, partName, record);
                    }
                    cell = new XlsbCell(row, column, XlsbCellValueKind.Text, richText, styleIndex);
                    break;
                case BrtFmlaNum:
                    cell = ParseFormulaCell(record, cursor, row, column, styleIndex, XlsbCellValueKind.Number, cursor.ReadDouble(), options, workbook, partName);
                    break;
                case BrtFmlaBool:
                    cell = ParseFormulaCell(record, cursor, row, column, styleIndex, XlsbCellValueKind.Boolean, cursor.ReadByte() != 0, options, workbook, partName);
                    break;
                case BrtFmlaError:
                    cell = ParseFormulaCell(record, cursor, row, column, styleIndex, XlsbCellValueKind.Error, BiffErrorValue.ToText(cursor.ReadByte()), options, workbook, partName);
                    break;
                case BrtFmlaString:
                    cell = ParseFormulaCell(record, cursor, row, column, styleIndex, XlsbCellValueKind.Text, cursor.ReadWideString(options.MaxStringCharacters), options, workbook, partName);
                    break;
                default:
                    throw new InvalidOperationException($"Unsupported XLSB cell record type {record.Type}.");
            }

            cell.SourceRecordType = record.Type;
            cell.SourceRecordData = (byte[])record.Data.Clone();
            int availableFormats = workbook.Stylesheet?.CellFormats.Count ?? 1;
            if (styleIndex >= availableFormats) {
                throw new InvalidDataException(
                    $"The XLSB cell at row {row}, column {column} refers to missing cell format {styleIndex}; the styles part exposes {availableFormats} format(s).");
            }
            return cell;
        }

        private static XlsbCell ParseFormulaCell(
            XlsbRecord record,
            XlsbBinaryCursor cursor,
            int row,
            int column,
            uint styleIndex,
            XlsbCellValueKind valueKind,
            object? cachedValue,
            XlsbImportOptions options,
            XlsbWorkbook workbook,
            string partName) {
            int formulaPayloadOffset = cursor.Position;
            cursor.ReadUInt16(); // grbit flags
            uint tokenCount = cursor.ReadUInt32();
            if (tokenCount > cursor.Remaining) {
                throw new InvalidDataException($"The XLSB formula record at offset {record.Offset} declares {tokenCount} token bytes but only {cursor.Remaining} remain.");
            }

            byte[] tokens = cursor.ReadBytes(checked((int)tokenCount));
            var cell = new XlsbCell(row, column, valueKind, cachedValue, styleIndex) {
                FormulaBytes = tokens,
                FormulaPayloadBytes = CopyTail(record.Data, formulaPayloadOffset)
            };
            if (XlsbFormulaTextReader.TryRead(tokens, out string? formulaText)) {
                cell.FormulaText = formulaText;
            } else {
                workbook.AddDiagnostic(new XlsbImportDiagnostic(
                    "XLSB-FORMULA-PRESERVED",
                    XlsbImportDiagnosticSeverity.Warning,
                    $"Preserved an unsupported BIFF12 formula token stream at row {row}, column {column}; its cached value was projected.",
                    partName,
                    record.Type,
                    record.Offset));
                PreserveRecord(options, workbook, partName, record);
            }

            return cell;
        }

        private static byte[] CopyTail(byte[] data, int offset) {
            if (offset < 0 || offset > data.Length) throw new ArgumentOutOfRangeException(nameof(offset));
            byte[] result = new byte[data.Length - offset];
            Buffer.BlockCopy(data, offset, result, 0, result.Length);
            return result;
        }

        private static IReadOnlyList<XlsbRecord> ReadRecords(byte[] bytes, XlsbImportOptions options) {
            using var stream = new MemoryStream(bytes, writable: false);
            return XlsbRecordReader.ReadAll(stream, options.MaxRecordBytes);
        }

        private static void PreserveRecord(
            XlsbImportOptions options,
            XlsbWorkbook workbook,
            string partName,
            XlsbRecord record) {
            if (!options.ReportPreservedRecords) return;
            workbook.AddPreservedRecord(new XlsbPreservedRecordInfo(partName, record.Type, record.Offset, record.Size));
        }
    }
}
