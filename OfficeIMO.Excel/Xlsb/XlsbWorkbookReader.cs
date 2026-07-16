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
        private const int BrtColInfo = 60;
        private const int BrtBeginSheet = 129;
        private const int BrtEndSheet = 130;
        private const int BrtBeginBook = 131;
        private const int BrtEndBook = 132;
        private const int BrtBeginSheetData = 145;
        private const int BrtEndSheetData = 146;
        private const int BrtWsDim = 148;
        private const int BrtPane = 151;
        private const int BrtWbProp = 153;
        private const int BrtBeginBundleShs = 143;
        private const int BrtEndBundleShs = 144;
        private const int BrtBundleSh = 156;
        private const int BrtBeginSst = 159;
        private const int BrtEndSst = 160;
        private const int BrtSstItem = 19;
        private const int BrtMergeCell = 176;
        private const int BrtBeginMergeCells = 177;
        private const int BrtEndMergeCells = 178;
        private const int BrtBeginColInfos = 390;
        private const int BrtEndColInfos = 391;
        private const int BrtWsFmtInfo = 485;
        private const int BrtHLink = 494;

        private const string WorksheetRelationshipSuffix = "/worksheet";
        private const string SharedStringsRelationshipSuffix = "/sharedStrings";
        private const string StylesRelationshipSuffix = "/styles";
        private const string HyperlinkRelationshipSuffix = "/hyperlink";

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
            int totalMergedRanges = 0;
            int totalHyperlinks = 0;
            foreach (XlsbWorksheet worksheet in workbook.Worksheets) {
                if (!relationships.TryGetValue(worksheet.RelationshipId, out XlsbPackageRelationship? relationship)
                    || relationship.IsExternal
                    || !relationship.Type.EndsWith(WorksheetRelationshipSuffix, StringComparison.Ordinal)) {
                    throw new InvalidDataException($"The XLSB worksheet '{worksheet.Name}' refers to missing or invalid relationship '{worksheet.RelationshipId}'.");
                }

                string sheetPartName = XlsbPackagePartReader.ResolveTarget(workbookPartName!, relationship.Target);
                worksheet.PartName = sheetPartName;
                IReadOnlyDictionary<string, XlsbPackageRelationship> worksheetRelationships = parts.ReadRelationships(sheetPartName);
                ParseWorksheetPart(
                    parts.ReadPart(sheetPartName),
                    sheetPartName,
                    worksheet,
                    worksheetRelationships,
                    sharedStrings,
                    resolved,
                    workbook,
                    ref totalCells,
                    ref totalMergedRanges,
                    ref totalHyperlinks);
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
            IReadOnlyDictionary<string, XlsbPackageRelationship> worksheetRelationships,
            IReadOnlyList<string> sharedStrings,
            XlsbImportOptions options,
            XlsbWorkbook workbook,
            ref int totalCells,
            ref int totalMergedRanges,
            ref int totalHyperlinks) {
            IReadOnlyList<XlsbRecord> records = ReadRecords(bytes, options);
            if (records.Count < 2 || records[0].Type != BrtBeginSheet || records[records.Count - 1].Type != BrtEndSheet) {
                throw new InvalidDataException($"The XLSB worksheet part '{partName}' is missing its BrtBeginSheet/BrtEndSheet boundaries.");
            }

            XlsbRowInfo? currentRow = null;
            int previousRow = -1;
            int previousColumnEnd = -1;
            bool inSheetData = false;
            bool inColumnInfos = false;
            bool inMergeCells = false;
            bool sawSheetData = false;
            bool sawColumnInfos = false;
            bool sawMergeCells = false;
            uint declaredMergeCount = 0;
            int actualMergeCount = 0;
            foreach (XlsbRecord record in records) {
                switch (record.Type) {
                    case BrtBeginSheet:
                    case BrtEndSheet:
                        break;
                    case BrtBeginSheetData:
                        if (inSheetData || sawSheetData) {
                            throw new InvalidDataException($"The XLSB worksheet part '{partName}' contains duplicate or nested BrtBeginSheetData records.");
                        }
                        inSheetData = true;
                        sawSheetData = true;
                        currentRow = null;
                        break;
                    case BrtEndSheetData:
                        if (!inSheetData) {
                            throw new InvalidDataException($"The XLSB worksheet part '{partName}' contains BrtEndSheetData without a matching begin record.");
                        }
                        inSheetData = false;
                        currentRow = null;
                        break;
                    case BrtWsDim:
                        if (worksheet.UsedRange != null) {
                            throw new InvalidDataException($"The XLSB worksheet part '{partName}' contains more than one BrtWsDim record.");
                        }
                        worksheet.UsedRange = ParseCellRange(record, "BrtWsDim");
                        break;
                    case BrtWsFmtInfo:
                        if (worksheet.FormatInfo != null) {
                            throw new InvalidDataException($"The XLSB worksheet part '{partName}' contains more than one BrtWsFmtInfo record.");
                        }
                        worksheet.FormatInfo = ParseWorksheetFormatInfo(record);
                        break;
                    case BrtPane:
                        if (worksheet.Pane != null) {
                            throw new InvalidDataException($"The XLSB worksheet part '{partName}' contains more than one BrtPane record.");
                        }
                        worksheet.Pane = ParsePane(record);
                        break;
                    case BrtBeginColInfos:
                        if (inColumnInfos || sawColumnInfos) {
                            throw new InvalidDataException($"The XLSB worksheet part '{partName}' contains duplicate or nested BrtBeginColInfos records.");
                        }
                        inColumnInfos = true;
                        sawColumnInfos = true;
                        previousColumnEnd = -1;
                        break;
                    case BrtColInfo:
                        if (!inColumnInfos) {
                            throw new InvalidDataException($"The BrtColInfo record at offset {record.Offset} appears outside its collection.");
                        }
                        XlsbColumnInfo column = ParseColumnInfo(record, workbook);
                        if (column.FirstColumn - 1 <= previousColumnEnd) {
                            throw new InvalidDataException($"The BrtColInfo record at offset {record.Offset} overlaps or precedes another column definition.");
                        }
                        previousColumnEnd = column.LastColumn - 1;
                        worksheet.AddColumn(column);
                        break;
                    case BrtEndColInfos:
                        if (!inColumnInfos) {
                            throw new InvalidDataException($"The XLSB worksheet part '{partName}' contains BrtEndColInfos without a matching begin record.");
                        }
                        inColumnInfos = false;
                        break;
                    case BrtRowHdr:
                        if (!inSheetData) {
                            throw new InvalidDataException($"The XLSB row record at offset {record.Offset} appears outside BrtBeginSheetData/BrtEndSheetData.");
                        }
                        currentRow = ParseRowInfo(record, workbook);
                        if (currentRow.Row - 1 <= previousRow) {
                            throw new InvalidDataException($"The XLSB row record at offset {record.Offset} is duplicated or out of order.");
                        }
                        previousRow = currentRow.Row - 1;
                        worksheet.AddRow(currentRow);
                        break;
                    case BrtBeginMergeCells:
                        if (inMergeCells || sawMergeCells) {
                            throw new InvalidDataException($"The XLSB worksheet part '{partName}' contains duplicate or nested BrtBeginMergeCells records.");
                        }
                        if (record.Data.Length != 4) {
                            throw new InvalidDataException($"The BrtBeginMergeCells record at offset {record.Offset} has invalid payload length {record.Data.Length}.");
                        }
                        var mergeCountCursor = new XlsbBinaryCursor(record.Data);
                        declaredMergeCount = mergeCountCursor.ReadUInt32();
                        if (declaredMergeCount > options.MaxMergedRanges - totalMergedRanges) {
                            throw new InvalidDataException($"The XLSB worksheet '{worksheet.Name}' declares {declaredMergeCount} merged ranges, exceeding the configured limit of {options.MaxMergedRanges}.");
                        }
                        inMergeCells = true;
                        sawMergeCells = true;
                        actualMergeCount = 0;
                        break;
                    case BrtMergeCell:
                        if (!inMergeCells) {
                            throw new InvalidDataException($"The BrtMergeCell record at offset {record.Offset} appears outside its collection.");
                        }
                        actualMergeCount = checked(actualMergeCount + 1);
                        if (actualMergeCount > options.MaxMergedRanges) {
                            throw new InvalidDataException($"The XLSB workbook exceeds the configured limit of {options.MaxMergedRanges} merged ranges.");
                        }
                        worksheet.AddMergedRange(ParseCellRange(record, "BrtMergeCell"));
                        break;
                    case BrtEndMergeCells:
                        if (!inMergeCells || record.Data.Length != 0) {
                            throw new InvalidDataException($"The BrtEndMergeCells record at offset {record.Offset} is invalid or has no matching begin record.");
                        }
                        if (actualMergeCount != declaredMergeCount) {
                            throw new InvalidDataException($"The XLSB worksheet '{worksheet.Name}' declares {declaredMergeCount} merged ranges but contains {actualMergeCount} records.");
                        }
                        totalMergedRanges = checked(totalMergedRanges + actualMergeCount);
                        inMergeCells = false;
                        break;
                    case BrtHLink:
                        totalHyperlinks = checked(totalHyperlinks + 1);
                        if (totalHyperlinks > options.MaxHyperlinks) {
                            throw new InvalidDataException($"The XLSB workbook exceeds the configured limit of {options.MaxHyperlinks} worksheet hyperlinks.");
                        }
                        worksheet.AddHyperlink(ParseHyperlink(record, worksheetRelationships, options));
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
                        if (!inSheetData || currentRow == null) {
                            throw new InvalidDataException($"The XLSB cell record at offset {record.Offset} appears before a row header.");
                        }

                        totalCells = checked(totalCells + 1);
                        if (totalCells > options.MaxCells) {
                            throw new InvalidDataException($"The XLSB workbook exceeds the configured limit of {options.MaxCells} populated cells.");
                        }

                        XlsbCell cell = ParseCell(record, currentRow.Row - 1, sharedStrings, options, workbook, partName);
                        if (!currentRow.ContainsZeroBasedColumn(cell.Column - 1)) {
                            throw new InvalidDataException($"The XLSB cell at row {cell.Row}, column {cell.Column} is not covered by its BrtRowHdr column spans.");
                        }
                        worksheet.AddCell(cell);
                        break;
                    default:
                        PreserveRecord(options, workbook, partName, record);
                        break;
                }
            }

            if (inSheetData || inColumnInfos || inMergeCells) {
                throw new InvalidDataException($"The XLSB worksheet part '{partName}' contains an unterminated record collection.");
            }

            if (worksheet.UsedRange == null) {
                throw new InvalidDataException($"The XLSB worksheet part '{partName}' does not contain BrtWsDim.");
            }
            if (!sawSheetData) {
                throw new InvalidDataException($"The XLSB worksheet part '{partName}' does not contain BrtBeginSheetData/BrtEndSheetData.");
            }
        }

        private static XlsbCellRange ParseCellRange(XlsbRecord record, string recordName) {
            if (record.Data.Length != 16) {
                throw new InvalidDataException($"The {recordName} record at offset {record.Offset} has invalid payload length {record.Data.Length}.");
            }

            var cursor = new XlsbBinaryCursor(record.Data);
            uint firstRow = cursor.ReadUInt32();
            uint lastRow = cursor.ReadUInt32();
            uint firstColumn = cursor.ReadUInt32();
            uint lastColumn = cursor.ReadUInt32();
            if (firstRow > lastRow || lastRow >= A1.MaxRows || firstColumn > lastColumn || lastColumn >= A1.MaxColumns) {
                throw new InvalidDataException($"The {recordName} record at offset {record.Offset} contains an invalid worksheet range.");
            }

            return new XlsbCellRange(
                checked((int)firstRow + 1),
                checked((int)lastRow + 1),
                checked((int)firstColumn + 1),
                checked((int)lastColumn + 1));
        }

        private static XlsbHyperlink ParseHyperlink(
            XlsbRecord record,
            IReadOnlyDictionary<string, XlsbPackageRelationship> relationships,
            XlsbImportOptions options) {
            if (record.Data.Length < 32) {
                throw new InvalidDataException($"The BrtHLink record at offset {record.Offset} is truncated.");
            }

            var cursor = new XlsbBinaryCursor(record.Data);
            byte[] rangePayload = cursor.ReadBytes(16);
            var rangeRecord = new XlsbRecord(record.Offset, record.HeaderSize, BrtHLink, rangePayload);
            XlsbCellRange range = ParseCellRange(rangeRecord, "BrtHLink");
            string relationshipId = cursor.ReadWideString(options.MaxStringCharacters);
            string location = cursor.ReadWideString(options.MaxStringCharacters);
            string tooltip = cursor.ReadWideString(options.MaxStringCharacters);
            string display = cursor.ReadWideString(options.MaxStringCharacters);
            if (cursor.Remaining != 0) {
                throw new InvalidDataException($"The BrtHLink record at offset {record.Offset} contains trailing payload bytes.");
            }

            string? externalTarget = null;
            if (!string.IsNullOrEmpty(relationshipId)) {
                if (!relationships.TryGetValue(relationshipId, out XlsbPackageRelationship? relationship)
                    || !relationship.IsExternal
                    || !relationship.Type.EndsWith(HyperlinkRelationshipSuffix, StringComparison.Ordinal)) {
                    throw new InvalidDataException($"The BrtHLink record at offset {record.Offset} refers to missing or invalid hyperlink relationship '{relationshipId}'.");
                }
                externalTarget = relationship.Target;
            } else if (string.IsNullOrWhiteSpace(location)) {
                throw new InvalidDataException($"The BrtHLink record at offset {record.Offset} has neither an external relationship nor an internal location.");
            }

            return new XlsbHyperlink(range, relationshipId, externalTarget, location, tooltip, display);
        }

        private static XlsbWorksheetFormatInfo ParseWorksheetFormatInfo(XlsbRecord record) {
            if (record.Data.Length != 12) {
                throw new InvalidDataException($"The BrtWsFmtInfo record at offset {record.Offset} has invalid payload length {record.Data.Length}.");
            }

            var cursor = new XlsbBinaryCursor(record.Data);
            uint encodedColumnWidth = cursor.ReadUInt32();
            ushort fallbackColumnWidth = cursor.ReadUInt16();
            ushort rowHeightTwips = cursor.ReadUInt16();
            uint flags = cursor.ReadUInt32();
            if ((flags & 0x0000FFFCU) != 0) {
                throw new InvalidDataException($"The BrtWsFmtInfo record at offset {record.Offset} sets reserved flags.");
            }

            byte maximumRowOutlineLevel = checked((byte)((flags >> 16) & 0xFFU));
            byte maximumColumnOutlineLevel = checked((byte)((flags >> 24) & 0xFFU));
            if (maximumRowOutlineLevel > 7 || maximumColumnOutlineLevel > 7) {
                throw new InvalidDataException($"The BrtWsFmtInfo record at offset {record.Offset} contains an invalid outline level.");
            }

            double columnWidth = encodedColumnWidth == uint.MaxValue
                ? fallbackColumnWidth
                : encodedColumnWidth / 256D;
            if (columnWidth < 0D || columnWidth > 255D) {
                throw new InvalidDataException($"The BrtWsFmtInfo record at offset {record.Offset} contains invalid default column width {columnWidth}.");
            }

            return new XlsbWorksheetFormatInfo(
                columnWidth,
                rowHeightTwips / 20D,
                (flags & 0x01U) != 0,
                (flags & 0x02U) != 0,
                maximumRowOutlineLevel,
                maximumColumnOutlineLevel);
        }

        private static XlsbPaneInfo ParsePane(XlsbRecord record) {
            if (record.Data.Length != 29) {
                throw new InvalidDataException($"The BrtPane record at offset {record.Offset} has invalid payload length {record.Data.Length}.");
            }

            var cursor = new XlsbBinaryCursor(record.Data);
            double horizontalSplit = cursor.ReadDouble();
            double verticalSplit = cursor.ReadDouble();
            uint topRow = cursor.ReadUInt32();
            uint leftColumn = cursor.ReadUInt32();
            uint activePane = cursor.ReadUInt32();
            byte flags = cursor.ReadByte();
            if (double.IsNaN(horizontalSplit) || double.IsInfinity(horizontalSplit) || horizontalSplit < 0D
                || double.IsNaN(verticalSplit) || double.IsInfinity(verticalSplit) || verticalSplit < 0D
                || topRow >= A1.MaxRows || leftColumn >= A1.MaxColumns || activePane > 3U || (flags & 0xFC) != 0) {
                throw new InvalidDataException($"The BrtPane record at offset {record.Offset} contains invalid pane metadata.");
            }

            return new XlsbPaneInfo(
                horizontalSplit,
                verticalSplit,
                checked((int)topRow),
                checked((int)leftColumn),
                activePane,
                (flags & 0x01) != 0,
                (flags & 0x02) != 0);
        }

        private static XlsbColumnInfo ParseColumnInfo(XlsbRecord record, XlsbWorkbook workbook) {
            if (record.Data.Length != 18) {
                throw new InvalidDataException($"The BrtColInfo record at offset {record.Offset} has invalid payload length {record.Data.Length}.");
            }

            var cursor = new XlsbBinaryCursor(record.Data);
            uint firstColumn = cursor.ReadUInt32();
            uint lastColumn = cursor.ReadUInt32();
            uint encodedWidth = cursor.ReadUInt32();
            uint styleIndex = cursor.ReadUInt32();
            ushort flags = cursor.ReadUInt16();
            if (firstColumn > lastColumn || lastColumn >= A1.MaxColumns || encodedWidth > 255U * 256U || (flags & ~0x170F) != 0) {
                throw new InvalidDataException($"The BrtColInfo record at offset {record.Offset} contains invalid column metadata.");
            }
            ValidateStyleIndex(styleIndex, workbook, $"The BrtColInfo record at offset {record.Offset}");

            return new XlsbColumnInfo(
                checked((int)firstColumn + 1),
                checked((int)lastColumn + 1),
                encodedWidth / 256D,
                styleIndex,
                (flags & 0x0001) != 0,
                (flags & 0x0002) != 0,
                (flags & 0x0004) != 0,
                (flags & 0x0008) != 0,
                checked((byte)((flags >> 8) & 0x07)),
                (flags & 0x1000) != 0);
        }

        private static XlsbRowInfo ParseRowInfo(XlsbRecord record, XlsbWorkbook workbook) {
            if (record.Data.Length < 17) {
                throw new InvalidDataException($"The BrtRowHdr record at offset {record.Offset} is truncated.");
            }

            var cursor = new XlsbBinaryCursor(record.Data);
            uint zeroBasedRow = cursor.ReadUInt32();
            uint styleIndex = cursor.ReadUInt32();
            ushort heightTwips = cursor.ReadUInt16();
            byte extraFlags = cursor.ReadByte();
            byte flags = cursor.ReadByte();
            byte phoneticFlags = cursor.ReadByte();
            uint spanCount = cursor.ReadUInt32();
            if (zeroBasedRow >= A1.MaxRows || (extraFlags & 0xFC) != 0 || (flags & 0x80) != 0 || (phoneticFlags & 0xFE) != 0 || spanCount > 16) {
                throw new InvalidDataException($"The BrtRowHdr record at offset {record.Offset} contains invalid row metadata.");
            }
            if (cursor.Remaining != checked((int)spanCount * 8)) {
                throw new InvalidDataException($"The BrtRowHdr record at offset {record.Offset} has an invalid column-span payload.");
            }
            ValidateStyleIndex(styleIndex, workbook, $"The BrtRowHdr record at offset {record.Offset}");

            var row = new XlsbRowInfo(
                checked((int)zeroBasedRow + 1),
                styleIndex,
                heightTwips,
                checked((byte)(flags & 0x07)),
                (flags & 0x08) != 0,
                (flags & 0x10) != 0,
                (flags & 0x20) != 0,
                (flags & 0x40) != 0,
                (phoneticFlags & 0x01) != 0);
            int previousLast = -1;
            for (uint index = 0; index < spanCount; index++) {
                uint firstColumn = cursor.ReadUInt32();
                uint lastColumn = cursor.ReadUInt32();
                if (firstColumn > lastColumn
                    || lastColumn >= A1.MaxColumns
                    || firstColumn / 1024U != lastColumn / 1024U
                    || firstColumn <= previousLast) {
                    throw new InvalidDataException($"The BrtRowHdr record at offset {record.Offset} contains an invalid column span.");
                }
                previousLast = checked((int)lastColumn);
                row.AddSpan(checked((int)firstColumn), checked((int)lastColumn));
            }
            return row;
        }

        private static void ValidateStyleIndex(uint styleIndex, XlsbWorkbook workbook, string context) {
            int availableFormats = workbook.Stylesheet?.CellFormats.Count ?? 1;
            if (styleIndex >= availableFormats) {
                throw new InvalidDataException($"{context} refers to missing cell format {styleIndex}; the styles part exposes {availableFormats} format(s).");
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
