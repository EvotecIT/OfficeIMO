using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Styles {
    /// <summary>Reads the core XLSB styles collections used by normal worksheet cells.</summary>
    internal static class XlsbStylesheetReader {
        private const int BrtBeginStyleSheet = 278;
        private const int BrtEndStyleSheet = 279;
        private const int BrtBeginFills = 603;
        private const int BrtEndFills = 604;
        private const int BrtBeginFonts = 611;
        private const int BrtEndFonts = 612;
        private const int BrtBeginBorders = 613;
        private const int BrtEndBorders = 614;
        private const int BrtBeginFmts = 615;
        private const int BrtEndFmts = 616;
        private const int BrtBeginCellXfs = 617;
        private const int BrtEndCellXfs = 618;
        private const int BrtBeginCellStyleXfs = 626;
        private const int BrtEndCellStyleXfs = 627;
        private const int BrtFont = 43;
        private const int BrtFmt = 44;
        private const int BrtFill = 45;
        private const int BrtBorder = 46;
        private const int BrtXf = 47;
        private const int BrtStyle = 48;
        private const int BrtBeginStyles = 619;
        private const int BrtEndStyles = 620;
        private const int BrtBeginDxfs = 505;
        private const int BrtEndDxfs = 506;
        private const int BrtFrtBegin = 35;
        private const int BrtFrtEnd = 36;
        private const int BrtAcBegin = 37;
        private const int BrtAcEnd = 38;
        private const int MaxStyleItems = 65_536;
        private const int MaxFonts = 0xFFD3;
        private const int MaxCellFormats = 0xFF96;

        internal static XlsbStylesheet Read(
            byte[] bytes,
            string partName,
            XlsbImportOptions options,
            XlsbWorkbook workbook,
            XlsbRecordReadBudget budget) {
            ArgumentNullException.ThrowIfNull(bytes);
            if (string.IsNullOrWhiteSpace(partName)) throw new ArgumentNullException(nameof(partName));
            ArgumentNullException.ThrowIfNull(options);
            ArgumentNullException.ThrowIfNull(workbook);

            IReadOnlyList<XlsbRecord> records;
            using (var stream = new MemoryStream(bytes, writable: false)) {
                records = XlsbRecordReader.ReadAll(stream, options.MaxRecordBytes, budget);
            }

            if (records.Count < 2
                || records[0].Type != BrtBeginStyleSheet
                || records[records.Count - 1].Type != BrtEndStyleSheet) {
                throw new InvalidDataException($"The XLSB styles part '{partName}' is missing its BrtBeginStyleSheet/BrtEndStyleSheet boundaries.");
            }

            var stylesheet = new XlsbStylesheet();
            CollectionKind active = CollectionKind.None;
            int expectedCount = 0;
            int actualCount = 0;
            var styleRecords = new List<XlsbRecord>();
            bool inStyles = false;
            var differentialFormats = new DifferentialFormatCollectionReader(stylesheet, workbook, partName, options);
            foreach (XlsbRecord record in records) {
                if (active == CollectionKind.None && differentialFormats.TryRead(record)) continue;
                switch (record.Type) {
                    case BrtBeginStyleSheet:
                        if (record.Offset != 0 || record.Data.Length != 0) {
                            throw new InvalidDataException($"The XLSB styles part '{partName}' contains an invalid BrtBeginStyleSheet record.");
                        }
                        break;
                    case BrtEndStyleSheet:
                        if (record.Offset + record.HeaderSize + record.Size != bytes.Length
                            || record.Data.Length != 0
                            || active != CollectionKind.None) {
                            throw new InvalidDataException($"The XLSB styles part '{partName}' contains an invalid BrtEndStyleSheet record.");
                        }
                        break;
                    case BrtBeginFmts:
                        BeginCollection(record, CollectionKind.NumberFormats, ref active, ref expectedCount, ref actualCount);
                        break;
                    case BrtEndFmts:
                        EndCollection(record, CollectionKind.NumberFormats, ref active, expectedCount, actualCount);
                        break;
                    case BrtBeginFonts:
                        BeginCollection(record, CollectionKind.Fonts, ref active, ref expectedCount, ref actualCount);
                        break;
                    case BrtEndFonts:
                        EndCollection(record, CollectionKind.Fonts, ref active, expectedCount, actualCount);
                        break;
                    case BrtBeginFills:
                        BeginCollection(record, CollectionKind.Fills, ref active, ref expectedCount, ref actualCount);
                        break;
                    case BrtEndFills:
                        EndCollection(record, CollectionKind.Fills, ref active, expectedCount, actualCount);
                        break;
                    case BrtBeginBorders:
                        BeginCollection(record, CollectionKind.Borders, ref active, ref expectedCount, ref actualCount);
                        break;
                    case BrtEndBorders:
                        EndCollection(record, CollectionKind.Borders, ref active, expectedCount, actualCount);
                        break;
                    case BrtBeginCellStyleXfs:
                        BeginCollection(record, CollectionKind.CellStyleFormats, ref active, ref expectedCount, ref actualCount);
                        break;
                    case BrtEndCellStyleXfs:
                        EndCollection(record, CollectionKind.CellStyleFormats, ref active, expectedCount, actualCount);
                        break;
                    case BrtBeginCellXfs:
                        BeginCollection(record, CollectionKind.CellFormats, ref active, ref expectedCount, ref actualCount);
                        break;
                    case BrtEndCellXfs:
                        EndCollection(record, CollectionKind.CellFormats, ref active, expectedCount, actualCount);
                        break;
                    case BrtFmt when active == CollectionKind.NumberFormats:
                        ReadNumberFormat(record, options, stylesheet);
                        actualCount++;
                        break;
                    case BrtFont when active == CollectionKind.Fonts:
                        stylesheet.AddFont(ReadFont(record, options));
                        actualCount++;
                        break;
                    case BrtFill when active == CollectionKind.Fills:
                        XlsbFill fill = ReadFill(record);
                        stylesheet.AddFill(fill);
                        actualCount++;
                        if (fill.GradientStopCount > 0) {
                            workbook.AddDiagnostic(new XlsbImportDiagnostic(
                                "xlsb-style-gradient-preserved",
                                XlsbImportDiagnosticSeverity.Warning,
                                "Preserved an XLSB gradient fill whose stop collection is not yet projected into the editable Open XML style model.",
                                partName,
                                record.Type,
                                record.Offset));
                            PreserveRecord(options, workbook, partName, record);
                        }
                        break;
                    case BrtBorder when active == CollectionKind.Borders:
                        stylesheet.AddBorder(ReadBorder(record));
                        actualCount++;
                        break;
                    case BrtXf when active == CollectionKind.CellStyleFormats:
                        stylesheet.AddCellStyleFormat(ReadCellFormat(record));
                        actualCount++;
                        break;
                    case BrtXf when active == CollectionKind.CellFormats:
                        stylesheet.AddCellFormat(ReadCellFormat(record));
                        actualCount++;
                        break;
                    case BrtBeginStyles when active == CollectionKind.None && !inStyles:
                    case BrtEndStyles when inStyles:
                        inStyles = record.Type == BrtBeginStyles;
                        styleRecords.Add(record);
                        break;
                    case BrtStyle when inStyles:
                        styleRecords.Add(record);
                        break;
                    default:
                        PreserveRecord(options, workbook, partName, record);
                        break;
                }
            }

            if (active != CollectionKind.None) {
                throw new InvalidDataException($"The XLSB styles part '{partName}' ended inside the {active} collection.");
            }
            differentialFormats.Complete();

            ValidateReferences(stylesheet, partName);
            ProjectCellStyles(stylesheet, styleRecords, options);
            if (stylesheet.CellStyles.Count == 0) {
                foreach (XlsbRecord record in styleRecords) PreserveRecord(options, workbook, partName, record);
            }
            return stylesheet;
        }

        // Projects BrtStyle (MS-XLSB 2.4.811). A collection that cannot be represented as Open XML cell styles
        // stays preserved as native records instead of failing the load.
        private static void ProjectCellStyles(XlsbStylesheet stylesheet, IReadOnlyList<XlsbRecord> records, XlsbImportOptions options) {
            if (records.Count < 2 || records[0].Type != BrtBeginStyles || records[records.Count - 1].Type != BrtEndStyles
                || records[0].Data.Length != 4) return;
            uint declared = new XlsbBinaryCursor(records[0].Data).ReadUInt32();
            if (declared != records.Count - 2) return;
            var styles = new List<XlsbCellStyle>(records.Count - 2);
            var formatIds = new HashSet<uint>();
            for (int index = 1; index < records.Count - 1; index++) {
                var cursor = new XlsbBinaryCursor(records[index].Data);
                if (cursor.Remaining < 12) return;
                uint formatId = cursor.ReadUInt32();
                ushort flags = cursor.ReadUInt16();
                byte builtInId = cursor.ReadByte();
                byte level = cursor.ReadByte();
                string name;
                try {
                    name = cursor.ReadWideString(Math.Min(options.MaxStringCharacters, 255));
                } catch (Exception exception) when (exception is InvalidDataException or EndOfStreamException) {
                    return;
                }
                bool builtIn = (flags & 0x0001) != 0;
                if (cursor.Remaining != 0 || name.Length == 0 || formatId >= stylesheet.CellStyleFormats.Count
                    || !formatIds.Add(formatId) || (!builtIn && (builtInId != 0 || (flags & 0x0004) != 0))) return;
                styles.Add(new XlsbCellStyle(formatId, builtIn, (flags & 0x0002) != 0, (flags & 0x0004) != 0, builtInId, level, name));
            }
            stylesheet.CellStyles = styles;
        }

        private static void BeginCollection(
            XlsbRecord record,
            CollectionKind kind,
            ref CollectionKind active,
            ref int expectedCount,
            ref int actualCount) {
            if (active != CollectionKind.None) {
                throw new InvalidDataException($"The XLSB styles part begins {kind} inside the active {active} collection.");
            }

            if (record.Data.Length != 4) {
                throw new InvalidDataException($"The XLSB {kind} collection header has invalid payload length {record.Data.Length}.");
            }

            var cursor = new XlsbBinaryCursor(record.Data);
            uint declared = cursor.ReadUInt32();
            int maximum = kind == CollectionKind.Fonts
                ? MaxFonts
                : kind == CollectionKind.CellFormats
                    ? MaxCellFormats
                    : MaxStyleItems;
            if (declared > maximum) {
                throw new InvalidDataException($"The XLSB {kind} collection declares {declared} items, exceeding the supported limit of {maximum}.");
            }

            active = kind;
            expectedCount = checked((int)declared);
            actualCount = 0;
        }

        private static void EndCollection(
            XlsbRecord record,
            CollectionKind kind,
            ref CollectionKind active,
            int expectedCount,
            int actualCount) {
            if (record.Data.Length != 0 || active != kind) {
                throw new InvalidDataException($"The XLSB styles part contains an invalid end marker for the {kind} collection.");
            }

            if (actualCount != expectedCount) {
                throw new InvalidDataException($"The XLSB {kind} collection declares {expectedCount} items but contains {actualCount} supported item records.");
            }

            active = CollectionKind.None;
        }

        private static void ReadNumberFormat(XlsbRecord record, XlsbImportOptions options, XlsbStylesheet stylesheet) {
            var cursor = new XlsbBinaryCursor(record.Data);
            ushort id = cursor.ReadUInt16();
            string code = cursor.ReadWideString(Math.Min(options.MaxStringCharacters, 255));
            if (cursor.Remaining != 0 || code.Length == 0) {
                throw new InvalidDataException($"The XLSB number format record at offset {record.Offset} is malformed.");
            }

            try {
                stylesheet.AddNumberFormat(id, code);
            } catch (ArgumentException exception) {
                throw new InvalidDataException($"The XLSB styles part contains duplicate number format id {id}.", exception);
            }
        }

        private static XlsbFont ReadFont(XlsbRecord record, XlsbImportOptions options) {
            var cursor = new XlsbBinaryCursor(record.Data);
            var font = new XlsbFont {
                HeightTwips = cursor.ReadUInt16(),
                Flags = cursor.ReadUInt16(),
                Weight = cursor.ReadUInt16(),
                Script = cursor.ReadUInt16(),
                Underline = cursor.ReadByte(),
                Family = cursor.ReadByte(),
                CharacterSet = cursor.ReadByte()
            };
            cursor.Skip(1);
            font.Color = ReadColor(cursor);
            font.Scheme = cursor.ReadByte();
            font.Name = cursor.ReadWideString(Math.Min(options.MaxStringCharacters, 31));
            if (cursor.Remaining != 0 || font.Name.Length == 0) {
                throw new InvalidDataException($"The XLSB font record at offset {record.Offset} is malformed.");
            }
            return font;
        }

        private static XlsbFill ReadFill(XlsbRecord record) {
            try {
                var cursor = new XlsbBinaryCursor(record.Data);
                var fill = new XlsbFill {
                    Pattern = cursor.ReadUInt32(),
                    Foreground = ReadColor(cursor),
                    Background = ReadColor(cursor),
                    GradientType = cursor.ReadInt32()
                };
                cursor.Skip(8 * 5);
                fill.GradientStopCount = cursor.ReadUInt32();
                if (fill.GradientStopCount > MaxStyleItems) {
                    throw new InvalidDataException($"The XLSB fill record at offset {record.Offset} declares too many gradient stops.");
                }
                for (uint index = 0; index < fill.GradientStopCount; index++) {
                    cursor.ReadDouble();
                    ReadColor(cursor);
                }
                if (cursor.Remaining != 0) {
                    throw new InvalidDataException($"The XLSB fill record at offset {record.Offset} has unexpected trailing data.");
                }
                return fill;
            } catch (EndOfStreamException exception) {
                throw new InvalidDataException(
                    $"The BrtFill record at offset {record.Offset} is truncated.",
                    exception);
            }
        }

        private static XlsbBorder ReadBorder(XlsbRecord record) {
            var cursor = new XlsbBinaryCursor(record.Data);
            byte flags = cursor.ReadByte();
            var border = new XlsbBorder {
                DiagonalDown = (flags & 0x01) != 0,
                DiagonalUp = (flags & 0x02) != 0,
                Top = ReadBorderSide(cursor),
                Bottom = ReadBorderSide(cursor),
                Left = ReadBorderSide(cursor),
                Right = ReadBorderSide(cursor),
                Diagonal = ReadBorderSide(cursor)
            };
            if (cursor.Remaining != 0) {
                throw new InvalidDataException($"The XLSB border record at offset {record.Offset} has unexpected trailing data.");
            }
            return border;
        }

        private static XlsbBorderSide ReadBorderSide(XlsbBinaryCursor cursor) {
            byte style = cursor.ReadByte();
            cursor.Skip(1);
            return new XlsbBorderSide(style, ReadColor(cursor));
        }

        private static XlsbCellFormat ReadCellFormat(XlsbRecord record) {
            if (record.Data.Length != 16) {
                throw new InvalidDataException($"The XLSB cell format record at offset {record.Offset} has invalid payload length {record.Data.Length}.");
            }

            var cursor = new XlsbBinaryCursor(record.Data);
            var format = new XlsbCellFormat {
                ParentFormatId = cursor.ReadUInt16(),
                NumberFormatId = cursor.ReadUInt16(),
                FontId = cursor.ReadUInt16(),
                FillId = cursor.ReadUInt16(),
                BorderId = cursor.ReadUInt16(),
                TextRotation = cursor.ReadByte(),
                Indent = cursor.ReadByte()
            };
            byte alignment = cursor.ReadByte();
            format.HorizontalAlignment = (byte)(alignment & 0x07);
            format.VerticalAlignment = (byte)((alignment >> 3) & 0x07);
            format.WrapText = (alignment & 0x40) != 0;
            format.JustifyLastLine = (alignment & 0x80) != 0;
            byte protection = cursor.ReadByte();
            format.ShrinkToFit = (protection & 0x01) != 0;
            format.Merged = (protection & 0x02) != 0;
            format.ReadingOrder = (byte)((protection >> 2) & 0x03);
            format.Locked = (protection & 0x10) != 0;
            format.Hidden = (protection & 0x20) != 0;
            format.PivotButton = (protection & 0x40) != 0;
            format.QuotePrefix = (protection & 0x80) != 0;
            format.ApplyFlags = (byte)(cursor.ReadUInt16() & 0x3F);
            return format;
        }

        internal static XlsbColor ReadColor(XlsbBinaryCursor cursor) {
            byte flags = cursor.ReadByte();
            byte index = cursor.ReadByte();
            short tint = cursor.ReadInt16();
            return new XlsbColor(
                (byte)(flags >> 1),
                index,
                tint,
                cursor.ReadByte(),
                cursor.ReadByte(),
                cursor.ReadByte(),
                cursor.ReadByte());
        }

        private static void ValidateReferences(XlsbStylesheet stylesheet, string partName) {
            if (stylesheet.Fonts.Count == 0
                || stylesheet.Fills.Count == 0
                || stylesheet.Borders.Count == 0
                || stylesheet.CellStyleFormats.Count == 0
                || stylesheet.CellFormats.Count == 0) {
                throw new InvalidDataException($"The XLSB styles part '{partName}' is missing one or more required formatting collections.");
            }

            foreach (XlsbCellFormat format in stylesheet.CellStyleFormats.Concat(stylesheet.CellFormats)) {
                if (format.FontId >= stylesheet.Fonts.Count
                    || format.FillId >= stylesheet.Fills.Count
                    || format.BorderId >= stylesheet.Borders.Count) {
                    throw new InvalidDataException($"The XLSB styles part '{partName}' contains a cell format with an out-of-range font, fill, or border reference.");
                }
                if (format.NumberFormatId >= ExcelBuiltInNumberFormats.FirstCustomId
                    && !stylesheet.NumberFormats.ContainsKey(format.NumberFormatId)) {
                    throw new InvalidDataException(
                        $"The XLSB styles part '{partName}' contains a cell format that references missing custom number format {format.NumberFormatId}.");
                }
            }

            if (stylesheet.CellStyleFormats.Any(format => format.ParentFormatId != ushort.MaxValue)
                || stylesheet.CellFormats.Any(format => format.ParentFormatId >= stylesheet.CellStyleFormats.Count)) {
                throw new InvalidDataException($"The XLSB styles part '{partName}' contains an invalid parent cell-style reference.");
            }
        }

        private static void PreserveRecord(
            XlsbImportOptions options,
            XlsbWorkbook workbook,
            string partName,
            XlsbRecord record) {
            if (options.ReportPreservedRecords) {
                workbook.AddPreservedRecord(new XlsbPreservedRecordInfo(partName, record.Type, record.Offset, record.Size));
            }
        }

        /// <summary>
        /// Reads DXFS (MS-XLSB 2.1.7.50): BrtBeginDXFs, then per DXF either a BrtDXF with optional FRT records or two
        /// alternate-content blocks (BrtACBegin BrtDXF BrtACEnd BrtACBegin FRTDXF BrtACEnd), where the second block wins.
        /// </summary>
        private sealed class DifferentialFormatCollectionReader {
            private readonly XlsbStylesheet _stylesheet;
            private readonly XlsbWorkbook _workbook;
            private readonly string _partName;
            private readonly XlsbImportOptions _options;
            private bool _active;
            private bool _seen;
            private uint _declared;
            private int _alternateDepth;
            private int _alternateBlocks;
            private int _futureDepth;
            private DocumentFormat.OpenXml.Spreadsheet.DifferentialFormat? _pending;

            internal DifferentialFormatCollectionReader(XlsbStylesheet stylesheet, XlsbWorkbook workbook, string partName, XlsbImportOptions options) {
                _stylesheet = stylesheet;
                _workbook = workbook;
                _partName = partName;
                _options = options;
            }

            internal bool TryRead(XlsbRecord record) {
                if (!_active) {
                    if (record.Type != BrtBeginDxfs) return false;
                    if (_seen || record.Data.Length != 4) {
                        throw new InvalidDataException($"The XLSB styles part '{_partName}' contains an invalid or repeated differential format collection.");
                    }
                    _declared = BitConverter.ToUInt32(record.Data, 0);
                    if (_declared > int.MaxValue) {
                        throw new InvalidDataException($"The XLSB styles part '{_partName}' declares too many differential formats.");
                    }
                    _seen = true;
                    _active = true;
                    return true;
                }
                if (_futureDepth > 0) {
                    if (record.Type == BrtFrtBegin) _futureDepth++;
                    else if (record.Type == BrtFrtEnd) _futureDepth--;
                    PreserveRecord(_options, _workbook, _partName, record);
                    return true;
                }
                switch (record.Type) {
                    case XlsbDifferentialFormatCodec.BrtDXF:
                        DocumentFormat.OpenXml.Spreadsheet.DifferentialFormat format = Decode(record);
                        if (_alternateDepth == 0) Add(format);
                        else _pending = format;
                        return true;
                    case BrtAcBegin:
                        _alternateDepth++;
                        return true;
                    case BrtAcEnd:
                        if (_alternateDepth == 0) throw new InvalidDataException($"The XLSB styles part '{_partName}' contains an unmatched BrtACEnd record.");
                        if (--_alternateDepth == 0 && ++_alternateBlocks == 2) {
                            if (_pending == null) {
                                throw new InvalidDataException($"The XLSB styles part '{_partName}' contains an alternate differential format without a BrtDXF record.");
                            }
                            Add(_pending);
                            _pending = null;
                            _alternateBlocks = 0;
                        }
                        return true;
                    case BrtFrtBegin:
                        _futureDepth = 1;
                        PreserveRecord(_options, _workbook, _partName, record);
                        return true;
                    case BrtEndDxfs:
                        if (_alternateDepth != 0 || _alternateBlocks != 0 || record.Data.Length != 0) {
                            throw new InvalidDataException($"The XLSB styles part '{_partName}' contains an incomplete differential format.");
                        }
                        if (_stylesheet.DifferentialFormats.Count != _declared) {
                            throw new InvalidDataException(
                                $"The XLSB differential format collection declares {_declared} items but contains {_stylesheet.DifferentialFormats.Count}.");
                        }
                        _active = false;
                        return true;
                    default:
                        throw new InvalidDataException($"The XLSB styles part '{_partName}' contains record {record.Type} inside the differential format collection.");
                }
            }

            internal void Complete() {
                if (_active) throw new InvalidDataException($"The XLSB styles part '{_partName}' ended inside the differential format collection.");
            }

            private void Add(DocumentFormat.OpenXml.Spreadsheet.DifferentialFormat format) {
                if (_stylesheet.DifferentialFormats.Count >= _declared) {
                    throw new InvalidDataException($"The XLSB differential format collection contains more than the {_declared} items it declares.");
                }
                _stylesheet.AddDifferentialFormat(format);
            }

            private DocumentFormat.OpenXml.Spreadsheet.DifferentialFormat Decode(XlsbRecord record) {
                try {
                    return XlsbDifferentialFormatCodec.Read(record.Data, type => _workbook.AddDiagnostic(new XlsbImportDiagnostic(
                        "xlsb-dxf-property-not-projected",
                        XlsbImportDiagnosticSeverity.Warning,
                        $"A differential format property of type 0x{type:X2} has no SpreadsheetML equivalent and is not projected.",
                        _partName,
                        record.Type,
                        record.Offset)));
                } catch (EndOfStreamException exception) {
                    throw new InvalidDataException($"The BrtDXF record at offset {record.Offset} in '{_partName}' is truncated.", exception);
                }
            }
        }

        private enum CollectionKind {
            None,
            NumberFormats,
            Fonts,
            Fills,
            Borders,
            CellStyleFormats,
            CellFormats
        }
    }
}
