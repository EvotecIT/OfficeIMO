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
        private const int MaxStyleItems = 65_536;

        internal static XlsbStylesheet Read(
            byte[] bytes,
            string partName,
            XlsbImportOptions options,
            XlsbWorkbook workbook,
            XlsbRecordReadBudget budget) {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
            if (string.IsNullOrWhiteSpace(partName)) throw new ArgumentNullException(nameof(partName));
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));

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
            foreach (XlsbRecord record in records) {
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
                                "XLSB-STYLE-GRADIENT-PRESERVED",
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
                    default:
                        PreserveRecord(options, workbook, partName, record);
                        break;
                }
            }

            if (active != CollectionKind.None) {
                throw new InvalidDataException($"The XLSB styles part '{partName}' ended inside the {active} collection.");
            }

            ValidateReferences(stylesheet, partName);
            return stylesheet;
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
            if (declared > MaxStyleItems) {
                throw new InvalidDataException($"The XLSB {kind} collection declares {declared} items, exceeding the supported limit of {MaxStyleItems}.");
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
            if (fill.GradientStopCount == 0 && cursor.Remaining != 0) {
                throw new InvalidDataException($"The XLSB fill record at offset {record.Offset} has unexpected trailing data.");
            }
            return fill;
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

        private static XlsbColor ReadColor(XlsbBinaryCursor cursor) {
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
