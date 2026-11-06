using OfficeIMO.Excel.LegacyXls.Biff;

namespace OfficeIMO.Excel.Legacy;

internal sealed partial class EarlyBiffWorkbook {
    private void ReadWorksheet(ushort bof, int start, int end, string name, ushort? inheritedCodePage) {
        ushort? codePage = null;
        ushort? dateMode = null;
        int position = start;
        bool sawEnd = false;
        var formatRecords = new List<Record>();
        var styles = new List<StyleFormat>();
        while (position < end) {
            Record record = ReadRecord(ref position, end);
            if (IsBof(record.Type)) {
                SkipEmbeddedChart(record, ref position, end, charge: true);
                continue;
            }
            if (record.Type == FilePass) RejectEncrypted();
            if ((record.Type == 0x001E && bof != 0x0409) || (record.Type == 0x041E && bof == 0x0409)) {
                formatRecords.Add(record);
            }
            if ((record.Type, bof) is (0x0043, 0x0009) or (0x0243, 0x0209) or (0x0443, 0x0409)) {
                styles.Add(ReadStyle(record, bof));
                _styleOmitted = true;
            }
            if (record.Type == CodePage) codePage = ObserveCodePage(record, codePage);
            if (record.Type == DateMode) {
                RequireLength(record, 2);
                ushort declared = ReadUInt16(record, 0);
                if (declared > 1 || (dateMode.HasValue && dateMode.Value != declared)) {
                    throw new InvalidDataException("Invalid or conflicting early BIFF date system.");
                }
                dateMode = declared;
            }
            if (record.Type != EndOfFile) continue;
            RequireLength(record, 0);
            sawEnd = position == end;
            break;
        }
        if (!sawEnd) throw new InvalidDataException("An early BIFF worksheet has an invalid EOF boundary.");

        ExcelDateSystem dates = dateMode == 1 ? ExcelDateSystem.NineteenFour : ExcelDateSystem.NineteenHundred;
        if (_model.DateSystem.HasValue && _model.DateSystem.Value != dates) {
            throw new NotSupportedException("A BIFF4 workbook with different worksheet date systems cannot be projected to one Excel date system.");
        }
        _model.DateSystem = dates;

        Encoding encoding = BiffCodePageEncoding.Resolve(codePage ?? inheritedCodePage);
        var sheet = new LegacySpreadsheetSheet(name);
        _model.Sheets.Add(sheet);
        var coordinates = new HashSet<int>();
        string[] formats = formatRecords.Select(record =>
            ReadText(record, record.Type == 0x041E ? 2 : 0, 1, encoding)).ToArray();
        int previousCell = -1;
        bool outOfOrder = false;
        Record? pendingFormula = null;
        ushort? extendedStyle = null;
        bool biff2 = bof == 0x0009;
        position = start;
        while (position < end) {
            Record record = ReadRecord(ref position, end, charge: false);
            ushort type = record.Type;
            if (pendingFormula.HasValue) {
                if (type == (biff2 ? 0x0007 : 0x0207)) {
                    AddCell(pendingFormula.Value, ReadText(record, 0, biff2 ? 1 : 2, encoding), false);
                    pendingFormula = null;
                    extendedStyle = null;
                    continue;
                }
                if (type is 0x0021 or 0x0221 or 0x0036 or 0x0037 or 0x0236) {
                    ObserveUnprojected(record);
                    continue;
                }
                throw new InvalidDataException("An early BIFF string formula is missing its cached STRING record.");
            }
            switch (type) {
                case EndOfFile:
                case CodePage:
                case DateMode:
                    continue;
                case 0x001E when bof != 0x0409:
                case 0x041E when bof == 0x0409:
                case 0x0043 when biff2:
                case 0x0243 when bof == 0x0209:
                case 0x0443 when bof == 0x0409:
                    continue;
                case 0x0044 when biff2:
                    RequireLength(record, 2);
                    extendedStyle = ReadUInt16(record, 0);
                    continue;
                case 0x0001 when biff2:
                case 0x0201 when !biff2:
                    RequireLength(record, biff2 ? 7 : 6);
                    AddCell(record, null, false);
                    break;
                case 0x0002 when biff2:
                    RequireLength(record, 9);
                    AddCell(record, (double)ReadUInt16(record, 7), false);
                    break;
                case 0x0003 when biff2:
                case 0x0203 when !biff2:
                    RequireLength(record, biff2 ? 15 : 14);
                    AddCell(record, BiffRecordReader.ReadDouble(_data, record.Offset + (biff2 ? 7 : 6)), false);
                    break;
                case 0x027E when !biff2:
                    RequireLength(record, 10);
                    AddCell(record, BiffRkNumberReader.ReadRkNumber(BiffRecordReader.ReadUInt32(_data, record.Offset + 6)), false);
                    break;
                case 0x0004 when biff2:
                case 0x0204 when !biff2:
                    AddCell(record, ReadText(record, biff2 ? 7 : 6, biff2 ? 1 : 2, encoding), false);
                    break;
                case 0x0005 when biff2:
                case 0x0205 when !biff2:
                    RequireLength(record, biff2 ? 9 : 8);
                    int flag = _data[record.Offset + record.Length - 1];
                    byte value = _data[record.Offset + record.Length - 2];
                    AddCell(record, DecodeBooleanOrError(value, flag), flag == 1);
                    break;
                case 0x0006 when biff2:
                case 0x0206 when bof == 0x0209:
                case 0x0406 when bof == 0x0409:
                    _formulaOmitted = true;
                    int cache = biff2 ? 7 : 6;
                    int tokenLengthSize = biff2 ? 1 : 2;
                    if (record.Length < 16 + tokenLengthSize) throw new InvalidDataException("Truncated early BIFF formula.");
                    int tokenLength = biff2 ? _data[record.Offset + 16] : ReadUInt16(record, 16);
                    if (tokenLength > record.Length - 16 - tokenLengthSize) throw new InvalidDataException("Truncated early BIFF formula tokens.");
                    if (ReadUInt16(record, cache + 6) != 0xFFFF) {
                        AddCell(record, BiffRecordReader.ReadDouble(_data, record.Offset + cache), false);
                    } else {
                        byte cacheType = _data[record.Offset + cache];
                        byte result = _data[record.Offset + cache + 2];
                        switch (cacheType) {
                            case 0:
                                pendingFormula = record;
                                continue;
                            case 1:
                            case 2:
                                AddCell(record, DecodeBooleanOrError(result, cacheType - 1), cacheType == 2);
                                break;
                            case 3:
                                AddCell(record, string.Empty, false);
                                break;
                            default:
                                throw new InvalidDataException("Unknown early BIFF formula cache type.");
                        }
                    }
                    break;
                case 0x0007:
                case 0x0207:
                    throw new InvalidDataException("An early BIFF STRING record has no matching formula.");
                case 0x0009:
                case 0x0209:
                case 0x0409:
                case 0x0809:
                    SkipEmbeddedChart(record, ref position, end, charge: false);
                    continue;
                default:
                    ObserveUnprojected(record);
                    continue;
            }
            extendedStyle = null;
        }

        if (outOfOrder) {
            sheet.Cells.Sort(static (left, right) => left.Row != right.Row
                ? left.Row.CompareTo(right.Row) : left.Column.CompareTo(right.Column));
        }

        void AddCell(Record record, object? value, bool error) {
            int row = ReadUInt16(record, 0);
            int column = ReadUInt16(record, 2);
            int coordinate = row * MaximumColumns + column;
            if (row >= MaximumRows || column >= MaximumColumns || !coordinates.Add(coordinate)) {
                throw new InvalidDataException("The early BIFF worksheet contains an invalid or duplicate cell address.");
            }
            outOfOrder |= coordinate < previousCell;
            previousCell = coordinate;
            if (_model.RecoveredCellCount >= _limits.MaxItems) {
                throw new InvalidDataException("Early BIFF input exceeds the configured cell limit.");
            }
            if (value is double number && (double.IsNaN(number) || double.IsInfinity(number))) {
                throw new InvalidDataException("The early BIFF worksheet contains a non-finite numeric value.");
            }
            int xf = biff2 ? _data[record.Offset + 4] & 0x3F : ReadUInt16(record, 4);
            int formatIndex = 0;
            if (biff2) {
                if (xf == 63 && !extendedStyle.HasValue) {
                    throw new InvalidDataException("An early BIFF cell references a missing extended style index.");
                }
                int style = extendedStyle ?? xf;
                formatIndex = style < styles.Count ? styles[style].NumberFormat : _data[record.Offset + 5] & 0x3F;
            } else if (styles.Count != 0 || xf != 0) {
                formatIndex = ResolveNumberFormat(styles, xf);
            }
            string? format = formatIndex > 0 && formatIndex < formats.Length ? formats[formatIndex] : null;
            if (formatIndex > 0 && format == null) {
                AddLoss("early-biff-number-format", "Formatting", "A referenced early BIFF number format could not be resolved.");
            }
            sheet.Cells.Add(new LegacySpreadsheetCell(row + 1, column + 1, value, numberFormat: format, isError: error));
            _model.RecoveredCellCount++;
        }
    }

    private StyleFormat ReadStyle(Record record, ushort bof) {
        if (bof == 0x0009) {
            RequireLength(record, 4);
            return new StyleFormat(_data[record.Offset + 2] & 0x3F, -1, false, true);
        }

        RequireLength(record, 12);
        bool biff3 = bof == 0x0209;
        return new StyleFormat(_data[record.Offset + 1], ReadUInt16(record, biff3 ? 4 : 2) >> 4,
            (_data[record.Offset + 2] & 0x04) != 0,
            (_data[record.Offset + (biff3 ? 3 : 5)] & 0x04) != 0);
    }

    private int ResolveNumberFormat(List<StyleFormat> styles, int index) {
        if (index < styles.Count) {
            StyleFormat style = styles[index];
            if (style.IsStyle != style.UsesOwnFormat) return style.NumberFormat;
            if (!style.IsStyle && style.Parent < styles.Count) {
                StyleFormat parent = styles[style.Parent];
                if (parent.IsStyle && !parent.UsesOwnFormat) return parent.NumberFormat;
            }
        }
        AddLoss("early-biff-number-format", "Formatting", "A referenced early BIFF number format or parent style could not be resolved.");
        return 0;
    }

    private readonly struct StyleFormat {
        internal StyleFormat(int numberFormat, int parent, bool isStyle, bool usesOwnFormat) {
            NumberFormat = numberFormat;
            Parent = parent;
            IsStyle = isStyle;
            UsesOwnFormat = usesOwnFormat;
        }
        internal int NumberFormat { get; }
        internal int Parent { get; }
        internal bool IsStyle { get; }
        internal bool UsesOwnFormat { get; }
    }

    private static object DecodeBooleanOrError(byte value, int errorFlag) => errorFlag switch {
        0 when value <= 1 => value == 1,
        1 => BiffErrorValue.ToText(value),
        _ => throw new InvalidDataException("Invalid early BIFF boolean or error value.")
    };
}
