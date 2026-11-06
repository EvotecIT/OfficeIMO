using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.Legacy;

internal sealed partial class EarlyBiffWorkbook {
    private LegacyXlsWorkbook ReadStyles(List<Record> records, ushort bof, Encoding encoding, string[] formats) {
        var workbook = new LegacyXlsWorkbook();
        // Excel 2-4 uses the first sixteen entries of the shared standard palette.
        workbook.MutablePaletteColors.AddRange(BiffColorPalette.DefaultPaletteColors.Take(16));
        foreach (Record record in records) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (record.Type is 0x0031 or 0x0231) {
                bool biff2 = record.Type == 0x0031;
                int index = workbook.Fonts.Count;
                if (index >= 4) index++;
                ushort flags = ReadUInt16(record, 2);
                double size = ReadUInt16(record, 0) / 20d;
                if (size <= 0) throw new InvalidDataException("An early BIFF font has a nonpositive size.");
                workbook.MutableFonts.Add(new LegacyXlsFont((ushort)index,
                    ReadText(record, biff2 ? 4 : 6, 1, encoding), size,
                    biff2 ? (ushort)0x7FFF : ReadUInt16(record, 4),
                    (flags & 1) != 0, (flags & 2) != 0, (flags & 4) != 0, (flags & 8) != 0,
                    outline: (flags & 0x10) != 0, shadow: (flags & 0x20) != 0));
            } else if (record.Type == 0x0045) {
                RequireLength(record, 2);
                if (workbook.Fonts.Count == 0) throw new InvalidDataException("An early BIFF font colour has no font.");
                LegacyXlsFont old = workbook.Fonts[workbook.Fonts.Count - 1];
                workbook.MutableFonts[workbook.Fonts.Count - 1] = new LegacyXlsFont(old.FontIndex,
                    old.Name, old.Size, ReadUInt16(record, 0), old.Bold, old.Italic, old.Underline, old.Strikeout);
            } else if (record.Type == 0x0092) {
                int count = ReadUInt16(record, 0);
                if (count != 16 || record.Length != 2 + count * 4) throw new InvalidDataException("Invalid early BIFF palette.");
                workbook.MutablePaletteColors.Clear();
                for (int i = 0; i < count; i++) {
                    int offset = record.Offset + 2 + i * 4;
                    workbook.MutablePaletteColors.Add($"FF{_data[offset]:X2}{_data[offset + 1]:X2}{_data[offset + 2]:X2}");
                }
            } else {
                workbook.MutableCellFormats.Add(ReadCellFormat(record, bof, formats, (ushort)workbook.CellFormats.Count));
            }
        }
        return workbook;
    }

    private LegacyXlsCellFormat ReadCellFormat(Record record, ushort bof, string[] formats, ushort index) {
        bool biff2 = bof == 0x0009;
        bool biff3 = bof == 0x0209;
        RequireLength(record, biff2 ? 4 : 12);
        int offset = record.Offset;
        if (biff2) {
            LegacyXlsCellFormat inline = CreateBiff2CellFormat(_data[offset], _data[offset + 2], _data[offset + 3], formats, index);
            if (inline.NumberFormatId != 0 && inline.NumberFormatCode == null) AddLoss("early-biff-number-format", "Formatting", "A referenced early BIFF number format could not be resolved.");
            return inline;
        }
        byte font = _data[offset];
        int number = _data[offset + 1];
        byte type = _data[offset + 2];
        bool style = (type & 4) != 0;
        ushort parent = (ushort)(ReadUInt16(record, biff3 ? 4 : 2) >> 4);
        byte used = _data[offset + (biff3 ? 3 : 5)];
        byte alignment = _data[offset + 4];
        ushort background = ReadUInt16(record, 6);
        uint border = BiffRecordReader.ReadUInt32(_data, offset + 8);
        byte vertical = biff3 ? (byte)2 : (byte)((alignment >> 4) & 3);
        int orientation = biff3 ? 0 : alignment >> 6;
        byte rotation = orientation switch { 1 => 255, 2 => 90, 3 => 180, _ => 0 };
        var lines = new LegacyXlsBorder((byte)((border >> 8) & 7), (byte)((border >> 24) & 7),
            (byte)(border & 7), (byte)((border >> 16) & 7),
            (ushort)((border >> 11) & 31), (ushort)(border >> 27),
            (ushort)((border >> 3) & 31), (ushort)((border >> 19) & 31), 0, 0, false, false);
        string? format = number < formats.Length ? formats[number] : null;
        if (number != 0 && format == null) AddLoss("early-biff-number-format", "Formatting", "A referenced early BIFF number format could not be resolved.");
        return new LegacyXlsCellFormat(index, font, (ushort)number, style, parent,
            Applies(4), Applies(8), Applies(64), (byte)(background & 0x3F),
            (ushort)((background >> 6) & 31), (ushort)(background >> 11),
            Applies(32), Applies(16), (byte)(alignment & 7), vertical,
            (alignment & 8) != 0, rotation, 0, false, 0, Applies(128),
            (type & 1) != 0, (type & 2) != 0, false, lines, format, false, false);

        bool Applies(int mask) => ((used & mask) != 0) != style;
    }

    private static LegacyXlsCellFormat CreateBiff2CellFormat(byte font, byte formatAndProtection, byte alignment, string[] formats, ushort index) {
        int number = formatAndProtection & 0x3F;
        return new LegacyXlsCellFormat(index, font, (ushort)number, false, ushort.MaxValue,
            true, true, true, (byte)((alignment & 0x80) == 0 ? 0 : 17), 8, 9,
            true, true, (byte)(alignment & 7), 2, false, 0, 0, false, 0, true,
            (formatAndProtection & 0x40) != 0, (formatAndProtection & 0x80) != 0, false,
            new LegacyXlsBorder((byte)((alignment >> 3) & 1), (byte)((alignment >> 4) & 1),
                (byte)((alignment >> 5) & 1), (byte)((alignment >> 6) & 1), 8, 8, 8, 8, 0, 0, false, false),
            number < formats.Length ? formats[number] : null, false, false);
    }

    private ushort ResolveBiff2InlineStyle(Record record, int attributeOffset, LegacyXlsWorkbook workbook, Dictionary<int, ushort> styles, string[] formats, string sheet, int row, int column) {
        byte protection = _data[record.Offset + attributeOffset];
        byte fontAndFormat = _data[record.Offset + attributeOffset + 1];
        byte alignment = _data[record.Offset + attributeOffset + 2];
        if ((protection & 0x3F) != 0) {
            AddCellLoss("early-biff-inline-style", "Formatting", "A BIFF2 worksheet without XF records has a nonzero style-index slot. Its defined inline attributes were retained.", sheet, row, column);
        }
        int number = fontAndFormat & 0x3F;
        if (number != 0 && number >= formats.Length) {
            AddCellLoss("early-biff-number-format", "Formatting", "A referenced early BIFF number format could not be resolved.", sheet, row, column);
        }
        int key = (protection & 0xC0) | (fontAndFormat << 8) | (alignment << 16);
        if (styles.TryGetValue(key, out ushort index)) return index;
        if (workbook.CellFormats.Count >= ushort.MaxValue) throw new InvalidDataException("Early BIFF exceeds the supported style-record count.");
        index = (ushort)workbook.CellFormats.Count;
        workbook.MutableCellFormats.Add(CreateBiff2CellFormat((byte)(fontAndFormat >> 6), (byte)(number | (protection & 0xC0)), alignment, formats, index));
        styles.Add(key, index);
        return index;
    }

    private ushort? ResolveStyle(LegacyXlsWorkbook workbook, Dictionary<ushort, LegacyXlsCellFormat> effectiveStyles, int index, string sheet, int row, int column) {
        if (index >= workbook.CellFormats.Count) {
            if (index != 0) AddCellLoss("early-biff-cell-style", "Formatting", "A cell references a missing XF record.", sheet, row, column);
            return null;
        }
        LegacyXlsCellFormat style = workbook.CellFormats[index];
        if (!style.IsStyle && style.ParentStyleIndex != 0xFFF && style.ParentStyleIndex != ushort.MaxValue
            && (style.ParentStyleIndex >= workbook.CellFormats.Count || !workbook.CellFormats[style.ParentStyleIndex].IsStyle)) {
            AddCellLoss("early-biff-style-inheritance", "Formatting", "The XF parent is missing or is not a style XF. Owned facets were retained.", sheet, row, column);
        }
        if (!effectiveStyles.TryGetValue((ushort)index, out LegacyXlsCellFormat? effective)) {
            effective = workbook.GetEffectiveCellFormat((ushort)index)!;
            effectiveStyles.Add((ushort)index, effective);
        }
        if (workbook.GetFont(effective.FontIndex) == null && workbook.Fonts.Count != 0) {
            AddCellLoss("early-biff-font-reference", "Formatting", "The XF references a missing font. Other supported style facets were retained.", sheet, row, column);
        }
        if (effective.FillPattern > 18) {
            AddCellLoss("early-biff-fill-pattern", "Formatting", "The XF fill pattern has no supported Excel representation.", sheet, row, column);
        }
        return (ushort)index;
    }

    private void AddCellLoss(string code, string category, string message, string sheet, int row, int column) {
        if (message.Length > _limits.MaxTextCharacters - _textCharacters) throw new InvalidDataException("Early BIFF input exceeds the configured text limit.");
        _textCharacters += message.Length;
        _model.Findings.Add(new OfficeCompatibilityFinding(code, category, message, OfficeCompatibilityState.Approximated,
            OfficeCompatibilitySeverity.Warning, OfficeCompatibilityImpact.Semantic | OfficeCompatibilityImpact.Visual, true,
            "'" + sheet.Replace("'", "''") + "'!" + ExcelA1.CellReference(row + 1, column + 1)));
    }
}
