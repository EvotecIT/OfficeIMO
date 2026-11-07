using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.Legacy;

internal sealed partial class EarlyBiffWorkbook {
    private bool ReadLayout(Record record, ushort bof, LegacySpreadsheetSheet sheet, Dictionary<int, ushort>? inlineStyles, string[] formats) {
        bool biff2 = bof == 0x0009;
        switch (record.Type) {
            case 0x0008 when biff2:
            case 0x0208 when !biff2:
                if (record.Length < (biff2 ? 11 : 16)) throw new InvalidDataException("Truncated early BIFF row layout.");
                int row = ReadUInt16(record, 0);
                if (row >= MaximumRows) throw new InvalidDataException("Invalid early BIFF row index.");
                ushort height = ReadUInt16(record, 6);
                uint flags = biff2 ? 0 : BiffRecordReader.ReadUInt32(_data, record.Offset + 12);
                if (biff2 && (_data[record.Offset + 10] & 1) != 0 && record.Length < 16) throw new InvalidDataException("Truncated BIFF2 row style.");
                ushort? style = biff2
                    ? (_data[record.Offset + 10] & 1) == 0 ? null : record.Length == 18 ? ReadUInt16(record, 16) : (ushort)(_data[record.Offset + 13] & 63)
                    : (flags & 128) == 0 ? null : (ushort)((flags >> 16) & 0x0FFF);
                if (biff2 && inlineStyles != null && style.HasValue) {
                    if (record.Length == 18) {
                        AddCellLoss("early-biff-row-style", "Formatting", "An extended BIFF2 row style references an absent XF table. The row geometry was retained.", sheet.Name, row, 0);
                        style = null;
                    } else {
                        style = ResolveBiff2InlineStyle(record, 13, sheet.StyleWorkbook!, inlineStyles, formats, sheet.Name, row, 0);
                    }
                }
                if (sheet.Rows.Count >= MaximumRows) throw new InvalidDataException("Too many early BIFF row definitions.");
                bool customHeight = (height & 0x8000) == 0;
                ushort rowHeight = (ushort)(height & 0x7FFF);
                sheet.Rows.Add(new LegacyXlsRowLayout(row + 1, customHeight ? ProjectHeight(rowHeight) : rowHeight / 20d,
                    (flags & 32) != 0, customHeight, style, (byte)(flags & 7), (flags & 16) != 0));
                return true;
            case 0x0024 when biff2:
                RequireLength(record, 4);
                AddColumns(_data[record.Offset], _data[record.Offset + 1], ReadUInt16(record, 2), false, null, 0, false);
                return true;
            case 0x007D when !biff2:
                if (record.Length < 10) throw new InvalidDataException("Truncated early BIFF column layout.");
                ushort columnFlags = ReadUInt16(record, 8);
                AddColumns(ReadUInt16(record, 0), ReadUInt16(record, 2), ReadUInt16(record, 4),
                    (columnFlags & 1) != 0, ReadUInt16(record, 6), (byte)((columnFlags >> 8) & 7), (columnFlags & 0x1000) != 0);
                return true;
            case 0x0025 when biff2:
            case 0x0225 when !biff2:
                RequireLength(record, biff2 ? 2 : 4);
                ushort defaultHeight = ReadUInt16(record, biff2 ? 0 : 2);
                if (biff2) defaultHeight &= 0x7FFF;
                sheet.DefaultRowHeight = ProjectHeight(defaultHeight);
                sheet.DefaultRowsHidden = !biff2 && (ReadUInt16(record, 0) & 2) != 0;
                return true;
            case 0x00AB:
                RequireLength(record, 34);
                if (ReadUInt16(record, 0) != 32) throw new InvalidDataException("Invalid early BIFF global column-width mask.");
                sheet.GlobalColumnWidthMask = new byte[32];
                Array.Copy(_data, record.Offset + 2, sheet.GlobalColumnWidthMask, 0, 32);
                return true;
            case 0x0055:
                RequireLength(record, 2);
                sheet.BaseColumnWidth = ReadUInt16(record, 0);
                return true;
            case 0x0099:
                RequireLength(record, 2);
                sheet.DefaultColumnWidth = ProjectWidth(ReadUInt16(record, 0));
                return true;
            default:
                return false;
        }

        void AddColumns(int first, int last, ushort width, bool hidden, ushort? styleIndex, byte outline, bool collapsed) {
            if (first > last || last > MaximumColumns) throw new InvalidDataException("Invalid early BIFF column range.");
            if (last == MaximumColumns) {
                AddLoss("early-biff-column-tail", "Layout", "A producer column definition extends to the unused 257th column. Its supported range was retained.");
                last--;
            }
            if (first > last) return;
            if (sheet.Columns.Count >= MaximumColumns) throw new InvalidDataException("Too many early BIFF column definitions.");
            sheet.Columns.Add(new LegacyXlsColumnLayout(first + 1, last + 1, ProjectWidth(width), hidden, styleIndex, outline, collapsed));
        }

        double ProjectHeight(ushort twips) {
            if (twips > 8180) {
                AddLoss("early-biff-row-height", "Layout", "One or more early BIFF row heights exceeded 409 points and were bounded to the editable worksheet limit.");
                return 409d;
            }
            return twips / 20d;
        }

        double ProjectWidth(ushort width) {
            if (width == 0) {
                AddLoss("early-biff-zero-column-width", "Layout", "An authored zero column width uses the editable worksheet default width. Its explicit visibility was retained.");
            } else if (width > 65280) {
                AddLoss("early-biff-column-width", "Layout", "One or more early BIFF column widths exceeded 255 character units and were bounded to the editable worksheet limit.");
                return 255d;
            }
            return width / 256d;
        }
    }
}
