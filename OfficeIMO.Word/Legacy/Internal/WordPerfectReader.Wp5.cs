using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace OfficeIMO.Word.Legacy;

internal sealed partial class WordPerfectReader {
    private static readonly int[] Fixed5 = { 4, 9, 11, 3, 3, 5, 6, 7, 4, 5, 6, 6, 8, 10, 10, 12 };

    private void ReadPrefix5(int documentOffset) {
        if (documentOffset == 16) return;
        var visited = new HashSet<int>();
        var ranges = new List<(int Start, int End)>();
        int block = 16;
        while (block != 0) {
            _budget.Record();
            if (!visited.Add(block)) throw new InvalidDataException("The WordPerfect 5 prefix directory contains a cycle.");
            if (block < 16 || block > documentOffset - 50 || U16(_data, block) != 0xfffb ||
                U16(_data, block + 2) != 5 || U16(_data, block + 4) != 50)
                throw new InvalidDataException("The WordPerfect 5 prefix directory is malformed.");
            ranges.Add((block, block + 50));
            for (int i = 1; i < 5; i++) {
                _budget.Record();
                int entry = block + i * 10, type = U16(_data, entry);
                if (type == 0 || type == 0xffff) continue;
                int length = I32(_data, entry + 2), offset = I32(_data, entry + 6);
                if (type > 0x2ff || offset < 16 || offset > documentOffset - length)
                    throw new InvalidDataException("A WordPerfect 5 prefix packet is outside the prefix.");
                if (_packets.ContainsKey(type))
                    throw new InvalidDataException("The WordPerfect 5 prefix contains duplicate packet types.");
                _packets.Add(type, new Packet { Type = type, Offset = offset, Length = length });
                if (length > 0) ranges.Add((offset, offset + length));
            }
            block = I32(_data, block + 6);
        }
        ranges.Sort((left, right) => left.Start.CompareTo(right.Start));
        for (int i = 1; i < ranges.Count; i++)
            if (ranges[i].Start < ranges[i - 1].End) throw new InvalidDataException("WordPerfect 5 prefix packets and indexes overlap.");
        _model.Metadata["PrefixPacketCount"] = _packets.Count.ToString(CultureInfo.InvariantCulture);
        if (_packets.ContainsKey(15) || _packets.ContainsKey(2)) Font5(0, null);
    }

    private void Font5(int index, double? size) {
        if (!_packets.TryGetValue(15, out Packet? fonts) && !_packets.TryGetValue(2, out fonts)) {
            Loss("WORDPERFECT_FONT_DIRECTORY", "Formatting", "A font change has no qualified font directory.");
            return;
        }
        if (fonts.Length % 86 != 0 || index >= fonts.Length / 86) throw new InvalidDataException("The WordPerfect 5 font index is outside its directory.");
        int entry = fonts.Offset + index * 86;
        FontSize(size ?? U16(_data, entry + (fonts.Type == 15 ? 47 : 22)) / 50d);
        if (!_packets.TryGetValue(7, out Packet? names)) { _font = null; return; }
        int offset = U16(_data, entry + 18);
        if (offset >= names.Length) throw new InvalidDataException("The WordPerfect 5 font name is outside its string pool.");
        int start = names.Offset + offset, end = start;
        while (end < names.Offset + names.Length && _data[end] != 0) { _budget.Record(); end++; }
        if (end == names.Offset + names.Length) throw new InvalidDataException("The WordPerfect 5 font name is unterminated.");
        _budget.Text(end - start);
        _font = Encoding.ASCII.GetString(_data, start, end - start);
    }

    private int ReadCode5(int at, int end) {
        int code = _data[at];
        if (code >= 0xc0 && code <= 0xcf) {
            int length = Fixed5[code - 0xc0];
            if (length > end - at || _data[at + length - 1] != code)
                throw new InvalidDataException("A WordPerfect 5 fixed function is truncated or has mismatched gates.");
            switch (code) {
                case 0xc0: Append(Character(_data[at + 1], _data[at + 2])); break;
                case 0xc1: Append("\t"); break;
                case 0xc3: case 0xc4: Attribute(_data[at + 1], code == 0xc3); break;
                default: Loss("WORDPERFECT_FIXED_" + code.ToString("X2"), "Formatting", "A fixed WordPerfect 5 function was omitted."); break;
            }
            return at + length;
        }
        if (code >= 0xd0) {
            if (end - at < 8) throw new InvalidDataException("A WordPerfect 5 variable function is truncated.");
            int storedSize = U16(_data, at + 2), length = storedSize + 4, sub = _data[at + 1];
            if (storedSize < 4 || length > end - at || U16(_data, at + length - 4) != storedSize ||
                _data[at + length - 2] != sub || _data[at + length - 1] != code)
                throw new InvalidDataException("A WordPerfect 5 variable function has an invalid size or mismatched gates.");
            Function5(code, sub, at + 4, length - 8);
            return at + length;
        }
        switch (code) {
            case 0: break;
            case 0x0a: case 0x8c: case 0x90: case 0x99: FinishParagraph(true); break;
            case 0x0c: BreakPage(); break;
            case 0x0b: case 0x0d: case 0x93: case 0x94: case 0x95: Append(" "); break;
            case 0xa0: Append("\u00a0"); break;
            case 0xa9: case 0xaa: case 0xab: Append("-"); break;
            case 0xac: case 0xad: case 0xae: Append("\u00ad"); break;
            default: Loss("WORDPERFECT_SINGLE_" + code.ToString("X2"), "Structure", "A single-byte WordPerfect 5 function was omitted."); break;
        }
        return at + 1;
    }

    private void Function5(int code, int sub, int data, int length) {
        void Need(int minimum) { if (length < minimum) throw new InvalidDataException("A WordPerfect 5 function's data is truncated."); }
        if (code == 0xd0 && (sub == 1 || sub == 5)) {
            Need(8); if (!ChangeSection()) return;
            if (sub == 1) { _section.LeftPoints = Points(U16(_data, data + 4)); _section.RightPoints = Points(U16(_data, data + 6)); }
            else { _section.TopPoints = Points(U16(_data, data + 4)); _section.BottomPoints = Points(U16(_data, data + 6)); }
        } else if (code == 0xd0 && sub == 6) { Need(2); Justification(_data[data + 1]); }
        else if (code == 0xd0 && sub == 0x0b) {
            Need(190); if (!ChangeSection()) return;
            double height = Points(U16(_data, data + 95)), width = Points(U16(_data, data + 97));
            if (_data[data + 189] == 1) { double swap = height; height = width; width = swap; }
            else if (_data[data + 189] != 0) throw new InvalidDataException("The WordPerfect 5 page orientation is unsupported.");
            if (height <= 0 || width <= 0) throw new InvalidDataException("WordPerfect 5 page dimensions must be positive.");
            _section.WidthPoints = width; _section.HeightPoints = height;
        } else if (code == 0xd1 && sub == 0) {
            Need(6); _color = _data[data + 3].ToString("X2") + _data[data + 4].ToString("X2") + _data[data + 5].ToString("X2");
        } else if (code == 0xd1 && sub == 1) {
            Need(26); Font5(_data[data + 25], length >= 30 ? U16(_data, data + 28) / 50d : null);
        } else if (code == 0xd5) {
            Need(18); int flags = _data[data + 7];
            int occurrence = (flags & 1) != 0 ? 3 : (flags & 2) != 0 ? 1 : (flags & 4) != 0 ? 2 : 0;
            SetStory(sub, occurrence, ReadStory(data + 18, data + length));
        } else if (code == 0xd6 && sub <= 1) {
            Need(4); int start = sub == 0 ? 15 + _data[data + 3] * 2 : 7;
            Need(start); AddNote(sub == 0 ? LegacyWordNoteKind.Footnote : LegacyWordNoteKind.Endnote, ReadStory(data + start, data + length));
            if ((_data[data] & 0x80) != 0) Loss("WORDPERFECT_NOTE_SYMBOL", "Notes", "A custom source note symbol was replaced by DOCX note numbering.");
        } else if (code == 0xd2 && sub == 0x0b) {
            Need(4); int newValues = 24 + U16(_data, data + 2) * 5;
            Need(newValues + 24); int columns = U16(_data, data + newValues + 2);
            if (columns < 1 || columns > 32 || columns * 5 > length - newValues - 24)
                throw new InvalidDataException("The WordPerfect 5 table column directory is malformed.");
            if (_table != null) EndTable();
            BeginTable();
            for (int i = 0; i < columns; i++) {
                _budget.Item(); double width = Points(U16(_data, data + newValues + 24 + i * 2));
                if (width <= 0) throw new InvalidDataException("WordPerfect 5 table widths must be positive.");
                _table!.ColumnWidthsPoints.Add(width);
            }
        } else if (code == 0xdc && sub == 0) {
            Need(11);
            if ((_data[data + 2] & 0x7f) != 1 || _data[data + 3] != 1 || (_data[data + 2] & 0x80) != 0)
                throw new InvalidDataException("WordPerfect 5 merged cells require a separately qualified profile.");
            BeginCell(); Justification(_data[data + 10]);
        } else if (code == 0xdc && sub == 1 || code == 0xdd && (sub == 1 || sub == 3)) BeginRow(false);
        else if ((code == 0xdc || code == 0xdd) && sub == 2) EndTable();
        else if (code == 0xd9) Inert("WORDPERFECT_MERGE_INERT", OfficeLegacyInertContentKind.EmbeddedCode, "WordPerfect 5 merge functions were not executed.");
        else Loss("WORDPERFECT_FUNCTION_" + code.ToString("X2") + sub.ToString("X2"), "Structure", "A WordPerfect 5 function is outside the recovered profile.");
    }
}
