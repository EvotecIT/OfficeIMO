using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeIMO.Word.Legacy;

internal sealed partial class WordPerfectReader {
    private static readonly int[] Fixed6 = { 4, 5, 3, 3, 3, 3, 4, 4, 4, 5, 5, 6, 6, 8, 8 };
    private const string International6 = "åÅæÆäÄáàâãÃçÇëéÉèêíñÑøØõÕöÖüÜúùß";
    private readonly Stack<bool> _invalidUndo = new();
    private bool Suppressed => _encasedNote >= 0 || (_invalidUndo.Count != 0 && _invalidUndo.Peek());

    private void ReadPrefix6(int documentOffset) {
        int index = Math.Max(16, U16(_data, 14));
        CheckRange(_data, index, 14);
        int count = U16(_data, index + 2);
        if (count < 1 || index > documentOffset - checked(count * 14))
            throw new InvalidDataException("The WordPerfect 6 prefix index is outside the prefix.");
        var ranges = new List<(int Start, int End)>();
        for (int id = 1; id < count; id++) {
            _budget.Record();
            int entry = index + id * 14;
            int length = I32(_data, entry + 6), offset = I32(_data, entry + 10);
            if (length == 0) continue;
            if (offset < index + count * 14 || offset > documentOffset - length)
                throw new InvalidDataException("A WordPerfect prefix packet overlaps the index or document area.");
            ranges.Add((offset, offset + length));
            var packet = new Packet { Type = _data[entry + 1], Offset = offset, Length = length };
            if ((_data[entry] & 1) != 0) {
                int children = U16(_data, offset), headerLength = checked(2 + children * 2);
                if (headerLength > length) throw new InvalidDataException("A WordPerfect prefix child directory is truncated.");
                packet.Children = new int[children];
                for (int child = 0; child < children; child++) {
                    _budget.Record();
                    packet.Children[child] = U16(_data, offset + 2 + child * 2);
                    if (packet.Children[child] >= count) throw new InvalidDataException("A WordPerfect child packet ID is outside the directory.");
                }
                packet.Offset += headerLength; packet.Length -= headerLength;
            }
            _packets.Add(id, packet);
            if (packet.Type == 0x70 || packet.Type == 0x71)
                Inert("WORDPERFECT_OLE_INERT", OfficeLegacyInertContentKind.EmbeddedObjects, "WordPerfect OLE packets were kept inert and were not activated.");
        }
        ranges.Sort((left, right) => left.Start.CompareTo(right.Start));
        for (int i = 1; i < ranges.Count; i++)
            if (ranges[i].Start < ranges[i - 1].End) throw new InvalidDataException("WordPerfect prefix packets overlap.");
        _model.Metadata["PrefixPacketCount"] = _packets.Count.ToString(CultureInfo.InvariantCulture);
        foreach (Packet packet in _packets.Values.Where(packet => packet.Type == 0x25)) {
            if (packet.Children.Length != 1 || packet.Length < 2) throw new InvalidDataException("The WordPerfect initial-font packet is malformed.");
            _font = Font6(packet.Children[0]);
            FontSize(U16(_data, packet.Offset) / 50d);
        }
    }

    private Packet Packet6(int id, int expectedType) {
        if (!_packets.TryGetValue(id, out Packet? packet) || packet.Type != expectedType)
            throw new InvalidDataException("A WordPerfect function references a missing or incorrectly typed prefix packet.");
        return packet;
    }

    private string? Font6(int id) {
        if (id == 0) return null;
        Packet packet = Packet6(id, 0x55);
        if (packet.Length < 24) throw new InvalidDataException("A WordPerfect font descriptor is truncated.");
        int length = U16(_data, packet.Offset + 22);
        if (length > packet.Length - 24 || (length & 1) != 0) throw new InvalidDataException("The WordPerfect font name is malformed.");
        return WordString(packet.Offset + 24, length);
    }

    private string WordString(int start, int length) {
        CheckRange(_data, start, length);
        if ((length & 1) != 0) throw new InvalidDataException("A WordPerfect character string has an odd byte count.");
        var text = new StringBuilder();
        for (int at = start; at < start + length; at += 2) {
            _budget.Record();
            int value = U16(_data, at);
            if (value == 0) break;
            string character = Character(value & 255, value >> 8);
            _budget.Text(character.Length);
            text.Append(character);
        }
        return text.ToString();
    }

    private string Character(int value, int characterSet) {
        if (characterSet == 0 && value >= 32 && value <= 126) return ((char)value).ToString();
        // WordPerfect's character sets are not Windows code pages. Unmapped glyphs are explicit losses.
        if (characterSet == 4) {
            string? mapped = value switch {
                0 => "●", 1 => "○", 2 => "■", 3 => "•", 4 => "*", 5 => "¶", 6 => "§",
                7 => "¡", 8 => "¿", 9 => "«", 10 => "»", 11 => "£", 12 => "¥", 13 => "₧", 14 => "ƒ",
                15 => "ª", 16 => "º", 17 => "½", 18 => "¼", 19 => "¢", 20 => "²", 21 => "ⁿ",
                22 => "®", 23 => "©", 24 => "¤", 25 => "¾", 26 => "³", 27 => "‛", 28 => "’", 29 => "‘",
                30 => "‟", 31 => "”", 32 => "“", 33 => "–", 34 => "—", 35 => "‹", 36 => "›",
                37 => "○", 38 => "□", 39 => "†", 40 => "‡", 41 => "™", 42 => "℠", 43 => "℞",
                44 => "●", 45 => "◦", 46 => "■", 47 => "▪", 48 => "□", 49 => "▫", 50 => "‒",
                51 => "ﬀ", 52 => "ﬃ", 53 => "ﬄ", 54 => "ﬁ", 55 => "ﬂ", 56 => "…", 57 => "$",
                58 => "₣", 59 => "₢", 60 => "₠", 61 => "₤", 62 => "‚", 63 => "„", 64 => "⅓",
                65 => "⅔", 66 => "⅛", 67 => "⅜", 68 => "⅝", 69 => "⅞", 70 => "Ⓜ", 71 => "Ⓟ", 72 => "€",
                _ => null
            };
            if (mapped != null) return mapped;
        }
        Loss("WORDPERFECT_CHARACTER_SET_" + characterSet, "Text",
            "An unmapped WordPerfect character in set " + characterSet + " was replaced by U+FFFD; it was not interpreted as an unrelated code page.");
        return "\ufffd";
    }

    private List<LegacyWordParagraph> TextPacket6(int id) {
        Packet packet = Packet6(id, 8);
        if (!_activePackets.Add(id)) throw new InvalidDataException("WordPerfect text packets contain a recursive reference.");
        try {
            if (packet.Length < 6) throw new InvalidDataException("A WordPerfect text packet is truncated.");
            int blocks = U16(_data, packet.Offset);
            int textOffset = I32(_data, packet.Offset + 2);
            if (textOffset < 6 + blocks * 4 || textOffset > packet.Length)
                throw new InvalidDataException("The WordPerfect text-block directory is malformed.");
            int total = 0;
            for (int i = 0; i < blocks; i++) {
                _budget.Record();
                int length = I32(_data, packet.Offset + 6 + i * 4);
                if (length > packet.Length - textOffset - total) throw new InvalidDataException("WordPerfect text blocks exceed their packet.");
                total += length;
            }
            return ReadStory(packet.Offset + textOffset, packet.Offset + textOffset + total);
        } finally { _activePackets.Remove(id); }
    }

    private int ReadCode6(int at, int end) {
        byte code = _data[at];
        if (code >= 0xf0 && code <= 0xfe) {
            int length = Fixed6[code - 0xf0];
            if (length > end - at || _data[at + length - 1] != code) throw new InvalidDataException("A WordPerfect fixed function is truncated or has mismatched gates.");
            if (code == 0xf1) {
                int type = _data[at + 1];
                // Undo levels identify edit operations, not paired range IDs. Producers may use
                // different levels on the two gates; visibility is determined by the range type.
                if (type == 0 || type == 2) {
                    _budget.Item(); _invalidUndo.Push(type == 0);
                    if (type == 0) Loss("WORDPERFECT_DELETED_TEXT", "Text", "Deleted undo text was excluded from recovery.");
                } else if (type == 1 || type == 3) {
                    if (_invalidUndo.Count == 0 || _invalidUndo.Pop() != (type == 1)) throw new InvalidDataException("WordPerfect undo ranges are unbalanced.");
                } else throw new InvalidDataException("The WordPerfect undo function is invalid.");
            } else if (!Suppressed) {
                if (code == 0xf2 || code == 0xf3) Attribute(_data[at + 1], code == 0xf2);
                else if (code == 0xf0) Append(Character(_data[at + 1], _data[at + 2]));
                else Loss("WORDPERFECT_FIXED_" + code.ToString("X2"), "Formatting", "Fixed function 0x" + code.ToString("X2") + " was omitted.");
            }
            return at + length;
        }
        if (code >= 0xd0 && code <= 0xef) {
            if (end - at < 10) throw new InvalidDataException("A WordPerfect variable function is truncated.");
            int length = U16(_data, at + 2), sub = _data[at + 1];
            if (length < 10 || length > end - at || U16(_data, at + length - 3) != length || _data[at + length - 1] != code)
                throw new InvalidDataException("A WordPerfect variable function has an invalid size or mismatched gates.");
            int cursor = at + 5;
            int[] ids = Array.Empty<int>();
            if ((_data[at + 4] & 0x80) != 0) {
                int count = _data[cursor++];
                if (count * 2 > at + length - 5 - cursor) throw new InvalidDataException("A WordPerfect function's prefix references are truncated.");
                ids = new int[count];
                for (int i = 0; i < count; i++) ids[i] = U16(_data, cursor + i * 2);
                cursor += count * 2;
            }
            if (cursor > at + length - 5) throw new InvalidDataException("A WordPerfect function header is truncated.");
            int dataLength = U16(_data, cursor); cursor += 2;
            if (dataLength > at + length - 3 - cursor) throw new InvalidDataException("A WordPerfect function's data exceeds its envelope.");
            if (_encasedNote >= 0 && code == 0xd7 && sub == _encasedNote) _encasedNote = -1;
            else if (!Suppressed) Function6(code, sub, cursor, dataLength, ids);
            return at + length;
        }
        if (Suppressed) return at + 1;
        if (code == 0) return at + 1;
        if (code >= 1 && code <= 32) { Append(International6[code - 1].ToString()); return at + 1; }
        switch (code) {
            case 0x80: case 0xcd: case 0xce: case 0xcf: Append(" "); break;
            case 0x81: Append("\u00a0"); break;
            case 0x82: case 0x83: Append("\u00ad"); break;
            case 0x84: Append("-"); break;
            case 0x85: break;
            case 0x87: case 0xb7: case 0xb8: case 0xb9: case 0xca: case 0xcb: case 0xcc: FinishParagraph(true); break;
            case 0xb4: case 0xc7: BreakPage(); break;
            case 0xbd: case 0xbe: case 0xbf: EndTable(); break;
            case 0xc0: case 0xc1: case 0xc2: case 0xc3: case 0xc4: case 0xc5: BeginRow(); break;
            case 0xc6: BeginCell(); break;
            default: Loss("WORDPERFECT_SINGLE_" + code.ToString("X2"), "Structure", "Single-byte function 0x" + code.ToString("X2") + " was omitted."); break;
        }
        return at + 1;
    }

    private void Function6(int code, int sub, int data, int length, int[] ids) {
        void Need(int minimum) { if (length < minimum) throw new InvalidDataException("A WordPerfect function's non-deletable data is truncated."); }
        if (code == 0xd1 && (sub == 0 || sub == 1)) {
            Need(2); if (!ChangeSection()) return;
            if (sub == 0) _section.TopPoints = Points(U16(_data, data)); else _section.BottomPoints = Points(U16(_data, data));
        } else if (code == 0xd1 && sub == 0x11) {
            Need(9); if (!ChangeSection()) return;
            double height = Points(U16(_data, data + 3)), width = Points(U16(_data, data + 5));
            if (width <= 0 || height <= 0) throw new InvalidDataException("WordPerfect page dimensions must be positive.");
            if (_data[data + 8] == 1) { double swap = width; width = height; height = swap; }
            else if (_data[data + 8] != 0) throw new InvalidDataException("WordPerfect page orientation is unsupported.");
            _section.WidthPoints = width; _section.HeightPoints = height;
        } else if (code == 0xd2 && sub <= 1) {
            Need(2); if (!ChangeSection()) return;
            if (sub == 0) _section.LeftPoints = Points(U16(_data, data)); else _section.RightPoints = Points(U16(_data, data));
        } else if (code == 0xd3 && sub == 5) { Need(1); Justification(_data[data]); }
        else if (code == 0xd3 && sub == 0x0a) {
            Need(4);
            if (length >= 6) {
                _spacingAfter = Points(U16(_data, data + 4));
                if (_paragraph != null) _paragraph.SpacingAfterPoints = _spacingAfter;
            } else Loss("WORDPERFECT_RELATIVE_SPACING", "Formatting", "Relative paragraph spacing without an absolute WPU value was not projected as points.");
        } else if (code == 0xd4 && sub == 0x18) {
            Need(3); _color = _data[data].ToString("X2") + _data[data + 1].ToString("X2") + _data[data + 2].ToString("X2");
        } else if (code == 0xd4 && (sub == 0x1a || sub == 0x1b)) {
            Need(8); if (ids.Length > 0 && ids[0] != 0) _font = Font6(ids[0]);
            FontSize(U16(_data, data + (sub == 0x1a ? 6 : 0)) / 50d);
        } else if (code == 0xd4 && sub == 0x2a) BeginTable();
        else if (code == 0xd4 && sub == 0x2b) { /* Ends the table definition, not the table. */ }
        else if (code == 0xd4 && sub == 0x2c) {
            Need(17); if (_table == null) throw new InvalidDataException("WordPerfect column definition appears outside a table.");
            _budget.Item(); double width = Points(U16(_data, data + 1));
            if (width <= 0) throw new InvalidDataException("WordPerfect table width must be positive.");
            _table.ColumnWidthsPoints.Add(width);
        } else if (code == 0xd6) {
            Need(1); if (ids.Length == 0) throw new InvalidDataException("A WordPerfect running story has no text packet.");
            SetStory(sub, _data[data] & 3, TextPacket6(ids[0]));
        } else if (code == 0xd7 && (sub == 0 || sub == 2)) {
            if (ids.Length == 0) throw new InvalidDataException("A WordPerfect note has no text packet.");
            AddNote(sub == 0 ? LegacyWordNoteKind.Footnote : LegacyWordNoteKind.Endnote, TextPacket6(ids[0]));
            _encasedNote = sub + 1;
        } else if (code == 0xd7 && (sub == 1 || sub == 3)) throw new InvalidDataException("WordPerfect note references are unbalanced.");
        else if (code == 0xe0) Append("\t");
        else if (code == 0xd0) Eol6(sub, data, length);
        else if (code == 0xde) Inert("WORDPERFECT_MERGE_INERT", OfficeLegacyInertContentKind.EmbeddedCode, "Merge commands were retained as inert source evidence and were not executed.");
        else if (code == 0xd4 && sub == 0x34) {
            Need(1); Inert("WORDPERFECT_HYPERTEXT_INERT", (_data[data] == 2 ? OfficeLegacyInertContentKind.Macros : OfficeLegacyInertContentKind.ExternalLinks), "Hypertext actions were kept inert and were not followed or executed.");
        } else if (code == 0xdf && sub <= 2) Graphic6(ids);
        else if (code == 0xdd) {
            Loss("WORDPERFECT_STYLE_BEHAVIOR", "Style", "Style packets, inheritance and automatic next-style behavior are not reconstructed; only separately decoded inline formatting is applied.");
        } else {
            Loss("WORDPERFECT_FUNCTION_" + code.ToString("X2") + sub.ToString("X2"), "Structure",
                "Function 0x" + code.ToString("X2") + sub.ToString("X2") + " is outside the recovered profile.");
        }
    }

    private void Eol6(int sub, int data, int length) {
        if (sub == 1 || sub == 2 || sub == 3 || sub == 0x14 || sub == 0x15 || sub == 0x16) { Append(" "); return; }
        if (sub >= 4 && sub <= 6 || sub >= 0x17 && sub <= 0x19) { FinishParagraph(true); return; }
        if (sub == 9 || sub == 0x1c) { BreakPage(); return; }
        if (sub >= 0x11 && sub <= 0x13) { EndTable(); return; }
        if (sub == 0x0a) BeginCell();
        else if (sub >= 0x0b && sub <= 0x10) BeginRow();
        else { Loss("WORDPERFECT_EOL_" + sub.ToString("X2"), "Layout", "A column or unsupported end-of-line function was approximated."); FinishParagraph(true); return; }
        // EOL non-deletable data begins with the byte count of internal formatter data.
        if (length == 0) return;
        if (length < 2) throw new InvalidDataException("WordPerfect cell information is truncated.");
        int cursor = data + 2 + U16(_data, data), end = data + length;
        if (cursor > end) throw new InvalidDataException("WordPerfect cell formatter data exceeds its function.");
        while (cursor < end) {
            _budget.Record(); int type = _data[cursor++], size;
            switch (type) {
                case 0x80: size = 4; if (end - cursor < size) throw new InvalidDataException("WordPerfect row information is truncated."); if (_row != null) _row.IsHeader = (_data[cursor] & 4) != 0; break;
                case 0x85:
                    size = 3;
                    if (end - cursor < size) throw new InvalidDataException("WordPerfect spanning information is truncated.");
                    if (_data[cursor] != 1 || _data[cursor + 1] != 1)
                        throw new InvalidDataException("WordPerfect merged cells require a separately qualified profile.");
                    break;
                default:
                    Loss("WORDPERFECT_CELL_ATTRIBUTES", "Table", "Advanced cell attributes were not projected.");
                    return;
            }
            cursor += size;
        }
    }
}
