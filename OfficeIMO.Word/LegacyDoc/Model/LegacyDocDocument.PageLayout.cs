using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.LegacyDoc.Model {
    public sealed partial class LegacyDocDocument {
        private static IReadOnlyList<LegacyDocSection> ApplyDopNoteSettings(
            IReadOnlyList<LegacyDocSection> sections, byte[] tableStream, LegacyDocFib fib) {
            // Later formats use section note SPRMs. Keep tolerating the short DOPs
            // emitted by older writers without overriding their section metadata.
            if (fib.NFib > 0x00D9 || fib.LcbDop < 500 || fib.FcDop < 0 || fib.FcDop > tableStream.Length - fib.LcbDop) return sections;
            int offset = fib.FcDop;
            int position = (LegacyDocFib.ReadUInt16(tableStream, offset) >> 5) & 3;
            ushort footnoteSequence = LegacyDocFib.ReadUInt16(tableStream, offset + 2);
            uint endnoteSequence = unchecked((uint)LegacyDocFib.ReadInt32(tableStream, offset + 52));
            ushort footnoteFormat = LegacyDocFib.ReadUInt16(tableStream, offset + 492);
            ushort endnoteFormat = LegacyDocFib.ReadUInt16(tableStream, offset + 494);
            return sections.Select(section => new LegacyDocSection(section.StartCharacter, section.EndCharacter,
                section.Format.WithNoteSettings(
                    footnotePosition: position == 1 ? FootnotePositionValues.PageBottom : position == 2 ? FootnotePositionValues.BeneathText : (FootnotePositionValues?)null,
                    footnoteRestart: ReadDopNoteRestart(footnoteSequence & 3, true),
                    footnoteStart: footnoteSequence >> 2 == 0 ? (int?)null : footnoteSequence >> 2,
                    footnoteNumberFormat: footnoteFormat <= byte.MaxValue ? LegacyDocNumberFormatMapper.FromNfc((byte)footnoteFormat) : null,
                    endnoteRestart: ReadDopNoteRestart((int)(endnoteSequence & 3), false),
                    endnoteStart: (endnoteSequence >> 2 & 0x3FFF) == 0 ? (int?)null : (int)(endnoteSequence >> 2 & 0x3FFF),
                    endnoteNumberFormat: endnoteFormat <= byte.MaxValue ? LegacyDocNumberFormatMapper.FromNfc((byte)endnoteFormat) : null))).ToArray();
        }

        private static RestartNumberValues? ReadDopNoteRestart(int value, bool footnotes) =>
            value == 0 ? RestartNumberValues.Continuous : value == 1 ? RestartNumberValues.EachSection :
            value == 2 && footnotes ? RestartNumberValues.EachPage : (RestartNumberValues?)null;

        private void ReadDopMarginSettings(byte[] tableStream, LegacyDocFib fib) {
            if (fib.FcDop < 0 || fib.FcDop > tableStream.Length - fib.LcbDop) return;

            const uint mirrorMarginsFlag = 0x00200000;
            if (fib.LcbDop >= 8) {
                uint flags = unchecked((uint)LegacyDocFib.ReadInt32(tableStream, fib.FcDop + 4));
                MirrorMargins = (flags & mirrorMarginsFlag) != 0;
            }
            // MS-DOC DopBase stores Copts60 at byte 8; bit 5 is fNoColumnBalance.
            const int compatibilityFlagsOffset = 8;
            const ushort noColumnBalanceFlag = 0x0020;
            if (fib.LcbDop >= compatibilityFlagsOffset + sizeof(ushort)) {
                ushort flags = LegacyDocFib.ReadUInt16(tableStream, fib.FcDop + compatibilityFlagsOffset);
                NoColumnBalance = (flags & noColumnBalanceFlag) != 0;
            }
            if (fib.LcbDop >= 12) {
                ushort interval = LegacyDocFib.ReadUInt16(tableStream, fib.FcDop + 10);
                if (interval <= short.MaxValue) DefaultTabStop = interval;
            }
            const int viewFlagsOffset = 82;
            const ushort gutterAtTopFlag = 0x8000;
            if (fib.LcbDop >= viewFlagsOffset + sizeof(ushort)) {
                ushort flags = LegacyDocFib.ReadUInt16(tableStream, fib.FcDop + viewFlagsOffset);
                GutterAtTop = (flags & gutterAtTopFlag) != 0;
            }
        }
    }
}
