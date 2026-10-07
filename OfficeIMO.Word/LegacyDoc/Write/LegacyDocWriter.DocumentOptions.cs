using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private const int Dop97Length = 500;
        private const int Dop2000Length = 544;
        private const int SttbfAssocLength = 42;
        private const int DopBaseEndnotePlacementOffset = 52;
        private const int DopBaseEndnotePlacementShift = 16;
        private const int DopBaseViewFlagsOffset = 82;
        private const int DopBaseCompatibilityFlagsOffset = 8;
        private const int Dop95CompatibilityFlagsOffset = 84;
        private const int Dop2000CompatibilityFlagsOffset = 508;
        private const ushort FacingPagesDopFlag = 0x0001;
        private const uint MirrorMarginsDopFlag = 0x00200000;
        private const ushort GutterAtTopDopFlag = 0x8000;
        private const ushort NoColumnBalanceDopFlag = 0x0020;

        private static byte[] CreateDopBase(LegacyDocWritableBody body) {
            var dop = new byte[body.DopLength];
            // Word 97/2000 reads document note options from DOP rather than section SPRMs.
            byte footnotePosition = GetFootnotePositionOperand(ReadDocumentNoteValue(body.Sections,
                section => section.FootnotePosition, FootnotePositionValues.PageBottom, "footnote placement"))!.Value;
            ushort baseFlags = (ushort)(footnotePosition << 5);
            if (body.FacingPages) baseFlags |= FacingPagesDopFlag;
            WriteUInt16(dop, 0, baseFlags);
            WriteUInt16(dop, 2, CreateDocumentNoteSequence(body.Sections, false));
            WriteUInt16(dop, 10, body.DefaultTabStop);
            WriteUInt16(dop, 492, GetPageNumberFormatOperand(ReadDocumentNoteValue(body.Sections,
                section => section.FootnoteNumberFormat, NumberFormatValues.Decimal, "footnote numbering format"))!.Value);
            WriteUInt16(dop, 494, GetPageNumberFormatOperand(ReadDocumentNoteValue(body.Sections,
                section => section.EndnoteNumberFormat, NumberFormatValues.LowerRoman, "endnote numbering format"))!.Value);

            // MS-DOC DopBase's second flags word contains revision tracking and mirrored margins.
            uint flags = 0;
            if (body.TrackRevisions) flags |= 0x00008000;
            if (body.LockRevisionTracking) flags |= 0x40000000;
            if (body.MirrorMargins) flags |= MirrorMarginsDopFlag;
            if (flags != 0) WriteUInt32(dop, 4, flags);

            if (body.NoColumnBalance) {
                // Dop95's Copts80 repeats Copts60; both copies must agree (MS-DOC).
                WriteUInt16(dop, DopBaseCompatibilityFlagsOffset, NoColumnBalanceDopFlag);
                WriteUInt16(dop, Dop95CompatibilityFlagsOffset, NoColumnBalanceDopFlag);
                if (body.RequiresWord2000Format) WriteUInt16(dop, Dop2000CompatibilityFlagsOffset, NoColumnBalanceDopFlag);
            }

            uint placement = body.EndnotePosition == null ? 3u : (uint)GetEndnotePositionOperand(body.EndnotePosition.Value)!.Value;
            WriteUInt32(dop, DopBaseEndnotePlacementOffset, (uint)CreateDocumentNoteSequence(body.Sections, true) | placement << DopBaseEndnotePlacementShift);
            if (body.GutterAtTop) WriteUInt16(dop, DopBaseViewFlagsOffset, GutterAtTopDopFlag);
            return dop;
        }

        private static ushort CreateDocumentNoteSequence(IReadOnlyList<LegacyDocWritableSection> sections, bool endnotes) {
            int start = ReadDocumentNoteValue(sections, section => endnotes ? section.EndnoteStart : section.FootnoteStart,
                1, endnotes ? "endnote starting number" : "footnote starting number");
            var restart = ReadDocumentNoteValue(sections, section => endnotes ? section.EndnoteRestart : section.FootnoteRestart,
                RestartNumberValues.Continuous, endnotes ? "endnote numbering restart" : "footnote numbering restart");
            return (ushort)((start << 2) | GetNoteRestartOperand(restart)!.Value);
        }

        private static T ReadDocumentNoteValue<T>(IReadOnlyList<LegacyDocWritableSection> sections,
            Func<LegacyDocSectionFormat, T?> selector, T defaultValue, string description) where T : struct {
            T value = sections.Count == 0 ? defaultValue : selector(sections[0].Format) ?? defaultValue;
            foreach (LegacyDocWritableSection section in sections) {
                if (!EqualityComparer<T>.Default.Equals(value, selector(section.Format) ?? defaultValue)) {
                    throw new NotSupportedException($"Native DOC saving supports only one {description} for the whole document. Save as DOCX to retain different section note settings.");
                }
            }
            return value;
        }
    }
}
