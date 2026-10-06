namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private const int DopBaseLength = 8;
        private const int DopBaseEndnotePlacementLength = 56;
        private const int DopBaseEndnotePlacementOffset = 52;
        private const int DopBaseEndnotePlacementShift = 16;
        private const int DopBaseFullLength = 84;
        private const int DopBaseViewFlagsOffset = 82;
        private const ushort FacingPagesDopFlag = 0x0001;
        private const uint MirrorMarginsDopFlag = 0x00200000;
        private const ushort GutterAtTopDopFlag = 0x8000;

        private static byte[] CreateDopBase(LegacyDocWritableBody body) {
            var dop = new byte[body.DopLength];
            if (body.FacingPages) WriteUInt16(dop, 0, FacingPagesDopFlag);

            // MS-DOC DopBase's second flags word contains revision tracking and mirrored margins.
            uint flags = 0;
            if (body.TrackRevisions) flags |= 0x00008000;
            if (body.LockRevisionTracking) flags |= 0x40000000;
            if (body.MirrorMargins) flags |= MirrorMarginsDopFlag;
            if (flags != 0) WriteUInt32(dop, 4, flags);

            if (body.EndnotePosition != null) {
                uint placement = (uint)GetEndnotePositionOperand(body.EndnotePosition.Value)!.Value;
                WriteUInt32(dop, DopBaseEndnotePlacementOffset, placement << DopBaseEndnotePlacementShift);
            }
            if (body.GutterAtTop) WriteUInt16(dop, DopBaseViewFlagsOffset, GutterAtTopDopFlag);
            return dop;
        }
    }
}
