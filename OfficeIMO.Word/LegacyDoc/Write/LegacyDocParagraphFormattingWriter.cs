namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static class LegacyDocParagraphFormattingWriter {
        private const int PapxFkpBxLength = 13;
        private const ushort SprmPFKeep = 0x2405;
        private const ushort SprmPFKeepFollow = 0x2406;
        private const ushort SprmPFPageBreakBefore = 0x2407;
        private const ushort SprmPJc = 0x2461;
        private const ushort SprmPDxaRight = 0x840E;
        private const ushort SprmPDxaLeft = 0x840F;
        private const ushort SprmPDxaLeft1 = 0x8411;
        private const ushort SprmPDyaLine = 0x6412;
        private const ushort SprmPDyaBefore = 0xA413;
        private const ushort SprmPDyaAfter = 0xA414;
        private const ushort SprmPFWidowControl = 0x2431;

        internal static void WritePapxFkp(byte[] stream, int pageOffset, int textOffset, int oleSectorSize, IReadOnlyList<LegacyDocWritableParagraphSegment> segments, int bytesPerCharacter) {
            if (segments.Count == 0 || segments.Count > byte.MaxValue) {
                throw new NotSupportedException("Native DOC saving currently supports paragraph formatting only when it fits in one paragraph-format page.");
            }

            int rgbxOffset = pageOffset + ((segments.Count + 1) * 4);
            int papxOffset = oleSectorSize - 1;

            if (rgbxOffset + (segments.Count * PapxFkpBxLength) > pageOffset + papxOffset) {
                throw new NotSupportedException("Native DOC saving currently supports paragraph formatting only when it fits in one paragraph-format page.");
            }

            for (int index = 0; index < segments.Count; index++) {
                LegacyDocWritableParagraphSegment segment = segments[index];
                WriteInt32(stream, pageOffset + (index * 4), textOffset + (segment.StartCharacter * bytesPerCharacter));
                if (!segment.Formatting.HasFormatting) {
                    continue;
                }

                byte[] papx = CreatePapx(segment.Formatting);
                papxOffset -= papx.Length;
                papxOffset = papxOffset % 2 == 0 ? papxOffset : papxOffset - 1;
                if (pageOffset + papxOffset <= (rgbxOffset + (segments.Count * PapxFkpBxLength)) || papxOffset / 2 > byte.MaxValue) {
                    throw new NotSupportedException("Native DOC saving currently supports paragraph formatting only when it fits in one paragraph-format page.");
                }

                Buffer.BlockCopy(papx, 0, stream, pageOffset + papxOffset, papx.Length);
                stream[rgbxOffset + (index * PapxFkpBxLength)] = (byte)(papxOffset / 2);
            }

            LegacyDocWritableParagraphSegment lastSegment = segments[segments.Count - 1];
            WriteInt32(stream, pageOffset + (segments.Count * 4), textOffset + (lastSegment.EndCharacter * bytesPerCharacter));
            stream[pageOffset + oleSectorSize - 1] = (byte)segments.Count;
        }

        private static byte[] CreatePapx(LegacyDocWritableParagraphFormatting formatting) {
            var grpprl = new List<byte>(6) {
                (byte)((formatting.StyleIndex ?? 0) & 0xFF),
                (byte)((formatting.StyleIndex ?? 0) >> 8)
            };

            if (formatting.Alignment != null) {
                AddSingleByteSprm(grpprl, SprmPJc, formatting.Alignment.Value);
            }

            if (formatting.KeepLinesTogether == true) {
                AddSingleByteSprm(grpprl, SprmPFKeep, 1);
            }

            if (formatting.KeepWithNext == true) {
                AddSingleByteSprm(grpprl, SprmPFKeepFollow, 1);
            }

            if (formatting.PageBreakBefore == true) {
                AddSingleByteSprm(grpprl, SprmPFPageBreakBefore, 1);
            }

            if (formatting.AvoidWidowAndOrphan == true) {
                AddSingleByteSprm(grpprl, SprmPFWidowControl, 1);
            }

            if (formatting.LeftIndentTwips != null) {
                AddInt16Sprm(grpprl, SprmPDxaLeft, formatting.LeftIndentTwips.Value);
            }

            if (formatting.RightIndentTwips != null) {
                AddInt16Sprm(grpprl, SprmPDxaRight, formatting.RightIndentTwips.Value);
            }

            if (formatting.FirstLineIndentTwips != null) {
                AddInt16Sprm(grpprl, SprmPDxaLeft1, formatting.FirstLineIndentTwips.Value);
            }

            if (formatting.SpacingBeforeTwips != null) {
                AddInt16Sprm(grpprl, SprmPDyaBefore, formatting.SpacingBeforeTwips.Value);
            }

            if (formatting.SpacingAfterTwips != null) {
                AddInt16Sprm(grpprl, SprmPDyaAfter, formatting.SpacingAfterTwips.Value);
            }

            if (formatting.LineSpacingTwips != null) {
                AddLineSpacingSprm(grpprl, formatting.LineSpacingTwips.Value);
            }

            if (grpprl.Count % 2 != 0) {
                grpprl.Add(0);
            }

            int cb = grpprl.Count / 2;
            if (cb > byte.MaxValue) {
                throw new NotSupportedException("Native DOC saving cannot write paragraph formatting because the PAPX record is too large.");
            }

            var papx = new byte[grpprl.Count + 2];
            papx[0] = 0;
            papx[1] = (byte)cb;
            grpprl.CopyTo(papx, 2);
            return papx;
        }

        private static void AddSingleByteSprm(List<byte> grpprl, ushort sprm, byte operand) {
            grpprl.Add((byte)(sprm & 0xFF));
            grpprl.Add((byte)(sprm >> 8));
            grpprl.Add(operand);
        }

        private static void AddInt16Sprm(List<byte> grpprl, ushort sprm, int operand) {
            if (operand < short.MinValue || operand > short.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports paragraph spacing and indentation only within the Word 97-2003 signed twip range.");
            }

            short value = checked((short)operand);
            grpprl.Add((byte)(sprm & 0xFF));
            grpprl.Add((byte)(sprm >> 8));
            grpprl.Add((byte)(value & 0xFF));
            grpprl.Add((byte)(value >> 8));
        }

        private static void AddLineSpacingSprm(List<byte> grpprl, int lineSpacingTwips) {
            AddInt16Sprm(grpprl, SprmPDyaLine, lineSpacingTwips);
            grpprl.Add(0);
            grpprl.Add(0);
        }

        private static void WriteInt32(byte[] bytes, int offset, int value) {
            bytes[offset] = (byte)value;
            bytes[offset + 1] = (byte)(value >> 8);
            bytes[offset + 2] = (byte)(value >> 16);
            bytes[offset + 3] = (byte)(value >> 24);
        }
    }

    internal readonly struct LegacyDocWritableParagraphFormatting : IEquatable<LegacyDocWritableParagraphFormatting> {
        internal static readonly LegacyDocWritableParagraphFormatting Plain = new LegacyDocWritableParagraphFormatting(null, null, null, null, null, null, null, null, null, null, null, null);

        internal LegacyDocWritableParagraphFormatting(
            byte? alignment,
            ushort? styleIndex,
            int? spacingBeforeTwips,
            int? spacingAfterTwips,
            int? lineSpacingTwips,
            int? leftIndentTwips,
            int? rightIndentTwips,
            int? firstLineIndentTwips,
            bool? keepLinesTogether,
            bool? keepWithNext,
            bool? pageBreakBefore,
            bool? avoidWidowAndOrphan) {
            Alignment = alignment;
            StyleIndex = styleIndex;
            SpacingBeforeTwips = spacingBeforeTwips;
            SpacingAfterTwips = spacingAfterTwips;
            LineSpacingTwips = lineSpacingTwips;
            LeftIndentTwips = leftIndentTwips;
            RightIndentTwips = rightIndentTwips;
            FirstLineIndentTwips = firstLineIndentTwips;
            KeepLinesTogether = keepLinesTogether;
            KeepWithNext = keepWithNext;
            PageBreakBefore = pageBreakBefore;
            AvoidWidowAndOrphan = avoidWidowAndOrphan;
        }

        internal byte? Alignment { get; }

        internal ushort? StyleIndex { get; }

        internal int? SpacingBeforeTwips { get; }

        internal int? SpacingAfterTwips { get; }

        internal int? LineSpacingTwips { get; }

        internal int? LeftIndentTwips { get; }

        internal int? RightIndentTwips { get; }

        internal int? FirstLineIndentTwips { get; }

        internal bool? KeepLinesTogether { get; }

        internal bool? KeepWithNext { get; }

        internal bool? PageBreakBefore { get; }

        internal bool? AvoidWidowAndOrphan { get; }

        internal bool HasFormatting => Alignment != null
            || StyleIndex != null
            || SpacingBeforeTwips != null
            || SpacingAfterTwips != null
            || LineSpacingTwips != null
            || LeftIndentTwips != null
            || RightIndentTwips != null
            || FirstLineIndentTwips != null
            || KeepLinesTogether != null
            || KeepWithNext != null
            || PageBreakBefore != null
            || AvoidWidowAndOrphan != null;

        public bool Equals(LegacyDocWritableParagraphFormatting other) {
            return Alignment == other.Alignment
                && StyleIndex == other.StyleIndex
                && SpacingBeforeTwips == other.SpacingBeforeTwips
                && SpacingAfterTwips == other.SpacingAfterTwips
                && LineSpacingTwips == other.LineSpacingTwips
                && LeftIndentTwips == other.LeftIndentTwips
                && RightIndentTwips == other.RightIndentTwips
                && FirstLineIndentTwips == other.FirstLineIndentTwips
                && KeepLinesTogether == other.KeepLinesTogether
                && KeepWithNext == other.KeepWithNext
                && PageBreakBefore == other.PageBreakBefore
                && AvoidWidowAndOrphan == other.AvoidWidowAndOrphan;
        }

        public override bool Equals(object? obj) {
            return obj is LegacyDocWritableParagraphFormatting other && Equals(other);
        }

        public override int GetHashCode() {
            int hash = 17;
            hash = (hash * 31) + Alignment.GetHashCode();
            hash = (hash * 31) + StyleIndex.GetHashCode();
            hash = (hash * 31) + SpacingBeforeTwips.GetHashCode();
            hash = (hash * 31) + SpacingAfterTwips.GetHashCode();
            hash = (hash * 31) + LineSpacingTwips.GetHashCode();
            hash = (hash * 31) + LeftIndentTwips.GetHashCode();
            hash = (hash * 31) + RightIndentTwips.GetHashCode();
            hash = (hash * 31) + FirstLineIndentTwips.GetHashCode();
            hash = (hash * 31) + KeepLinesTogether.GetHashCode();
            hash = (hash * 31) + KeepWithNext.GetHashCode();
            hash = (hash * 31) + PageBreakBefore.GetHashCode();
            hash = (hash * 31) + AvoidWidowAndOrphan.GetHashCode();
            return hash;
        }
    }

    internal readonly struct LegacyDocWritableParagraph {
        internal LegacyDocWritableParagraph(int startCharacter, int length, LegacyDocWritableParagraphFormatting formatting) {
            StartCharacter = startCharacter;
            Length = length;
            Formatting = formatting;
        }

        internal int StartCharacter { get; }

        internal int Length { get; }

        internal int EndCharacter => StartCharacter + Length;

        internal LegacyDocWritableParagraphFormatting Formatting { get; }
    }

    internal readonly struct LegacyDocWritableParagraphSegment {
        internal LegacyDocWritableParagraphSegment(int startCharacter, int length, LegacyDocWritableParagraphFormatting formatting) {
            StartCharacter = startCharacter;
            Length = length;
            Formatting = formatting;
        }

        internal int StartCharacter { get; }

        internal int Length { get; }

        internal int EndCharacter => StartCharacter + Length;

        internal LegacyDocWritableParagraphFormatting Formatting { get; }

        internal LegacyDocWritableParagraphSegment Extend(int additionalLength) {
            return new LegacyDocWritableParagraphSegment(StartCharacter, Length + additionalLength, Formatting);
        }
    }
}
