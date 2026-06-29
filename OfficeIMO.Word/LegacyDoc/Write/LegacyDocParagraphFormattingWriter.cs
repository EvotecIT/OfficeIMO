namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static class LegacyDocParagraphFormattingWriter {
        private const int PapxFkpBxLength = 13;
        private const ushort SprmPJc = 0x2461;

        internal static void WritePapxFkp(byte[] stream, int pageOffset, int textOffset, int oleSectorSize, IReadOnlyList<LegacyDocWritableParagraphSegment> segments) {
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
                WriteInt32(stream, pageOffset + (index * 4), textOffset + (segment.StartCharacter * 2));
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
            WriteInt32(stream, pageOffset + (segments.Count * 4), textOffset + (lastSegment.EndCharacter * 2));
            stream[pageOffset + oleSectorSize - 1] = (byte)segments.Count;
        }

        private static byte[] CreatePapx(LegacyDocWritableParagraphFormatting formatting) {
            var grpprl = new List<byte>(6) {
                0,
                0
            };

            if (formatting.Alignment != null) {
                AddSingleByteSprm(grpprl, SprmPJc, formatting.Alignment.Value);
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

        private static void WriteInt32(byte[] bytes, int offset, int value) {
            bytes[offset] = (byte)value;
            bytes[offset + 1] = (byte)(value >> 8);
            bytes[offset + 2] = (byte)(value >> 16);
            bytes[offset + 3] = (byte)(value >> 24);
        }
    }

    internal readonly struct LegacyDocWritableParagraphFormatting : IEquatable<LegacyDocWritableParagraphFormatting> {
        internal static readonly LegacyDocWritableParagraphFormatting Plain = new LegacyDocWritableParagraphFormatting(null);

        internal LegacyDocWritableParagraphFormatting(byte? alignment) {
            Alignment = alignment;
        }

        internal byte? Alignment { get; }

        internal bool HasFormatting => Alignment != null;

        public bool Equals(LegacyDocWritableParagraphFormatting other) {
            return Alignment == other.Alignment;
        }

        public override bool Equals(object? obj) {
            return obj is LegacyDocWritableParagraphFormatting other && Equals(other);
        }

        public override int GetHashCode() {
            return Alignment.GetHashCode();
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
