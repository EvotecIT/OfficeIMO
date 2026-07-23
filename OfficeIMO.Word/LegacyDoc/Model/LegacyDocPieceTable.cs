namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static class LegacyDocPieceTable {
        private const byte ClxPieceTableMarker = 0x02;
        private const byte ClxPrlMarker = 0x01;
        private const uint CompressedTextFlag = 0x40000000;
        private const uint FcMask = 0x3FFFFFFF;

        internal static bool TryRead(byte[] wordDocumentStream, byte[] tableStream, LegacyDocFib fib, int maxDecodedCharacters, out LegacyDocTextContent content, out string? error) {
            content = new LegacyDocTextContent(string.Empty, Array.Empty<LegacyDocTextCharacter>());
            error = null;
            long totalCharacterCountLong = (long)fib.CcpText
                + fib.CcpFtn
                + fib.CcpHdd
                + fib.CcpAtn
                + fib.CcpEdn
                + fib.CcpTxbx
                + fib.CcpHdrTxbx;

            if (totalCharacterCountLong == 0) {
                return true;
            }
            if (totalCharacterCountLong < 0 || totalCharacterCountLong > maxDecodedCharacters || totalCharacterCountLong > int.MaxValue) {
                error = "The FIB decoded character count exceeds MaxDecodedCharacters.";
                return false;
            }
            int totalCharacterCount = (int)totalCharacterCountLong;

            if (fib.FcClx < 0 || fib.LcbClx <= 0 || fib.FcClx + fib.LcbClx > tableStream.Length) {
                error = "The FIB points outside the selected table stream for the CLX piece table.";
                return false;
            }

            int clxOffset = fib.FcClx;
            int clxEnd = fib.FcClx + fib.LcbClx;
            while (clxOffset < clxEnd && tableStream[clxOffset] == ClxPrlMarker) {
                if (clxOffset + 3 > clxEnd) {
                    error = "The CLX property modifier block is truncated.";
                    return false;
                }

                int cbGrpprl = LegacyDocFib.ReadUInt16(tableStream, clxOffset + 1);
                clxOffset += 3 + cbGrpprl;
            }

            if (clxOffset >= clxEnd || tableStream[clxOffset] != ClxPieceTableMarker) {
                error = "The CLX does not contain a PLCFPCD piece-table block.";
                return false;
            }

            if (clxOffset + 5 > clxEnd) {
                error = "The PLCFPCD block header is truncated.";
                return false;
            }

            int pcdByteCount = LegacyDocFib.ReadInt32(tableStream, clxOffset + 1);
            int pcdOffset = clxOffset + 5;
            if (pcdByteCount < 4 || pcdOffset + pcdByteCount > clxEnd || (pcdByteCount - 4) % 12 != 0) {
                error = "The PLCFPCD block has an invalid length.";
                return false;
            }

            int pieceCount = (pcdByteCount - 4) / 12;
            var allCharacters = new List<LegacyDocTextCharacter>(totalCharacterCount);
            int cpArrayOffset = pcdOffset;
            int pcdArrayOffset = cpArrayOffset + ((pieceCount + 1) * 4);

            long appendedCharacterCount = 0;
            for (int i = 0; i < pieceCount; i++) {
                int cpStart = LegacyDocFib.ReadInt32(tableStream, cpArrayOffset + (i * 4));
                int cpEnd = LegacyDocFib.ReadInt32(tableStream, cpArrayOffset + ((i + 1) * 4));
                if (cpStart < 0 || cpEnd < cpStart) {
                    error = "The PLCFPCD character positions are not monotonic.";
                    return false;
                }
                if (cpEnd == cpStart) {
                    continue;
                }

                int decodedCharacterCount = Math.Min(cpEnd, totalCharacterCount) - cpStart;
                if (decodedCharacterCount <= 0) {
                    break;
                }
                appendedCharacterCount += decodedCharacterCount;
                if (appendedCharacterCount > maxDecodedCharacters || appendedCharacterCount > totalCharacterCount) {
                    error = "The PLCFPCD decoded character count exceeds MaxDecodedCharacters.";
                    return false;
                }

                uint fcCompressed = unchecked((uint)LegacyDocFib.ReadInt32(tableStream, pcdArrayOffset + (i * 8) + 2));
                bool compressed = (fcCompressed & CompressedTextFlag) != 0;
                uint fileCharacterPosition = fcCompressed & FcMask;
                int byteOffset = compressed
                    ? checked((int)(fileCharacterPosition / 2))
                    : checked((int)fileCharacterPosition);

                if (compressed) {
                    if (byteOffset + decodedCharacterCount > wordDocumentStream.Length) {
                        error = "A compressed text piece points outside the WordDocument stream.";
                        return false;
                    }

                    AppendWindows1252(allCharacters, wordDocumentStream, byteOffset, decodedCharacterCount, cpStart);
                } else {
                    int byteCount = checked(decodedCharacterCount * 2);
                    if (byteOffset + byteCount > wordDocumentStream.Length) {
                        error = "A Unicode text piece points outside the WordDocument stream.";
                        return false;
                    }

                    AppendUnicode(allCharacters, wordDocumentStream, byteOffset, decodedCharacterCount, cpStart);
                }
            }

            LegacyDocTextCharacter[] bodyCharacters = allCharacters
                .Where(character => character.CharacterPosition < fib.CcpText)
                .ToArray();
            content = new LegacyDocTextContent(new string(bodyCharacters.Select(character => character.Character).ToArray()), bodyCharacters, allCharacters);
            return true;
        }

        private static void AppendWindows1252(List<LegacyDocTextCharacter> characters, byte[] bytes, int offset, int count, int characterPositionStart) {
            for (int i = 0; i < count; i++) {
                char character = DecodeWindows1252(bytes[offset + i]);
                characters.Add(new LegacyDocTextCharacter(character, offset + i, characterPositionStart + i));
            }
        }

        private static void AppendUnicode(List<LegacyDocTextCharacter> characters, byte[] bytes, int offset, int count, int characterPositionStart) {
            for (int i = 0; i < count; i++) {
                int byteOffset = offset + (i * 2);
                char character = (char)(bytes[byteOffset] | (bytes[byteOffset + 1] << 8));
                characters.Add(new LegacyDocTextCharacter(character, byteOffset, characterPositionStart + i));
            }
        }

        private static char DecodeWindows1252(byte value) {
            if (value < 0x80 || value >= 0xA0) {
                return (char)value;
            }

            return value switch {
                0x80 => '\u20AC',
                0x82 => '\u201A',
                0x83 => '\u0192',
                0x84 => '\u201E',
                0x85 => '\u2026',
                0x86 => '\u2020',
                0x87 => '\u2021',
                0x88 => '\u02C6',
                0x89 => '\u2030',
                0x8A => '\u0160',
                0x8B => '\u2039',
                0x8C => '\u0152',
                0x8E => '\u017D',
                0x91 => '\u2018',
                0x92 => '\u2019',
                0x93 => '\u201C',
                0x94 => '\u201D',
                0x95 => '\u2022',
                0x96 => '\u2013',
                0x97 => '\u2014',
                0x98 => '\u02DC',
                0x99 => '\u2122',
                0x9A => '\u0161',
                0x9B => '\u203A',
                0x9C => '\u0153',
                0x9E => '\u017E',
                0x9F => '\u0178',
                _ => '\uFFFD'
            };
        }
    }
}
