namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static class LegacyDocCharacterFormattingReader {
        private const int OleSectorSize = 512;
        private const ushort SprmCFBold = 0x0835;
        private const ushort SprmCFItalic = 0x0836;

        internal static IReadOnlyList<LegacyDocCharacterFormatRange> ReadCharacterFormatting(
            byte[] wordDocumentStream,
            byte[] tableStream,
            LegacyDocFib fib,
            out string? warning) {
            warning = null;

            if (fib.LcbPlcfBteChpx == 0) {
                return Array.Empty<LegacyDocCharacterFormatRange>();
            }

            if (fib.FcPlcfBteChpx < 0
                || fib.LcbPlcfBteChpx < 4
                || fib.FcPlcfBteChpx + fib.LcbPlcfBteChpx > tableStream.Length
                || (fib.LcbPlcfBteChpx - 4) % 8 != 0) {
                warning = "The FIB points outside the selected table stream for the character-format bin table.";
                return Array.Empty<LegacyDocCharacterFormatRange>();
            }

            int binCount = (fib.LcbPlcfBteChpx - 4) / 8;
            int cpArrayOffset = fib.FcPlcfBteChpx;
            int bteArrayOffset = cpArrayOffset + ((binCount + 1) * 4);
            var ranges = new List<LegacyDocCharacterFormatRange>();

            for (int binIndex = 0; binIndex < binCount; binIndex++) {
                int fcStart = LegacyDocFib.ReadInt32(tableStream, cpArrayOffset + (binIndex * 4));
                int fcEnd = LegacyDocFib.ReadInt32(tableStream, cpArrayOffset + ((binIndex + 1) * 4));
                int pageNumber = LegacyDocFib.ReadInt32(tableStream, bteArrayOffset + (binIndex * 4));
                if (fcEnd <= fcStart) {
                    continue;
                }

                int pageOffset = checked(pageNumber * OleSectorSize);
                if (pageOffset < 0 || pageOffset + OleSectorSize > wordDocumentStream.Length) {
                    warning = "A character-format bin table entry points outside the WordDocument stream.";
                    return ranges;
                }

                ReadChpxFkp(wordDocumentStream, pageOffset, ranges);
            }

            return ranges
                .OrderBy(range => range.FileOffsetStart)
                .ThenBy(range => range.FileOffsetEnd)
                .ToArray();
        }

        private static void ReadChpxFkp(byte[] wordDocumentStream, int pageOffset, List<LegacyDocCharacterFormatRange> ranges) {
            int crun = wordDocumentStream[pageOffset + OleSectorSize - 1];
            if (crun <= 0) {
                return;
            }

            int rgfcOffset = pageOffset;
            int rgbOffset = pageOffset + ((crun + 1) * 4);
            if (rgbOffset + crun > pageOffset + OleSectorSize - 1) {
                return;
            }

            for (int runIndex = 0; runIndex < crun; runIndex++) {
                int fcStart = LegacyDocFib.ReadInt32(wordDocumentStream, rgfcOffset + (runIndex * 4));
                int fcEnd = LegacyDocFib.ReadInt32(wordDocumentStream, rgfcOffset + ((runIndex + 1) * 4));
                if (fcEnd <= fcStart) {
                    continue;
                }

                int chpxOffset = wordDocumentStream[rgbOffset + runIndex] * 2;
                if (chpxOffset == 0) {
                    continue;
                }

                int absoluteChpxOffset = pageOffset + chpxOffset;
                if (absoluteChpxOffset >= pageOffset + OleSectorSize - 1) {
                    continue;
                }

                int cbGrpprl = wordDocumentStream[absoluteChpxOffset];
                int grpprlOffset = absoluteChpxOffset + 1;
                if (cbGrpprl <= 0 || grpprlOffset + cbGrpprl > pageOffset + OleSectorSize - 1) {
                    continue;
                }

                LegacyDocCharacterFormat format = ReadGrpprl(wordDocumentStream, grpprlOffset, cbGrpprl);
                if (format.Bold || format.Italic) {
                    ranges.Add(new LegacyDocCharacterFormatRange(fcStart, fcEnd, format));
                }
            }
        }

        private static LegacyDocCharacterFormat ReadGrpprl(byte[] bytes, int offset, int count) {
            int end = offset + count;
            bool bold = false;
            bool italic = false;

            while (offset + 2 <= end) {
                ushort sprm = LegacyDocFib.ReadUInt16(bytes, offset);
                if (sprm == SprmCFBold || sprm == SprmCFItalic) {
                    if (offset + 3 > end) {
                        break;
                    }

                    bool enabled = bytes[offset + 2] != 0;
                    if (sprm == SprmCFBold) {
                        bold = enabled;
                    } else {
                        italic = enabled;
                    }

                    offset += 3;
                    continue;
                }

                if (!TryGetSprmOperandLength(bytes, offset, end, out int operandLength)) {
                    break;
                }

                offset += 2 + operandLength;
            }

            return new LegacyDocCharacterFormat(bold, italic);
        }

        private static bool TryGetSprmOperandLength(byte[] bytes, int sprmOffset, int end, out int operandLength) {
            operandLength = 0;
            ushort sprm = LegacyDocFib.ReadUInt16(bytes, sprmOffset);
            int spra = (sprm >> 13) & 0x7;
            switch (spra) {
                case 0:
                case 1:
                    operandLength = 1;
                    return sprmOffset + 2 + operandLength <= end;
                case 2:
                case 4:
                case 5:
                    operandLength = 2;
                    return sprmOffset + 2 + operandLength <= end;
                case 3:
                    operandLength = 4;
                    return sprmOffset + 2 + operandLength <= end;
                case 6:
                    if (sprmOffset + 3 > end) {
                        return false;
                    }

                    operandLength = 1 + bytes[sprmOffset + 2];
                    return sprmOffset + 2 + operandLength <= end;
                case 7:
                    operandLength = 3;
                    return sprmOffset + 2 + operandLength <= end;
                default:
                    return false;
            }
        }
    }
}
