namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static class LegacyDocCharacterFormattingReader {
        private const int OleSectorSize = 512;
        private const ushort SprmCFBold = 0x0835;
        private const ushort SprmCFItalic = 0x0836;
        private const ushort SprmCFStrike = 0x0837;
        private const ushort SprmCKul = 0x2A3E;
        private const ushort SprmCIco = 0x2A42;
        private const ushort SprmCHps = 0x4A43;
        private const ushort SprmCRgFtc0 = 0x4A4F;
        private const ushort SprmCCv = 0x6870;

        internal static IReadOnlyList<LegacyDocCharacterFormatRange> ReadCharacterFormatting(
            byte[] wordDocumentStream,
            byte[] tableStream,
            LegacyDocFib fib,
            IReadOnlyList<string> fontFamilies,
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

                ReadChpxFkp(wordDocumentStream, pageOffset, ranges, fontFamilies);
            }

            return ranges
                .OrderBy(range => range.FileOffsetStart)
                .ThenBy(range => range.FileOffsetEnd)
                .ToArray();
        }

        private static void ReadChpxFkp(byte[] wordDocumentStream, int pageOffset, List<LegacyDocCharacterFormatRange> ranges, IReadOnlyList<string> fontFamilies) {
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

                LegacyDocCharacterFormat format = ReadGrpprl(wordDocumentStream, grpprlOffset, cbGrpprl, fontFamilies);
                if (format.HasFormatting) {
                    ranges.Add(new LegacyDocCharacterFormatRange(fcStart, fcEnd, format));
                }
            }
        }

        internal static LegacyDocCharacterFormat ReadGrpprl(byte[] bytes, int offset, int count, IReadOnlyList<string> fontFamilies) {
            int end = offset + count;
            bool bold = false;
            bool italic = false;
            bool strike = false;
            LegacyDocUnderlineKind? underline = null;
            int? fontSizeHalfPoints = null;
            string? colorHex = null;
            string? fontFamily = null;

            while (offset + 2 <= end) {
                ushort sprm = LegacyDocFib.ReadUInt16(bytes, offset);
                if (sprm == SprmCFBold || sprm == SprmCFItalic || sprm == SprmCFStrike) {
                    if (offset + 3 > end) {
                        break;
                    }

                    bool enabled = bytes[offset + 2] != 0;
                    if (sprm == SprmCFBold) {
                        bold = enabled;
                    } else if (sprm == SprmCFItalic) {
                        italic = enabled;
                    } else {
                        strike = enabled;
                    }

                    offset += 3;
                    continue;
                }

                if (sprm == SprmCKul) {
                    if (offset + 3 > end) {
                        break;
                    }

                    underline = MapUnderline(bytes[offset + 2]);
                    offset += 3;
                    continue;
                }

                if (sprm == SprmCIco) {
                    if (offset + 3 > end) {
                        break;
                    }

                    colorHex = MapIndexedColor(bytes[offset + 2]);
                    offset += 3;
                    continue;
                }

                if (sprm == SprmCHps) {
                    if (offset + 4 > end) {
                        break;
                    }

                    fontSizeHalfPoints = LegacyDocFib.ReadUInt16(bytes, offset + 2);
                    offset += 4;
                    continue;
                }

                if (sprm == SprmCRgFtc0) {
                    if (offset + 4 > end) {
                        break;
                    }

                    int fontIndex = LegacyDocFib.ReadUInt16(bytes, offset + 2);
                    if (fontIndex >= 0 && fontIndex < fontFamilies.Count && !string.IsNullOrWhiteSpace(fontFamilies[fontIndex])) {
                        fontFamily = fontFamilies[fontIndex];
                    }

                    offset += 4;
                    continue;
                }

                if (sprm == SprmCCv) {
                    if (offset + 6 > end) {
                        break;
                    }

                    colorHex = ReadColorRef(bytes, offset + 2);
                    offset += 6;
                    continue;
                }

                if (!TryGetSprmOperandLength(bytes, offset, end, out int operandLength)) {
                    break;
                }

                offset += 2 + operandLength;
            }

            return new LegacyDocCharacterFormat(bold, italic, strike, underline, fontSizeHalfPoints, colorHex, fontFamily);
        }

        private static LegacyDocUnderlineKind? MapUnderline(byte value) {
            switch (value) {
                case 0:
                case 5:
                    return null;
                case 1:
                    return LegacyDocUnderlineKind.Single;
                case 2:
                    return LegacyDocUnderlineKind.Words;
                case 3:
                    return LegacyDocUnderlineKind.Double;
                case 4:
                    return LegacyDocUnderlineKind.Dotted;
                case 6:
                    return LegacyDocUnderlineKind.Thick;
                case 7:
                    return LegacyDocUnderlineKind.Dash;
                case 8:
                    return LegacyDocUnderlineKind.DotDash;
                case 9:
                    return LegacyDocUnderlineKind.DotDotDash;
                case 10:
                    return LegacyDocUnderlineKind.Wave;
                case 11:
                    return LegacyDocUnderlineKind.DottedHeavy;
                case 12:
                    return LegacyDocUnderlineKind.DashedHeavy;
                case 13:
                    return LegacyDocUnderlineKind.DashDotHeavy;
                case 14:
                    return LegacyDocUnderlineKind.DashDotDotHeavy;
                case 15:
                    return LegacyDocUnderlineKind.WavyHeavy;
                case 16:
                    return LegacyDocUnderlineKind.DashLong;
                case 17:
                    return LegacyDocUnderlineKind.WavyDouble;
                case 18:
                    return LegacyDocUnderlineKind.DashLongHeavy;
                default:
                    return null;
            }
        }

        private static string? MapIndexedColor(byte value) {
            switch (value) {
                case 0:
                    return null;
                case 1:
                    return "000000";
                case 2:
                    return "0000ff";
                case 3:
                    return "00ffff";
                case 4:
                    return "00ff00";
                case 5:
                    return "ff00ff";
                case 6:
                    return "ff0000";
                case 7:
                    return "ffff00";
                case 8:
                    return "ffffff";
                case 9:
                    return "000080";
                case 10:
                    return "008080";
                case 11:
                    return "008000";
                case 12:
                    return "800080";
                case 13:
                    return "800000";
                case 14:
                    return "808000";
                case 15:
                    return "808080";
                case 16:
                    return "c0c0c0";
                default:
                    return null;
            }
        }

        private static string ReadColorRef(byte[] bytes, int offset) {
            var chars = new char[6];
            WriteHexByte(chars, 0, bytes[offset]);
            WriteHexByte(chars, 2, bytes[offset + 1]);
            WriteHexByte(chars, 4, bytes[offset + 2]);
            return new string(chars);
        }

        private static void WriteHexByte(char[] destination, int offset, byte value) {
            const string hex = "0123456789abcdef";
            destination[offset] = hex[value >> 4];
            destination[offset + 1] = hex[value & 0x0F];
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
