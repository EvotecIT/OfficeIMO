namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static class LegacyDocCharacterFormattingReader {
        private const int OleSectorSize = 512;
        private const ushort SprmCFBold = 0x0835;
        private const ushort SprmCFItalic = 0x0836;
        private const ushort SprmCFStrike = 0x0837;
        private const ushort SprmCFOutline = 0x0838;
        private const ushort SprmCFShadow = 0x0839;
        private const ushort SprmCFImprint = 0x0854;
        private const ushort SprmCFSmallCaps = 0x083A;
        private const ushort SprmCFCaps = 0x083B;
        private const ushort SprmCFVanish = 0x083C;
        private const ushort SprmCFEmboss = 0x0858;
        private const ushort SprmCFNoProof = 0x0875;
        private const ushort SprmCHighlight = 0x2A0C;
        private const ushort SprmCKul = 0x2A3E;
        private const ushort SprmCDxaSpace = 0x8840;
        private const ushort SprmCIco = 0x2A42;
        private const ushort SprmCIss = 0x2A48;
        private const ushort SprmCHps = 0x4A43;
        private const ushort SprmCRgLid0 = 0x486D;
        private const ushort SprmCRgLid1 = 0x486E;
        private const ushort SprmCRgFtc0 = 0x4A4F;
        private const ushort SprmCFDStrike = 0x2A53;
        private const ushort SprmCCv = 0x6870;
        private const ushort SprmCPicLocation = 0x6A03;
        private const ushort SprmCFRMarkDel = 0x0800;
        private const ushort SprmCFRMarkIns = 0x0801;
        private const ushort SprmCIbstRMark = 0x4804;
        private const ushort SprmCDttmRMark = 0x6805;
        private const ushort SprmCIbstRMarkDel = 0x4863;
        private const ushort SprmCDttmRMarkDel = 0x6864;

        internal static IReadOnlyList<LegacyDocCharacterFormatRange> ReadCharacterFormatting(
            byte[] wordDocumentStream,
            byte[] tableStream,
            LegacyDocFib fib,
            IReadOnlyList<string> fontFamilies,
            IReadOnlyList<string> revisionAuthors,
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

                ReadChpxFkp(wordDocumentStream, pageOffset, ranges, fontFamilies, revisionAuthors);
            }

            return ranges
                .OrderBy(range => range.FileOffsetStart)
                .ThenBy(range => range.FileOffsetEnd)
                .ToArray();
        }

        private static void ReadChpxFkp(
            byte[] wordDocumentStream,
            int pageOffset,
            List<LegacyDocCharacterFormatRange> ranges,
            IReadOnlyList<string> fontFamilies,
            IReadOnlyList<string> revisionAuthors) {
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

                LegacyDocCharacterFormat format = ReadGrpprl(wordDocumentStream, grpprlOffset, cbGrpprl, fontFamilies, revisionAuthors);
                if (format.HasFormatting) {
                    ranges.Add(new LegacyDocCharacterFormatRange(fcStart, fcEnd, format));
                }
            }
        }

        internal static LegacyDocCharacterFormat ReadGrpprl(
            byte[] bytes,
            int offset,
            int count,
            IReadOnlyList<string> fontFamilies,
            IReadOnlyList<string>? revisionAuthors = null) {
            int end = offset + count;
            bool bold = false;
            bool italic = false;
            bool strike = false;
            bool doubleStrike = false;
            bool outline = false;
            bool shadow = false;
            bool emboss = false;
            bool imprint = false;
            bool hidden = false;
            bool noProof = false;
            bool smallCaps = false;
            bool caps = false;
            LegacyDocVerticalPositionKind? verticalPosition = null;
            LegacyDocUnderlineKind? underline = null;
            LegacyDocHighlightColorKind? highlight = null;
            int? fontSizeHalfPoints = null;
            string? colorHex = null;
            string? fontFamily = null;
            int? characterSpacingTwips = null;
            string? language = null;
            string? eastAsiaLanguage = null;
            int? pictureDataOffset = null;
            bool inserted = false;
            bool deleted = false;
            int insertedAuthorIndex = 0;
            int deletedAuthorIndex = 0;
            DateTime? insertedDate = null;
            DateTime? deletedDate = null;
            LegacyDocCharacterFormatProperties specified = LegacyDocCharacterFormatProperties.None;

            while (offset + 2 <= end) {
                ushort sprm = LegacyDocFib.ReadUInt16(bytes, offset);
                if (sprm == SprmCFRMarkIns || sprm == SprmCFRMarkDel) {
                    if (offset + 3 > end) {
                        break;
                    }

                    bool enabled = bytes[offset + 2] != 0;
                    if (sprm == SprmCFRMarkIns) {
                        inserted = enabled;
                    } else {
                        deleted = enabled;
                    }

                    offset += 3;
                    continue;
                }

                if (sprm == SprmCIbstRMark || sprm == SprmCIbstRMarkDel) {
                    if (offset + 4 > end) {
                        break;
                    }

                    int authorIndex = unchecked((short)LegacyDocFib.ReadUInt16(bytes, offset + 2));
                    if (sprm == SprmCIbstRMark) {
                        insertedAuthorIndex = authorIndex;
                    } else {
                        deletedAuthorIndex = authorIndex;
                    }

                    offset += 4;
                    continue;
                }

                if (sprm == SprmCDttmRMark || sprm == SprmCDttmRMarkDel) {
                    if (offset + 6 > end) {
                        break;
                    }

                    DateTime? date = ReadDttm(unchecked((uint)LegacyDocFib.ReadInt32(bytes, offset + 2)));
                    if (sprm == SprmCDttmRMark) {
                        insertedDate = date;
                    } else {
                        deletedDate = date;
                    }

                    offset += 6;
                    continue;
                }

                if (sprm == SprmCFBold || sprm == SprmCFItalic || sprm == SprmCFStrike || sprm == SprmCFOutline || sprm == SprmCFShadow || sprm == SprmCFSmallCaps || sprm == SprmCFCaps || sprm == SprmCFVanish || sprm == SprmCFImprint || sprm == SprmCFEmboss || sprm == SprmCFNoProof || sprm == SprmCFDStrike) {
                    if (offset + 3 > end) {
                        break;
                    }

                    bool enabled = bytes[offset + 2] != 0;
                    if (sprm == SprmCFBold) {
                        bold = enabled;
                        specified |= LegacyDocCharacterFormatProperties.Bold;
                    } else if (sprm == SprmCFItalic) {
                        italic = enabled;
                        specified |= LegacyDocCharacterFormatProperties.Italic;
                    } else if (sprm == SprmCFStrike) {
                        strike = enabled;
                        specified |= LegacyDocCharacterFormatProperties.Strike;
                    } else if (sprm == SprmCFOutline) {
                        outline = enabled;
                        specified |= LegacyDocCharacterFormatProperties.Outline;
                    } else if (sprm == SprmCFShadow) {
                        shadow = enabled;
                        specified |= LegacyDocCharacterFormatProperties.Shadow;
                    } else if (sprm == SprmCFEmboss) {
                        emboss = enabled;
                        specified |= LegacyDocCharacterFormatProperties.Emboss;
                    } else if (sprm == SprmCFImprint) {
                        imprint = enabled;
                        specified |= LegacyDocCharacterFormatProperties.Imprint;
                    } else if (sprm == SprmCFVanish) {
                        hidden = enabled;
                        specified |= LegacyDocCharacterFormatProperties.Hidden;
                    } else if (sprm == SprmCFNoProof) {
                        noProof = enabled;
                        specified |= LegacyDocCharacterFormatProperties.NoProof;
                    } else if (sprm == SprmCFSmallCaps) {
                        smallCaps = enabled;
                        specified |= LegacyDocCharacterFormatProperties.SmallCaps;
                    } else if (sprm == SprmCFDStrike) {
                        doubleStrike = enabled;
                        specified |= LegacyDocCharacterFormatProperties.DoubleStrike;
                    } else {
                        caps = enabled;
                        specified |= LegacyDocCharacterFormatProperties.Caps;
                    }

                    offset += 3;
                    continue;
                }

                if (sprm == SprmCHighlight) {
                    if (offset + 3 > end) {
                        break;
                    }

                    highlight = MapHighlight(bytes[offset + 2]);
                    specified |= LegacyDocCharacterFormatProperties.Highlight;
                    offset += 3;
                    continue;
                }

                if (sprm == SprmCKul) {
                    if (offset + 3 > end) {
                        break;
                    }

                    underline = MapUnderline(bytes[offset + 2]);
                    specified |= LegacyDocCharacterFormatProperties.Underline;
                    offset += 3;
                    continue;
                }

                if (sprm == SprmCIss) {
                    if (offset + 3 > end) {
                        break;
                    }

                    verticalPosition = MapVerticalPosition(bytes[offset + 2]);
                    specified |= LegacyDocCharacterFormatProperties.VerticalPosition;
                    offset += 3;
                    continue;
                }

                if (sprm == SprmCIco) {
                    if (offset + 3 > end) {
                        break;
                    }

                    colorHex = MapIndexedColor(bytes[offset + 2]);
                    specified |= LegacyDocCharacterFormatProperties.Color;
                    offset += 3;
                    continue;
                }

                if (sprm == SprmCDxaSpace) {
                    if (offset + 4 > end) {
                        break;
                    }

                    characterSpacingTwips = unchecked((short)LegacyDocFib.ReadUInt16(bytes, offset + 2));
                    specified |= LegacyDocCharacterFormatProperties.CharacterSpacing;
                    offset += 4;
                    continue;
                }

                if (sprm == SprmCHps) {
                    if (offset + 4 > end) {
                        break;
                    }

                    fontSizeHalfPoints = LegacyDocFib.ReadUInt16(bytes, offset + 2);
                    specified |= LegacyDocCharacterFormatProperties.FontSize;
                    offset += 4;
                    continue;
                }

                if (sprm == SprmCRgLid0 || sprm == SprmCRgLid1) {
                    if (offset + 4 > end) {
                        break;
                    }

                    string? languageTag = LegacyDocLanguageMapper.TryGetLanguageTag(LegacyDocFib.ReadUInt16(bytes, offset + 2));
                    if (languageTag != null) {
                        if (sprm == SprmCRgLid0) {
                            language = languageTag;
                        } else {
                            eastAsiaLanguage = languageTag;
                        }

                        specified |= LegacyDocCharacterFormatProperties.Language;
                    }

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
                        specified |= LegacyDocCharacterFormatProperties.FontFamily;
                    }

                    offset += 4;
                    continue;
                }

                if (sprm == SprmCCv) {
                    if (offset + 6 > end) {
                        break;
                    }

                    colorHex = ReadColorRef(bytes, offset + 2);
                    specified |= LegacyDocCharacterFormatProperties.Color;
                    offset += 6;
                    continue;
                }

                if (sprm == SprmCPicLocation) {
                    if (offset + 6 > end) {
                        break;
                    }

                    pictureDataOffset = LegacyDocFib.ReadInt32(bytes, offset + 2);
                    offset += 6;
                    continue;
                }

                if (!TryGetSprmOperandLength(bytes, offset, end, out int operandLength)) {
                    break;
                }

                offset += 2 + operandLength;
            }

            LegacyDocCapsKind? capsKind = caps
                ? LegacyDocCapsKind.Caps
                : smallCaps ? LegacyDocCapsKind.SmallCaps : null;
            LegacyDocRevision revision = deleted
                ? new LegacyDocRevision(LegacyDocRevisionKind.Deleted, ResolveRevisionAuthor(revisionAuthors, deletedAuthorIndex), deletedDate)
                : inserted
                    ? new LegacyDocRevision(LegacyDocRevisionKind.Inserted, ResolveRevisionAuthor(revisionAuthors, insertedAuthorIndex), insertedDate)
                    : LegacyDocRevision.None;

            return new LegacyDocCharacterFormat(
                bold,
                italic,
                strike,
                doubleStrike,
                outline,
                shadow,
                emboss,
                imprint,
                hidden,
                noProof,
                capsKind,
                verticalPosition,
                underline,
                highlight,
                fontSizeHalfPoints,
                colorHex,
                fontFamily,
                characterSpacingTwips,
                language,
                eastAsiaLanguage,
                specified,
                pictureDataOffset,
                revision);
        }

        private static string ResolveRevisionAuthor(IReadOnlyList<string>? revisionAuthors, int authorIndex) {
            if (revisionAuthors != null
                && authorIndex >= 0
                && authorIndex < revisionAuthors.Count
                && !string.IsNullOrWhiteSpace(revisionAuthors[authorIndex])) {
                return revisionAuthors[authorIndex];
            }

            return LegacyDocRevisionAuthorReader.UnknownAuthor;
        }

        private static DateTime? ReadDttm(uint value) {
            if (value == 0) {
                return null;
            }

            int minute = (int)(value & 0x3F);
            int hour = (int)((value >> 6) & 0x1F);
            int day = (int)((value >> 11) & 0x1F);
            int month = (int)((value >> 16) & 0x0F);
            int year = 1900 + (int)((value >> 20) & 0x1FF);
            if (minute > 59 || hour > 23 || month < 1 || month > 12 || day < 1 || day > DateTime.DaysInMonth(year, month)) {
                return null;
            }

            return new DateTime(year, month, day, hour, minute, 0, DateTimeKind.Unspecified);
        }

        private static LegacyDocVerticalPositionKind? MapVerticalPosition(byte value) {
            switch (value) {
                case 1:
                    return LegacyDocVerticalPositionKind.Superscript;
                case 2:
                    return LegacyDocVerticalPositionKind.Subscript;
                default:
                    return null;
            }
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

        private static LegacyDocHighlightColorKind? MapHighlight(byte value) {
            switch (value) {
                case 1:
                    return LegacyDocHighlightColorKind.Black;
                case 2:
                    return LegacyDocHighlightColorKind.Blue;
                case 3:
                    return LegacyDocHighlightColorKind.Cyan;
                case 4:
                    return LegacyDocHighlightColorKind.Green;
                case 5:
                    return LegacyDocHighlightColorKind.Magenta;
                case 6:
                    return LegacyDocHighlightColorKind.Red;
                case 7:
                    return LegacyDocHighlightColorKind.Yellow;
                case 8:
                    return LegacyDocHighlightColorKind.White;
                case 9:
                    return LegacyDocHighlightColorKind.DarkBlue;
                case 10:
                    return LegacyDocHighlightColorKind.DarkCyan;
                case 11:
                    return LegacyDocHighlightColorKind.DarkGreen;
                case 12:
                    return LegacyDocHighlightColorKind.DarkMagenta;
                case 13:
                    return LegacyDocHighlightColorKind.DarkRed;
                case 14:
                    return LegacyDocHighlightColorKind.DarkYellow;
                case 15:
                    return LegacyDocHighlightColorKind.DarkGray;
                case 16:
                    return LegacyDocHighlightColorKind.LightGray;
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
