using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static void AppendSupportedRunText(StringBuilder text, List<LegacyDocWritableRun> runs, Run run) {
            LegacyDocWritableFormatting formatting = ReadSupportedRunFormatting(run.RunProperties);

            foreach (OpenXmlElement child in run.ChildElements) {
                switch (child) {
                    case RunProperties:
                        break;
                    case Text textNode:
                        AppendFormattedText(text, runs, textNode.Text, formatting);
                        break;
                    case TabChar:
                        AppendFormattedText(text, runs, "\t", formatting);
                        break;
                    case Break breakNode:
                        AppendSupportedBreak(text, runs, breakNode, formatting);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving currently supports text, tabs, line breaks, and page breaks only. Unsupported run element: {child.LocalName}.");
                }
            }
        }

        private static void AppendSupportedBreak(StringBuilder text, List<LegacyDocWritableRun> runs, Break breakNode, LegacyDocWritableFormatting formatting) {
            BreakValues? breakType = breakNode.Type?.Value;
            if (breakType == null || breakType == BreakValues.TextWrapping) {
                AppendFormattedText(text, runs, "\v", formatting);
                return;
            }

            if (breakType == BreakValues.Page) {
                AppendFormattedText(text, runs, "\f", formatting);
                return;
            }

            throw new NotSupportedException($"Native DOC saving currently supports text-wrapping and page breaks only. Unsupported break type: {breakType}.");
        }

        private static LegacyDocWritableFormatting ReadSupportedRunFormatting(RunProperties? runProperties) {
            if (runProperties == null || !runProperties.HasChildren) {
                return LegacyDocWritableFormatting.Plain;
            }

            bool bold = false;
            bool italic = false;
            bool strike = false;
            byte? caps = null;
            byte? verticalPosition = null;
            byte? underline = null;
            byte? highlight = null;
            int? fontSizeHalfPoints = null;
            string? colorHex = null;
            string? fontFamily = null;
            foreach (OpenXmlElement property in runProperties.ChildElements) {
                switch (property) {
                    case Bold boldProperty:
                        bold = IsEnabled(boldProperty);
                        break;
                    case BoldComplexScript boldComplexScript:
                        bold = IsEnabled(boldComplexScript);
                        break;
                    case Italic italicProperty:
                        italic = IsEnabled(italicProperty);
                        break;
                    case ItalicComplexScript italicComplexScript:
                        italic = IsEnabled(italicComplexScript);
                        break;
                    case Strike strikeProperty:
                        strike = IsEnabled(strikeProperty);
                        break;
                    case Caps capsProperty:
                        caps = IsEnabled(capsProperty) ? (byte)1 : null;
                        break;
                    case SmallCaps smallCapsProperty:
                        caps = IsEnabled(smallCapsProperty) ? (byte)2 : null;
                        break;
                    case VerticalTextAlignment verticalTextAlignment:
                        verticalPosition = ReadSupportedVerticalPosition(verticalTextAlignment);
                        break;
                    case Underline underlineProperty:
                        underline = ReadSupportedUnderline(underlineProperty);
                        break;
                    case Highlight highlightProperty:
                        highlight = ReadSupportedHighlight(highlightProperty);
                        break;
                    case FontSize fontSize:
                        fontSizeHalfPoints = ReadFontSizeHalfPoints(fontSize);
                        break;
                    case Color color:
                        colorHex = ReadSupportedColorHex(color);
                        break;
                    case RunFonts runFonts:
                        fontFamily = ReadSupportedRunFontFamily(runFonts);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving currently supports only bold, italic, strikethrough, caps/small-caps, superscript/subscript, underline, highlight, font size, color, and font family run formatting. Unsupported run property: {property.LocalName}.");
                }
            }

            return new LegacyDocWritableFormatting(bold, italic, strike, caps, verticalPosition, underline, highlight, fontSizeHalfPoints, colorHex, fontFamily);
        }

        private static bool IsEnabled(OnOffType property) {
            return property.Val == null || property.Val.Value;
        }

        private static byte? ReadSupportedUnderline(Underline underline) {
            UnderlineValues value = underline.Val?.Value ?? UnderlineValues.Single;
            if (value == UnderlineValues.None) {
                return null;
            } else if (value == UnderlineValues.Single) {
                return 1;
            } else if (value == UnderlineValues.Words) {
                return 2;
            } else if (value == UnderlineValues.Double) {
                return 3;
            } else if (value == UnderlineValues.Dotted) {
                return 4;
            } else if (value == UnderlineValues.Thick) {
                return 6;
            } else if (value == UnderlineValues.Dash) {
                return 7;
            } else if (value == UnderlineValues.DotDash) {
                return 8;
            } else if (value == UnderlineValues.DotDotDash) {
                return 9;
            } else if (value == UnderlineValues.Wave) {
                return 10;
            } else if (value == UnderlineValues.DottedHeavy) {
                return 11;
            } else if (value == UnderlineValues.DashedHeavy) {
                return 12;
            } else if (value == UnderlineValues.DashDotHeavy) {
                return 13;
            } else if (value == UnderlineValues.DashDotDotHeavy) {
                return 14;
            } else if (value == UnderlineValues.WavyHeavy) {
                return 15;
            } else if (value == UnderlineValues.DashLong) {
                return 16;
            } else if (value == UnderlineValues.WavyDouble) {
                return 17;
            } else if (value == UnderlineValues.DashLongHeavy) {
                return 18;
            }

            throw new NotSupportedException($"Native DOC saving does not support underline style '{value}'.");
        }

        private static byte? ReadSupportedVerticalPosition(VerticalTextAlignment verticalTextAlignment) {
            VerticalPositionValues? value = verticalTextAlignment.Val?.Value;
            if (value == null) {
                return null;
            }

            if (value == VerticalPositionValues.Baseline) {
                return null;
            } else if (value == VerticalPositionValues.Superscript) {
                return 1;
            } else if (value == VerticalPositionValues.Subscript) {
                return 2;
            }

            throw new NotSupportedException($"Native DOC saving does not support vertical text alignment '{value}'.");
        }

        private static byte? ReadSupportedHighlight(Highlight highlight) {
            HighlightColorValues? value = highlight.Val?.Value;
            if (value == null || value == HighlightColorValues.None) {
                return null;
            }

            if (value == HighlightColorValues.Black) return 1;
            if (value == HighlightColorValues.Blue) return 2;
            if (value == HighlightColorValues.Cyan) return 3;
            if (value == HighlightColorValues.Green) return 4;
            if (value == HighlightColorValues.Magenta) return 5;
            if (value == HighlightColorValues.Red) return 6;
            if (value == HighlightColorValues.Yellow) return 7;
            if (value == HighlightColorValues.White) return 8;
            if (value == HighlightColorValues.DarkBlue) return 9;
            if (value == HighlightColorValues.DarkCyan) return 10;
            if (value == HighlightColorValues.DarkGreen) return 11;
            if (value == HighlightColorValues.DarkMagenta) return 12;
            if (value == HighlightColorValues.DarkRed) return 13;
            if (value == HighlightColorValues.DarkYellow) return 14;
            if (value == HighlightColorValues.DarkGray) return 15;
            if (value == HighlightColorValues.LightGray) return 16;

            throw new NotSupportedException($"Native DOC saving does not support highlight color '{value}'.");
        }

        private static int ReadFontSizeHalfPoints(FontSize fontSize) {
            string? value = fontSize.Val?.Value;
            if (string.IsNullOrWhiteSpace(value) || !int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int halfPoints)) {
                throw new NotSupportedException("Native DOC saving supports font size only when it is stored as a numeric half-point value.");
            }

            return halfPoints;
        }

        private static string? ReadSupportedColorHex(Color color) {
            string? value = color.Val?.Value;
            if (string.IsNullOrWhiteSpace(value) || string.Equals(value, "auto", StringComparison.OrdinalIgnoreCase)) {
                return null;
            }

            string colorValue = value!;
            string hex = colorValue.Trim().TrimStart('#').ToLowerInvariant();
            if (hex.Length != 6 || hex.Any(character => !Uri.IsHexDigit(character))) {
                throw new NotSupportedException("Native DOC saving supports text color only when it is stored as a 6-digit RGB hex value.");
            }

            return hex;
        }

        private static string? ReadSupportedRunFontFamily(RunFonts runFonts) {
            string? ascii = NormalizeFontFamily(runFonts.Ascii?.Value);
            string? highAnsi = NormalizeFontFamily(runFonts.HighAnsi?.Value);
            string? eastAsia = NormalizeFontFamily(runFonts.EastAsia?.Value);
            string? complexScript = NormalizeFontFamily(runFonts.ComplexScript?.Value);

            string? fontFamily = ascii ?? highAnsi;
            if (fontFamily == null) {
                if (eastAsia != null || complexScript != null) {
                    throw new NotSupportedException("Native DOC saving currently supports font family only for ASCII/HighAnsi text runs.");
                }

                return null;
            }

            if ((highAnsi != null && !string.Equals(fontFamily, highAnsi, StringComparison.OrdinalIgnoreCase))
                || (eastAsia != null && !string.Equals(fontFamily, eastAsia, StringComparison.OrdinalIgnoreCase))
                || (complexScript != null && !string.Equals(fontFamily, complexScript, StringComparison.OrdinalIgnoreCase))) {
                throw new NotSupportedException("Native DOC saving currently supports a single font family per text run. Multiple script-specific font families are not supported yet.");
            }

            return fontFamily;
        }

        private static string? NormalizeFontFamily(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            return value!.Trim();
        }

        private static void AppendFormattedText(
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            string? value,
            LegacyDocWritableFormatting formatting) {
            if (string.IsNullOrEmpty(value)) {
                return;
            }

            string textValue = value!;
            int start = text.Length;
            text.Append(textValue);
            if (!formatting.HasFormatting) {
                return;
            }

            int length = textValue.Length;
            if (runs.Count > 0) {
                LegacyDocWritableRun previous = runs[runs.Count - 1];
                if (previous.EndCharacter == start && previous.Formatting.Equals(formatting)) {
                    runs[runs.Count - 1] = previous.Extend(length);
                    return;
                }
            }

            runs.Add(new LegacyDocWritableRun(start, length, formatting));
        }

        private static void WriteChpxFkp(byte[] stream, int pageOffset, IReadOnlyList<LegacyDocWritableSegment> segments, IReadOnlyDictionary<string, int> fontFamilyIndexes, int bytesPerCharacter) {
            if (segments.Count == 0 || segments.Count > byte.MaxValue) {
                throw new NotSupportedException("Native DOC saving currently supports run formatting only when it fits in one character-format page.");
            }

            int rgbOffset = pageOffset + ((segments.Count + 1) * 4);
            int chpxOffset = AlignToEven((segments.Count + 1) * 4 + segments.Count);

            for (int index = 0; index < segments.Count; index++) {
                LegacyDocWritableSegment segment = segments[index];
                WriteInt32(stream, pageOffset + (index * 4), TextOffset + (segment.StartCharacter * bytesPerCharacter));
                if (segment.Formatting.HasFormatting) {
                    byte[] chpx = CreateChpx(segment.Formatting, fontFamilyIndexes);
                    chpxOffset = AlignToEven(chpxOffset);
                    if (chpxOffset + chpx.Length >= OleSectorSize - 1 || chpxOffset / 2 > byte.MaxValue) {
                        throw new NotSupportedException("Native DOC saving currently supports run formatting only when it fits in one character-format page.");
                    }

                    Buffer.BlockCopy(chpx, 0, stream, pageOffset + chpxOffset, chpx.Length);
                    stream[rgbOffset + index] = (byte)(chpxOffset / 2);
                    chpxOffset += chpx.Length;
                }
            }

            LegacyDocWritableSegment lastSegment = segments[segments.Count - 1];
            WriteInt32(stream, pageOffset + (segments.Count * 4), TextOffset + (lastSegment.EndCharacter * bytesPerCharacter));
            stream[pageOffset + OleSectorSize - 1] = (byte)segments.Count;
        }

        private static byte[] CreateChpx(LegacyDocWritableFormatting formatting, IReadOnlyDictionary<string, int> fontFamilyIndexes) {
            var grpprl = new List<byte>(18);
            if (formatting.Bold) {
                AddSingleByteSprm(grpprl, SprmCFBold, 1);
            }

            if (formatting.Italic) {
                AddSingleByteSprm(grpprl, SprmCFItalic, 1);
            }

            if (formatting.Strike) {
                AddSingleByteSprm(grpprl, SprmCFStrike, 1);
            }

            if (formatting.Caps == 1) {
                AddSingleByteSprm(grpprl, SprmCFCaps, 1);
            } else if (formatting.Caps == 2) {
                AddSingleByteSprm(grpprl, SprmCFSmallCaps, 1);
            }

            if (formatting.VerticalPosition != null) {
                AddSingleByteSprm(grpprl, SprmCIss, formatting.VerticalPosition.Value);
            }

            if (formatting.Underline != null) {
                AddSingleByteSprm(grpprl, SprmCKul, formatting.Underline.Value);
            }

            if (formatting.Highlight != null) {
                AddSingleByteSprm(grpprl, SprmCHighlight, formatting.Highlight.Value);
            }

            if (formatting.FontSizeHalfPoints != null) {
                AddUInt16Sprm(grpprl, SprmCHps, checked((ushort)formatting.FontSizeHalfPoints.Value));
            }

            if (formatting.ColorHex != null) {
                AddColorRefSprm(grpprl, formatting.ColorHex);
            }

            if (formatting.FontFamily != null) {
                if (!fontFamilyIndexes.TryGetValue(formatting.FontFamily, out int fontIndex)) {
                    throw new InvalidOperationException("The DOC font table does not contain a formatted run font family.");
                }

                AddUInt16Sprm(grpprl, SprmCRgFtc0, checked((ushort)fontIndex));
            }

            var chpx = new byte[grpprl.Count + 1];
            chpx[0] = (byte)grpprl.Count;
            grpprl.CopyTo(chpx, 1);
            return chpx;
        }

        private static byte[] CreateFontTable(IReadOnlyList<string> fontFamilies) {
            if (fontFamilies.Count == 0) {
                return Array.Empty<byte>();
            }

            if (fontFamilies.Count > ushort.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports only documents whose font table fits in a Word 97-2003 STTBF.");
            }

            using var stream = new MemoryStream();
            WriteUInt16(stream, checked((ushort)fontFamilies.Count));
            WriteUInt16(stream, 0);

            foreach (string fontFamily in fontFamilies) {
                byte[] ffn = CreateFfn(fontFamily);
                if (ffn.Length > byte.MaxValue) {
                    throw new NotSupportedException($"Native DOC saving cannot write font family '{fontFamily}' because its DOC font-table record is too long.");
                }

                stream.WriteByte(checked((byte)ffn.Length));
                stream.Write(ffn, 0, ffn.Length);
            }

            return stream.ToArray();
        }

        private static byte[] CreateFfn(string fontFamily) {
            if (string.IsNullOrWhiteSpace(fontFamily)) {
                throw new NotSupportedException("Native DOC saving cannot write an empty font family name.");
            }

            byte[] nameBytes = Encoding.Unicode.GetBytes(fontFamily + '\0');
            var ffn = new byte[39 + nameBytes.Length];
            ffn[1] = 0x90;
            ffn[2] = 0x01;
            Buffer.BlockCopy(nameBytes, 0, ffn, 39, nameBytes.Length);
            return ffn;
        }

        private static void AddSingleByteSprm(List<byte> grpprl, ushort sprm, byte operand) {
            grpprl.Add((byte)(sprm & 0xFF));
            grpprl.Add((byte)(sprm >> 8));
            grpprl.Add(operand);
        }

        private static void AddUInt16Sprm(List<byte> grpprl, ushort sprm, ushort operand) {
            grpprl.Add((byte)(sprm & 0xFF));
            grpprl.Add((byte)(sprm >> 8));
            grpprl.Add((byte)(operand & 0xFF));
            grpprl.Add((byte)(operand >> 8));
        }

        private static void AddColorRefSprm(List<byte> grpprl, string colorHex) {
            grpprl.Add((byte)(SprmCCv & 0xFF));
            grpprl.Add((byte)(SprmCCv >> 8));
            grpprl.Add(Convert.ToByte(colorHex.Substring(0, 2), 16));
            grpprl.Add(Convert.ToByte(colorHex.Substring(2, 2), 16));
            grpprl.Add(Convert.ToByte(colorHex.Substring(4, 2), 16));
            grpprl.Add(0);
        }

        private readonly struct LegacyDocWritableFormatting : IEquatable<LegacyDocWritableFormatting> {
            internal static readonly LegacyDocWritableFormatting Plain = new LegacyDocWritableFormatting(false, false, false, null, null, null, null, null, null, null);

            internal LegacyDocWritableFormatting(bool bold, bool italic, bool strike, byte? caps, byte? verticalPosition, byte? underline, byte? highlight, int? fontSizeHalfPoints, string? colorHex, string? fontFamily) {
                Bold = bold;
                Italic = italic;
                Strike = strike;
                Caps = caps;
                VerticalPosition = verticalPosition;
                Underline = underline;
                Highlight = highlight;
                FontSizeHalfPoints = fontSizeHalfPoints;
                ColorHex = colorHex;
                FontFamily = fontFamily;
            }

            internal bool Bold { get; }

            internal bool Italic { get; }

            internal bool Strike { get; }

            internal byte? Caps { get; }

            internal byte? VerticalPosition { get; }

            internal byte? Underline { get; }

            internal byte? Highlight { get; }

            internal int? FontSizeHalfPoints { get; }

            internal string? ColorHex { get; }

            internal string? FontFamily { get; }

            internal bool HasFormatting => Bold || Italic || Strike || Caps != null || VerticalPosition != null || Underline != null || Highlight != null || FontSizeHalfPoints != null || ColorHex != null || FontFamily != null;

            public bool Equals(LegacyDocWritableFormatting other) {
                return Bold == other.Bold
                    && Italic == other.Italic
                    && Strike == other.Strike
                    && Caps == other.Caps
                    && VerticalPosition == other.VerticalPosition
                    && Underline == other.Underline
                    && Highlight == other.Highlight
                    && FontSizeHalfPoints == other.FontSizeHalfPoints
                    && string.Equals(ColorHex, other.ColorHex, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(FontFamily, other.FontFamily, StringComparison.OrdinalIgnoreCase);
            }

            public override bool Equals(object? obj) {
                return obj is LegacyDocWritableFormatting other && Equals(other);
            }

            public override int GetHashCode() {
                int hash = 17;
                hash = (hash * 31) + Bold.GetHashCode();
                hash = (hash * 31) + Italic.GetHashCode();
                hash = (hash * 31) + Strike.GetHashCode();
                hash = (hash * 31) + Caps.GetHashCode();
                hash = (hash * 31) + VerticalPosition.GetHashCode();
                hash = (hash * 31) + Underline.GetHashCode();
                hash = (hash * 31) + Highlight.GetHashCode();
                hash = (hash * 31) + FontSizeHalfPoints.GetHashCode();
                hash = (hash * 31) + StringComparer.OrdinalIgnoreCase.GetHashCode(ColorHex ?? string.Empty);
                hash = (hash * 31) + StringComparer.OrdinalIgnoreCase.GetHashCode(FontFamily ?? string.Empty);
                return hash;
            }
        }

        private readonly struct LegacyDocWritableRun {
            internal LegacyDocWritableRun(int startCharacter, int length, LegacyDocWritableFormatting formatting) {
                StartCharacter = startCharacter;
                Length = length;
                Formatting = formatting;
            }

            internal int StartCharacter { get; }

            internal int Length { get; }

            internal int EndCharacter => StartCharacter + Length;

            internal LegacyDocWritableFormatting Formatting { get; }

            internal LegacyDocWritableRun Extend(int additionalLength) {
                return new LegacyDocWritableRun(StartCharacter, Length + additionalLength, Formatting);
            }
        }

        private readonly struct LegacyDocWritableSegment {
            internal LegacyDocWritableSegment(int startCharacter, int length, LegacyDocWritableFormatting formatting) {
                StartCharacter = startCharacter;
                Length = length;
                Formatting = formatting;
            }

            internal int StartCharacter { get; }

            internal int Length { get; }

            internal int EndCharacter => StartCharacter + Length;

            internal LegacyDocWritableFormatting Formatting { get; }

            internal LegacyDocWritableSegment Extend(int additionalLength) {
                return new LegacyDocWritableSegment(StartCharacter, Length + additionalLength, Formatting);
            }
        }
    }
}
