using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static void AppendSupportedRunText(StringBuilder text, List<LegacyDocWritableRun> runs, Run run, LegacyDocWritableFootnotes footnotes, LegacyDocWritableEndnotes endnotes) {
            AppendSupportedRunText(text, runs, run, footnotes, endnotes, LegacyDocWritableFormatting.Plain);
        }

        private static void AppendSupportedRunText(StringBuilder text, List<LegacyDocWritableRun> runs, Run run, LegacyDocWritableFootnotes footnotes, LegacyDocWritableEndnotes endnotes, LegacyDocWritableFormatting inheritedFormatting) {
            AppendSupportedRunText(text, runs, run, footnotes, endnotes, inheritedFormatting, allowHyperlinkRunStyle: false);
        }

        private static void AppendSupportedRunText(StringBuilder text, List<LegacyDocWritableRun> runs, Run run, LegacyDocWritableFootnotes footnotes, LegacyDocWritableEndnotes endnotes, LegacyDocWritableFormatting inheritedFormatting, bool allowHyperlinkRunStyle) {
            if (run.Elements<FootnoteReference>().Any()) {
                AppendFootnoteReferenceRun(text, runs, footnotes, run);
                return;
            }

            if (run.Elements<EndnoteReference>().Any()) {
                AppendEndnoteReferenceRun(text, runs, endnotes, run);
                return;
            }

            LegacyDocWritableFormatting formatting = ReadSupportedRunFormatting(run.RunProperties, allowHyperlinkRunStyle)
                .WithInheritedFormatting(inheritedFormatting);

            foreach (OpenXmlElement child in run.ChildElements) {
                switch (child) {
                    case RunProperties:
                        break;
                    case LastRenderedPageBreak:
                        break;
                    case DocumentFormat.OpenXml.Wordprocessing.PageNumber:
                        AppendSupportedPageNumberField(text, runs, formatting);
                        break;
                    case Text textNode:
                        AppendFormattedText(text, runs, textNode.Text, formatting);
                        break;
                    case TabChar:
                        AppendFormattedText(text, runs, "\t", formatting);
                        break;
                    case CarriageReturn:
                        AppendFormattedText(text, runs, LegacyDocSpecialCharacters.TextWrappingBreak.ToString(), formatting);
                        break;
                    case NoBreakHyphen:
                        AppendFormattedText(text, runs, LegacyDocSpecialCharacters.NoBreakHyphen.ToString(), formatting);
                        break;
                    case SoftHyphen:
                        AppendFormattedText(text, runs, LegacyDocSpecialCharacters.SoftHyphen.ToString(), formatting);
                        break;
                    case Break breakNode:
                        AppendSupportedBreak(text, runs, breakNode, formatting);
                        break;
                    case FootnoteReference footnoteReference:
                        AppendFootnoteReference(text, runs, footnotes, footnoteReference);
                        break;
                    case EndnoteReference endnoteReference:
                        AppendEndnoteReference(text, runs, endnotes, endnoteReference);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving currently supports text, tabs, page-number fields, carriage returns, soft/no-break hyphens, text-wrapping/page/column breaks, and simple footnote/endnote references only. Unsupported run element: {child.LocalName}.");
                }
            }
        }

        private static void AppendFootnoteReferenceRun(StringBuilder text, List<LegacyDocWritableRun> runs, LegacyDocWritableFootnotes footnotes, Run run) {
            foreach (OpenXmlElement child in run.ChildElements) {
                switch (child) {
                    case RunProperties:
                        break;
                    case LastRenderedPageBreak:
                        break;
                    case FootnoteReference footnoteReference:
                        AppendFootnoteReference(text, runs, footnotes, footnoteReference);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports footnote reference runs only when they contain footnote references. Unsupported footnote reference run element: {child.LocalName}.");
                }
            }
        }

        private static void AppendEndnoteReferenceRun(StringBuilder text, List<LegacyDocWritableRun> runs, LegacyDocWritableEndnotes endnotes, Run run) {
            foreach (OpenXmlElement child in run.ChildElements) {
                switch (child) {
                    case RunProperties:
                        break;
                    case LastRenderedPageBreak:
                        break;
                    case EndnoteReference endnoteReference:
                        AppendEndnoteReference(text, runs, endnotes, endnoteReference);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports endnote reference runs only when they contain endnote references. Unsupported endnote reference run element: {child.LocalName}.");
                }
            }
        }

        private static void AppendFootnoteReference(StringBuilder text, List<LegacyDocWritableRun> runs, LegacyDocWritableFootnotes footnotes, FootnoteReference footnoteReference) {
            long? id = footnoteReference.Id?.Value;
            if (id == null || id.Value <= 0) {
                throw new NotSupportedException("Native DOC saving supports footnote references only when they use a positive identifier.");
            }

            int referencePosition = text.Length;
            footnotes.AddReference(id.Value, referencePosition);
            AppendFormattedText(text, runs, LegacyDocFootnoteReader.FootnoteReferenceCharacter.ToString(), LegacyDocWritableFormatting.SpecialCharacter);
        }

        private static void AppendEndnoteReference(StringBuilder text, List<LegacyDocWritableRun> runs, LegacyDocWritableEndnotes endnotes, EndnoteReference endnoteReference) {
            long? id = endnoteReference.Id?.Value;
            if (id == null || id.Value <= 0) {
                throw new NotSupportedException("Native DOC saving supports endnote references only when they use a positive identifier.");
            }

            int referencePosition = text.Length;
            endnotes.AddReference(id.Value, referencePosition);
            AppendFormattedText(text, runs, LegacyDocFootnoteReader.FootnoteReferenceCharacter.ToString(), LegacyDocWritableFormatting.SpecialCharacter);
        }

        private static void AppendSupportedBreak(StringBuilder text, List<LegacyDocWritableRun> runs, Break breakNode, LegacyDocWritableFormatting formatting) {
            BreakValues? breakType = breakNode.Type?.Value;
            if (breakType == null || breakType == BreakValues.TextWrapping) {
                AppendFormattedText(text, runs, LegacyDocSpecialCharacters.TextWrappingBreak.ToString(), formatting);
                return;
            }

            if (breakType == BreakValues.Page) {
                AppendFormattedText(text, runs, LegacyDocSpecialCharacters.PageBreak.ToString(), formatting);
                return;
            }

            if (breakType == BreakValues.Column) {
                AppendFormattedText(text, runs, LegacyDocSpecialCharacters.ColumnBreak.ToString(), formatting);
                return;
            }

            throw new NotSupportedException($"Native DOC saving currently supports text-wrapping, page, and column breaks only. Unsupported break type: {breakType}.");
        }

        private static LegacyDocWritableFormatting ReadSupportedRunFormatting(OpenXmlCompositeElement? runProperties) {
            return ReadSupportedRunFormatting(runProperties, allowHyperlinkRunStyle: false);
        }

        private static LegacyDocWritableFormatting ReadSupportedRunFormatting(OpenXmlCompositeElement? runProperties, bool allowHyperlinkRunStyle) {
            if (runProperties == null || !runProperties.HasChildren) {
                return LegacyDocWritableFormatting.Plain;
            }

            bool? bold = null;
            bool? italic = null;
            bool strike = false;
            bool doubleStrike = false;
            bool outline = false;
            bool shadow = false;
            bool emboss = false;
            bool imprint = false;
            bool hidden = false;
            bool noProof = false;
            byte? caps = null;
            byte? verticalPosition = null;
            byte? underline = null;
            byte? highlight = null;
            int? fontSizeHalfPoints = null;
            string? colorHex = null;
            string? fontFamily = null;
            LegacyDocWritableFormattingProperties specified = LegacyDocWritableFormattingProperties.None;
            foreach (OpenXmlElement property in runProperties.ChildElements) {
                switch (property) {
                    case Bold boldProperty:
                        specified |= LegacyDocWritableFormattingProperties.Bold;
                        bold = MergeSingleRunToggle(bold, IsEnabled(boldProperty), "bold", "Bold", "BoldComplexScript");
                        break;
                    case BoldComplexScript boldComplexScript:
                        specified |= LegacyDocWritableFormattingProperties.Bold;
                        bold = MergeSingleRunToggle(bold, IsEnabled(boldComplexScript), "bold", "Bold", "BoldComplexScript");
                        break;
                    case Italic italicProperty:
                        specified |= LegacyDocWritableFormattingProperties.Italic;
                        italic = MergeSingleRunToggle(italic, IsEnabled(italicProperty), "italic", "Italic", "ItalicComplexScript");
                        break;
                    case ItalicComplexScript italicComplexScript:
                        specified |= LegacyDocWritableFormattingProperties.Italic;
                        italic = MergeSingleRunToggle(italic, IsEnabled(italicComplexScript), "italic", "Italic", "ItalicComplexScript");
                        break;
                    case Strike strikeProperty:
                        specified |= LegacyDocWritableFormattingProperties.Strike;
                        strike = IsEnabled(strikeProperty);
                        break;
                    case DoubleStrike doubleStrikeProperty:
                        specified |= LegacyDocWritableFormattingProperties.DoubleStrike;
                        doubleStrike = IsEnabled(doubleStrikeProperty);
                        break;
                    case Outline outlineProperty:
                        specified |= LegacyDocWritableFormattingProperties.Outline;
                        outline = IsEnabled(outlineProperty);
                        break;
                    case Shadow shadowProperty:
                        specified |= LegacyDocWritableFormattingProperties.Shadow;
                        shadow = IsEnabled(shadowProperty);
                        break;
                    case Emboss embossProperty:
                        specified |= LegacyDocWritableFormattingProperties.Emboss;
                        emboss = IsEnabled(embossProperty);
                        break;
                    case Imprint imprintProperty:
                        specified |= LegacyDocWritableFormattingProperties.Imprint;
                        imprint = IsEnabled(imprintProperty);
                        break;
                    case Vanish vanishProperty:
                        specified |= LegacyDocWritableFormattingProperties.Hidden;
                        hidden = IsEnabled(vanishProperty);
                        break;
                    case NoProof noProofProperty:
                        specified |= LegacyDocWritableFormattingProperties.NoProof;
                        noProof = IsEnabled(noProofProperty);
                        break;
                    case Caps capsProperty:
                        specified |= LegacyDocWritableFormattingProperties.Caps;
                        caps = MergeCapsKind(caps, IsEnabled(capsProperty), 1);
                        break;
                    case SmallCaps smallCapsProperty:
                        specified |= LegacyDocWritableFormattingProperties.Caps;
                        caps = MergeCapsKind(caps, IsEnabled(smallCapsProperty), 2);
                        break;
                    case VerticalTextAlignment verticalTextAlignment:
                        specified |= LegacyDocWritableFormattingProperties.VerticalPosition;
                        verticalPosition = ReadSupportedVerticalPosition(verticalTextAlignment);
                        break;
                    case Underline underlineProperty:
                        specified |= LegacyDocWritableFormattingProperties.Underline;
                        underline = ReadSupportedUnderline(underlineProperty);
                        break;
                    case Highlight highlightProperty:
                        specified |= LegacyDocWritableFormattingProperties.Highlight;
                        highlight = ReadSupportedHighlight(highlightProperty);
                        break;
                    case FontSize fontSize:
                        specified |= LegacyDocWritableFormattingProperties.FontSize;
                        fontSizeHalfPoints = MergeFontSizeHalfPoints(fontSizeHalfPoints, ReadFontSizeHalfPoints(fontSize.Val?.Value));
                        break;
                    case FontSizeComplexScript fontSizeComplexScript:
                        specified |= LegacyDocWritableFormattingProperties.FontSize;
                        fontSizeHalfPoints = MergeFontSizeHalfPoints(fontSizeHalfPoints, ReadFontSizeHalfPoints(fontSizeComplexScript.Val?.Value));
                        break;
                    case Color color:
                        specified |= LegacyDocWritableFormattingProperties.Color;
                        colorHex = ReadSupportedColorHex(color);
                        break;
                    case RunFonts runFonts:
                        specified |= LegacyDocWritableFormattingProperties.FontFamily;
                        fontFamily = ReadSupportedRunFontFamily(runFonts);
                        break;
                    case RunStyle runStyle when allowHyperlinkRunStyle && string.Equals(runStyle.Val?.Value, "Hyperlink", StringComparison.OrdinalIgnoreCase):
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving currently supports only bold, italic, strikethrough, double-strikethrough, outline, shadow, emboss, imprint, hidden text, proofing exclusion, caps/small-caps, superscript/subscript, underline, highlight, font size, color, and font family run formatting. Unsupported run property: {property.LocalName}.");
                }
            }

            return new LegacyDocWritableFormatting(bold == true, italic == true, strike, doubleStrike, outline, shadow, emboss, imprint, hidden, noProof, false, caps, verticalPosition, underline, highlight, fontSizeHalfPoints, colorHex, fontFamily, specified);
        }

        private static bool IsEnabled(OnOffType property) {
            return property.Val == null || property.Val.Value;
        }

        private static bool MergeSingleRunToggle(bool? currentValue, bool nextValue, string description, string directPropertyName, string complexScriptPropertyName) {
            if (currentValue != null && currentValue.Value != nextValue) {
                throw new NotSupportedException($"Native DOC saving supports one {description} value per text run. {directPropertyName} and {complexScriptPropertyName} must match.");
            }

            return nextValue;
        }

        private static byte? MergeCapsKind(byte? currentKind, bool enabled, byte nextKind) {
            if (!enabled) {
                return currentKind;
            }

            if (currentKind != null && currentKind.Value != nextKind) {
                throw new NotSupportedException("Native DOC saving supports either all-caps or small-caps per text run. Caps and SmallCaps cannot both be enabled.");
            }

            return nextKind;
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

        private static int ReadFontSizeHalfPoints(string? value) {
            if (string.IsNullOrWhiteSpace(value) || !int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int halfPoints)) {
                throw new NotSupportedException("Native DOC saving supports font size only when it is stored as a numeric half-point value.");
            }

            return halfPoints;
        }

        private static int MergeFontSizeHalfPoints(int? currentHalfPoints, int nextHalfPoints) {
            if (currentHalfPoints != null && currentHalfPoints.Value != nextHalfPoints) {
                throw new NotSupportedException("Native DOC saving supports one font size per text run. FontSize and FontSizeComplexScript must match.");
            }

            return nextHalfPoints;
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

            string? fontFamily = ascii ?? highAnsi ?? eastAsia ?? complexScript;
            if (fontFamily == null) {
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
            List<byte> grpprl = CreateCharacterGrpprl(formatting, fontFamilyIndexes);
            var chpx = new byte[grpprl.Count + 1];
            chpx[0] = (byte)grpprl.Count;
            grpprl.CopyTo(chpx, 1);
            return chpx;
        }

        private static byte[] CreateStyleCharacterUpx(LegacyDocWritableFormatting formatting, IReadOnlyDictionary<string, int> fontFamilyIndexes) {
            if (!formatting.HasFormatting) {
                return Array.Empty<byte>();
            }

            return CreateCharacterGrpprl(formatting, fontFamilyIndexes).ToArray();
        }

        private static List<byte> CreateCharacterGrpprl(LegacyDocWritableFormatting formatting, IReadOnlyDictionary<string, int> fontFamilyIndexes) {
            var grpprl = new List<byte>(18);
            if (formatting.Bold || formatting.IsSpecified(LegacyDocWritableFormattingProperties.Bold)) {
                AddSingleByteSprm(grpprl, SprmCFBold, formatting.Bold ? (byte)1 : (byte)0);
            }

            if (formatting.Italic || formatting.IsSpecified(LegacyDocWritableFormattingProperties.Italic)) {
                AddSingleByteSprm(grpprl, SprmCFItalic, formatting.Italic ? (byte)1 : (byte)0);
            }

            if (formatting.Strike || formatting.IsSpecified(LegacyDocWritableFormattingProperties.Strike)) {
                AddSingleByteSprm(grpprl, SprmCFStrike, formatting.Strike ? (byte)1 : (byte)0);
            }

            if (formatting.DoubleStrike || formatting.IsSpecified(LegacyDocWritableFormattingProperties.DoubleStrike)) {
                AddSingleByteSprm(grpprl, SprmCFDStrike, formatting.DoubleStrike ? (byte)1 : (byte)0);
            }

            if (formatting.Outline || formatting.IsSpecified(LegacyDocWritableFormattingProperties.Outline)) {
                AddSingleByteSprm(grpprl, SprmCFOutline, formatting.Outline ? (byte)1 : (byte)0);
            }

            if (formatting.Shadow || formatting.IsSpecified(LegacyDocWritableFormattingProperties.Shadow)) {
                AddSingleByteSprm(grpprl, SprmCFShadow, formatting.Shadow ? (byte)1 : (byte)0);
            }

            if (formatting.Emboss || formatting.IsSpecified(LegacyDocWritableFormattingProperties.Emboss)) {
                AddSingleByteSprm(grpprl, SprmCFEmboss, formatting.Emboss ? (byte)1 : (byte)0);
            }

            if (formatting.Imprint || formatting.IsSpecified(LegacyDocWritableFormattingProperties.Imprint)) {
                AddSingleByteSprm(grpprl, SprmCFImprint, formatting.Imprint ? (byte)1 : (byte)0);
            }

            if (formatting.Hidden || formatting.IsSpecified(LegacyDocWritableFormattingProperties.Hidden)) {
                AddSingleByteSprm(grpprl, SprmCFVanish, formatting.Hidden ? (byte)1 : (byte)0);
            }

            if (formatting.NoProof || formatting.IsSpecified(LegacyDocWritableFormattingProperties.NoProof)) {
                AddSingleByteSprm(grpprl, SprmCFNoProof, formatting.NoProof ? (byte)1 : (byte)0);
            }

            if (formatting.Special) {
                AddSingleByteSprm(grpprl, SprmCFSpec, 1);
            }

            if (formatting.Caps == 1) {
                AddSingleByteSprm(grpprl, SprmCFCaps, 1);
            } else if (formatting.Caps == 2) {
                AddSingleByteSprm(grpprl, SprmCFSmallCaps, 1);
            } else if (formatting.IsSpecified(LegacyDocWritableFormattingProperties.Caps)) {
                AddSingleByteSprm(grpprl, SprmCFCaps, 0);
                AddSingleByteSprm(grpprl, SprmCFSmallCaps, 0);
            }

            if (formatting.VerticalPosition != null || formatting.IsSpecified(LegacyDocWritableFormattingProperties.VerticalPosition)) {
                AddSingleByteSprm(grpprl, SprmCIss, formatting.VerticalPosition ?? 0);
            }

            if (formatting.Underline != null || formatting.IsSpecified(LegacyDocWritableFormattingProperties.Underline)) {
                AddSingleByteSprm(grpprl, SprmCKul, formatting.Underline ?? 0);
            }

            if (formatting.Highlight != null || formatting.IsSpecified(LegacyDocWritableFormattingProperties.Highlight)) {
                AddSingleByteSprm(grpprl, SprmCHighlight, formatting.Highlight ?? 0);
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

            return grpprl;
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

        [Flags]
        private enum LegacyDocWritableFormattingProperties {
            None = 0,
            Bold = 1 << 0,
            Italic = 1 << 1,
            Strike = 1 << 2,
            DoubleStrike = 1 << 3,
            Outline = 1 << 4,
            Shadow = 1 << 5,
            Emboss = 1 << 6,
            Imprint = 1 << 7,
            Hidden = 1 << 8,
            NoProof = 1 << 9,
            Special = 1 << 10,
            Caps = 1 << 11,
            VerticalPosition = 1 << 12,
            Underline = 1 << 13,
            Highlight = 1 << 14,
            FontSize = 1 << 15,
            Color = 1 << 16,
            FontFamily = 1 << 17
        }

        private readonly struct LegacyDocWritableFormatting : IEquatable<LegacyDocWritableFormatting> {
            internal static readonly LegacyDocWritableFormatting Plain = new LegacyDocWritableFormatting(false, false, false, false, false, false, false, false, false, false, false, null, null, null, null, null, null, null);
            internal static readonly LegacyDocWritableFormatting SpecialCharacter = new LegacyDocWritableFormatting(false, false, false, false, false, false, false, false, false, false, true, null, null, null, null, null, null, null, LegacyDocWritableFormattingProperties.Special);

            internal LegacyDocWritableFormatting(bool bold, bool italic, bool strike, bool doubleStrike, bool outline, bool shadow, bool emboss, bool imprint, bool hidden, bool noProof, bool special, byte? caps, byte? verticalPosition, byte? underline, byte? highlight, int? fontSizeHalfPoints, string? colorHex, string? fontFamily, LegacyDocWritableFormattingProperties specified = LegacyDocWritableFormattingProperties.None) {
                Bold = bold;
                Italic = italic;
                Strike = strike;
                DoubleStrike = doubleStrike;
                Outline = outline;
                Shadow = shadow;
                Emboss = emboss;
                Imprint = imprint;
                Hidden = hidden;
                NoProof = noProof;
                Special = special;
                Caps = caps;
                VerticalPosition = verticalPosition;
                Underline = underline;
                Highlight = highlight;
                FontSizeHalfPoints = fontSizeHalfPoints;
                ColorHex = colorHex;
                FontFamily = fontFamily;
                Specified = specified;
            }

            internal bool Bold { get; }

            internal bool Italic { get; }

            internal bool Strike { get; }

            internal bool DoubleStrike { get; }

            internal bool Outline { get; }

            internal bool Shadow { get; }

            internal bool Emboss { get; }

            internal bool Imprint { get; }

            internal bool Hidden { get; }

            internal bool NoProof { get; }

            internal bool Special { get; }

            internal byte? Caps { get; }

            internal byte? VerticalPosition { get; }

            internal byte? Underline { get; }

            internal byte? Highlight { get; }

            internal int? FontSizeHalfPoints { get; }

            internal string? ColorHex { get; }

            internal string? FontFamily { get; }

            internal bool HasFormatting => Bold || Italic || Strike || DoubleStrike || Outline || Shadow || Emboss || Imprint || Hidden || NoProof || Special || Caps != null || VerticalPosition != null || Underline != null || Highlight != null || FontSizeHalfPoints != null || ColorHex != null || FontFamily != null || HasExplicitOffFormatting;

            private LegacyDocWritableFormattingProperties Specified { get; }

            internal LegacyDocWritableFormatting WithInheritedFormatting(LegacyDocWritableFormatting inherited) {
                if (!inherited.HasFormatting || Special) {
                    return this;
                }

                return new LegacyDocWritableFormatting(
                    IsSpecified(LegacyDocWritableFormattingProperties.Bold) ? Bold : inherited.Bold,
                    IsSpecified(LegacyDocWritableFormattingProperties.Italic) ? Italic : inherited.Italic,
                    IsSpecified(LegacyDocWritableFormattingProperties.Strike) ? Strike : inherited.Strike,
                    IsSpecified(LegacyDocWritableFormattingProperties.DoubleStrike) ? DoubleStrike : inherited.DoubleStrike,
                    IsSpecified(LegacyDocWritableFormattingProperties.Outline) ? Outline : inherited.Outline,
                    IsSpecified(LegacyDocWritableFormattingProperties.Shadow) ? Shadow : inherited.Shadow,
                    IsSpecified(LegacyDocWritableFormattingProperties.Emboss) ? Emboss : inherited.Emboss,
                    IsSpecified(LegacyDocWritableFormattingProperties.Imprint) ? Imprint : inherited.Imprint,
                    IsSpecified(LegacyDocWritableFormattingProperties.Hidden) ? Hidden : inherited.Hidden,
                    IsSpecified(LegacyDocWritableFormattingProperties.NoProof) ? NoProof : inherited.NoProof,
                    Special,
                    IsSpecified(LegacyDocWritableFormattingProperties.Caps) ? Caps : inherited.Caps,
                    IsSpecified(LegacyDocWritableFormattingProperties.VerticalPosition) ? VerticalPosition : inherited.VerticalPosition,
                    IsSpecified(LegacyDocWritableFormattingProperties.Underline) ? Underline : inherited.Underline,
                    IsSpecified(LegacyDocWritableFormattingProperties.Highlight) ? Highlight : inherited.Highlight,
                    IsSpecified(LegacyDocWritableFormattingProperties.FontSize) ? FontSizeHalfPoints : inherited.FontSizeHalfPoints,
                    IsSpecified(LegacyDocWritableFormattingProperties.Color) ? ColorHex : inherited.ColorHex,
                    IsSpecified(LegacyDocWritableFormattingProperties.FontFamily) ? FontFamily : inherited.FontFamily,
                    Specified | inherited.Specified);
            }

            private bool HasExplicitOffFormatting =>
                (IsSpecified(LegacyDocWritableFormattingProperties.Bold) && !Bold)
                || (IsSpecified(LegacyDocWritableFormattingProperties.Italic) && !Italic)
                || (IsSpecified(LegacyDocWritableFormattingProperties.Strike) && !Strike)
                || (IsSpecified(LegacyDocWritableFormattingProperties.DoubleStrike) && !DoubleStrike)
                || (IsSpecified(LegacyDocWritableFormattingProperties.Outline) && !Outline)
                || (IsSpecified(LegacyDocWritableFormattingProperties.Shadow) && !Shadow)
                || (IsSpecified(LegacyDocWritableFormattingProperties.Emboss) && !Emboss)
                || (IsSpecified(LegacyDocWritableFormattingProperties.Imprint) && !Imprint)
                || (IsSpecified(LegacyDocWritableFormattingProperties.Hidden) && !Hidden)
                || (IsSpecified(LegacyDocWritableFormattingProperties.NoProof) && !NoProof)
                || (IsSpecified(LegacyDocWritableFormattingProperties.Caps) && Caps == null)
                || (IsSpecified(LegacyDocWritableFormattingProperties.VerticalPosition) && VerticalPosition == null)
                || (IsSpecified(LegacyDocWritableFormattingProperties.Underline) && Underline == null)
                || (IsSpecified(LegacyDocWritableFormattingProperties.Highlight) && Highlight == null);

            internal bool IsSpecified(LegacyDocWritableFormattingProperties property) {
                return (Specified & property) != 0;
            }

            public bool Equals(LegacyDocWritableFormatting other) {
                return Bold == other.Bold
                    && Italic == other.Italic
                    && Strike == other.Strike
                    && DoubleStrike == other.DoubleStrike
                    && Outline == other.Outline
                    && Shadow == other.Shadow
                    && Emboss == other.Emboss
                    && Imprint == other.Imprint
                    && Hidden == other.Hidden
                    && NoProof == other.NoProof
                    && Special == other.Special
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
                hash = (hash * 31) + DoubleStrike.GetHashCode();
                hash = (hash * 31) + Outline.GetHashCode();
                hash = (hash * 31) + Shadow.GetHashCode();
                hash = (hash * 31) + Emboss.GetHashCode();
                hash = (hash * 31) + Imprint.GetHashCode();
                hash = (hash * 31) + Hidden.GetHashCode();
                hash = (hash * 31) + NoProof.GetHashCode();
                hash = (hash * 31) + Special.GetHashCode();
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
