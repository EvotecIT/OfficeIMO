using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static void AppendSupportedRunText(StringBuilder text, List<LegacyDocWritableRun> runs, LegacyDocElement run, LegacyDocWritableFootnotes footnotes, LegacyDocWritableEndnotes endnotes, LegacyDocWritablePictures pictures, LegacyDocPart ownerPart) {
            AppendSupportedRunText(text, runs, run, footnotes, endnotes, LegacyDocWritableFormatting.Plain, allowHyperlinkRunStyle: false, pictures, ownerPart);
        }

        private static void AppendSupportedRunText(StringBuilder text, List<LegacyDocWritableRun> runs, LegacyDocElement run, LegacyDocWritableFootnotes footnotes, LegacyDocWritableEndnotes endnotes) {
            AppendSupportedRunText(text, runs, run, footnotes, endnotes, LegacyDocWritableFormatting.Plain);
        }

        private static void AppendSupportedRunText(StringBuilder text, List<LegacyDocWritableRun> runs, LegacyDocElement run, LegacyDocWritableFootnotes footnotes, LegacyDocWritableEndnotes endnotes, LegacyDocWritableFormatting inheritedFormatting) {
            AppendSupportedRunText(text, runs, run, footnotes, endnotes, inheritedFormatting, allowHyperlinkRunStyle: false);
        }

        private static void AppendSupportedRunText(StringBuilder text, List<LegacyDocWritableRun> runs, LegacyDocElement run, LegacyDocWritableFootnotes footnotes, LegacyDocWritableEndnotes endnotes, LegacyDocWritableFormatting inheritedFormatting, bool allowHyperlinkRunStyle, LegacyDocWritablePictures? pictures = null, LegacyDocPart? ownerPart = null) {
            if (run.HasChild("footnoteReference")) {
                AppendFootnoteReferenceRun(text, runs, footnotes, run, inheritedFormatting);
                return;
            }

            if (run.HasChild("endnoteReference")) {
                AppendEndnoteReferenceRun(text, runs, endnotes, run, inheritedFormatting);
                return;
            }

            LegacyDocWritableFormatting formatting = ReadSupportedRunFormatting(run.Child("rPr"), allowHyperlinkRunStyle)
                .WithInheritedFormatting(inheritedFormatting);

            foreach (LegacyDocElement child in run.Elements()) {
                switch (child.WName) {
                    case "rPr":
                        break;
                    case "lastRenderedPageBreak":
                        break;
                    case "pgNum":
                        AppendSupportedPageNumberField(text, runs, formatting);
                        break;
                    case "t":
                        AppendFormattedText(text, runs, child.Text, formatting);
                        break;
                    case "delText":
                        AppendFormattedText(text, runs, child.Text, formatting);
                        break;
                    case "tab":
                        AppendFormattedText(text, runs, "\t", formatting);
                        break;
                    case "cr":
                        AppendFormattedText(text, runs, LegacyDocSpecialCharacters.TextWrappingBreak.ToString(), formatting);
                        break;
                    case "noBreakHyphen":
                        AppendFormattedText(text, runs, LegacyDocSpecialCharacters.NoBreakHyphen.ToString(), formatting);
                        break;
                    case "softHyphen":
                        AppendFormattedText(text, runs, LegacyDocSpecialCharacters.SoftHyphen.ToString(), formatting);
                        break;
                    case "br":
                        AppendSupportedBreak(text, runs, child, formatting);
                        break;
                    case "footnoteReference":
                        AppendFootnoteReference(text, runs, footnotes, child, formatting);
                        break;
                    case "endnoteReference":
                        AppendEndnoteReference(text, runs, endnotes, child, formatting);
                        break;
                    case "commentReference":
                        AppendFormattedText(
                            text,
                            runs,
                            LegacyDocCommentReader.CommentReferenceCharacter.ToString(),
                            formatting.WithSpecialCharacter());
                        break;
                    case "drawing":
                        if (pictures == null || ownerPart == null) {
                            throw new NotSupportedException(
                                "Native DOC saving does not support inline pictures inside this run context.");
                        }

                        int picturePosition = text.Length;
                        int pictureDataOffset = pictures.AddInlinePicture(child, ownerPart.Value);
                        text.Append('\u0001');
                        runs.Add(new LegacyDocWritableRun(
                            picturePosition,
                            1,
                            LegacyDocWritableFormatting.SpecialCharacter.WithRevision(formatting.Revision),
                            pictureDataOffset));
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving currently supports text, embedded inline pictures, tabs, page-number fields, carriage returns, soft/no-break hyphens, text-wrapping/page/column breaks, and simple footnote/endnote/comment references only. Unsupported run element: {child.LocalName}.");
                }
            }
        }

        private static void AppendSupportedBreak(StringBuilder text, List<LegacyDocWritableRun> runs, LegacyDocElement breakNode, LegacyDocWritableFormatting formatting) {
            string? breakType = breakNode.Attr("type");
            if (breakType == null || breakType == "textWrapping") {
                AppendFormattedText(text, runs, LegacyDocSpecialCharacters.TextWrappingBreak.ToString(), formatting);
                return;
            }

            if (breakType == "page") {
                AppendFormattedText(text, runs, LegacyDocSpecialCharacters.PageBreak.ToString(), formatting);
                return;
            }

            if (breakType == "column") {
                AppendFormattedText(text, runs, LegacyDocSpecialCharacters.ColumnBreak.ToString(), formatting);
                return;
            }

            throw new NotSupportedException($"Native DOC saving currently supports text-wrapping, page, and column breaks only. Unsupported break type: {breakType}.");
        }

        private static LegacyDocWritableFormatting ReadSupportedRunFormatting(LegacyDocElement runProperties) {
            return ReadSupportedRunFormatting(runProperties, allowHyperlinkRunStyle: false);
        }

        private static LegacyDocWritableFormatting ReadSupportedParagraphMarkRunFormatting(LegacyDocElement paragraphProperties) {
            return ReadSupportedRunFormatting(paragraphProperties.Child("rPr"));
        }

        private static void AddParagraphMarkRunFormatting(List<LegacyDocWritableRun> runs, int paragraphMarkPosition, LegacyDocWritableFormatting formatting) {
            if (formatting.HasFormatting) {
                runs.Add(new LegacyDocWritableRun(paragraphMarkPosition, 1, formatting));
            }
        }

        private static LegacyDocWritableFormatting ReadSupportedRunFormatting(LegacyDocElement runProperties, bool allowHyperlinkRunStyle, bool allowNoteReferenceRunStyle = false) {
            if (runProperties.IsNone || !runProperties.HasElements) {
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
            int? characterSpacingTwips = null;
            int? characterScalePercentage = null;
            int? kerningMinimumFontSizeHalfPoints = null;
            ushort? languageId = null;
            ushort? eastAsiaLanguageId = null;
            LegacyDocWritableFormattingProperties specified = LegacyDocWritableFormattingProperties.None;
            foreach (LegacyDocElement property in runProperties.Elements()) {
                switch (property.WName) {
                    case "b":
                        specified |= LegacyDocWritableFormattingProperties.Bold;
                        bold = MergeSingleRunToggle(bold, IsEnabled(property), "bold", "Bold", "BoldComplexScript");
                        break;
                    case "bCs":
                        specified |= LegacyDocWritableFormattingProperties.Bold;
                        bold = MergeSingleRunToggle(bold, IsEnabled(property), "bold", "Bold", "BoldComplexScript");
                        break;
                    case "i":
                        specified |= LegacyDocWritableFormattingProperties.Italic;
                        italic = MergeSingleRunToggle(italic, IsEnabled(property), "italic", "Italic", "ItalicComplexScript");
                        break;
                    case "iCs":
                        specified |= LegacyDocWritableFormattingProperties.Italic;
                        italic = MergeSingleRunToggle(italic, IsEnabled(property), "italic", "Italic", "ItalicComplexScript");
                        break;
                    case "strike":
                        specified |= LegacyDocWritableFormattingProperties.Strike;
                        strike = IsEnabled(property);
                        break;
                    case "dstrike":
                        specified |= LegacyDocWritableFormattingProperties.DoubleStrike;
                        doubleStrike = IsEnabled(property);
                        break;
                    case "outline":
                        specified |= LegacyDocWritableFormattingProperties.Outline;
                        outline = IsEnabled(property);
                        break;
                    case "shadow":
                        specified |= LegacyDocWritableFormattingProperties.Shadow;
                        shadow = IsEnabled(property);
                        break;
                    case "emboss":
                        specified |= LegacyDocWritableFormattingProperties.Emboss;
                        emboss = IsEnabled(property);
                        break;
                    case "imprint":
                        specified |= LegacyDocWritableFormattingProperties.Imprint;
                        imprint = IsEnabled(property);
                        break;
                    case "vanish":
                        specified |= LegacyDocWritableFormattingProperties.Hidden;
                        hidden = IsEnabled(property);
                        break;
                    case "noProof":
                        specified |= LegacyDocWritableFormattingProperties.NoProof;
                        noProof = IsEnabled(property);
                        break;
                    case "caps":
                        specified |= LegacyDocWritableFormattingProperties.Caps;
                        caps = MergeCapsKind(caps, IsEnabled(property), 1);
                        break;
                    case "smallCaps":
                        specified |= LegacyDocWritableFormattingProperties.Caps;
                        caps = MergeCapsKind(caps, IsEnabled(property), 2);
                        break;
                    case "vertAlign":
                        specified |= LegacyDocWritableFormattingProperties.VerticalPosition;
                        verticalPosition = ReadSupportedVerticalPosition(property);
                        break;
                    case "u":
                        specified |= LegacyDocWritableFormattingProperties.Underline;
                        underline = ReadSupportedUnderline(property);
                        break;
                    case "highlight":
                        specified |= LegacyDocWritableFormattingProperties.Highlight;
                        highlight = ReadSupportedHighlight(property);
                        break;
                    case "sz":
                        specified |= LegacyDocWritableFormattingProperties.FontSize;
                        fontSizeHalfPoints = MergeFontSizeHalfPoints(fontSizeHalfPoints, ReadFontSizeHalfPoints(property.Val));
                        break;
                    case "szCs":
                        specified |= LegacyDocWritableFormattingProperties.FontSize;
                        fontSizeHalfPoints = MergeFontSizeHalfPoints(fontSizeHalfPoints, ReadFontSizeHalfPoints(property.Val));
                        break;
                    case "color":
                        specified |= LegacyDocWritableFormattingProperties.Color;
                        colorHex = ReadSupportedColorHex(property);
                        break;
                    case "rFonts":
                        specified |= LegacyDocWritableFormattingProperties.FontFamily;
                        fontFamily = ReadSupportedRunFontFamily(property);
                        break;
                    case "kern":
                        specified |= LegacyDocWritableFormattingProperties.Kerning;
                        kerningMinimumFontSizeHalfPoints = ReadSupportedKerning(property);
                        break;
                    case "w":
                        specified |= LegacyDocWritableFormattingProperties.CharacterScale;
                        characterScalePercentage = ReadSupportedCharacterScale(property);
                        break;
                    case "spacing":
                        specified |= LegacyDocWritableFormattingProperties.CharacterSpacing;
                        characterSpacingTwips = ReadSupportedCharacterSpacing(property);
                        break;
                    case "lang":
                        LegacyDocWritableLanguageIds languageIds = ReadSupportedRunLanguages(property);
                        if (languageIds.HasAny) {
                            specified |= LegacyDocWritableFormattingProperties.Language;
                            languageId = languageIds.LanguageId;
                            eastAsiaLanguageId = languageIds.EastAsiaLanguageId;
                        }

                        break;
                    case "rStyle" when allowNoteReferenceRunStyle &&
                        (string.Equals(property.Val, "FootnoteReference", StringComparison.Ordinal) ||
                         string.Equals(property.Val, "EndnoteReference", StringComparison.Ordinal)):
                        break;
                    case "rStyle" when allowHyperlinkRunStyle && string.Equals(property.Val, "Hyperlink", StringComparison.Ordinal):
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving currently supports only bold, italic, strikethrough, double-strikethrough, outline, shadow, emboss, imprint, hidden text, proofing exclusion, caps/small-caps, superscript/subscript, underline, highlight, font size, color, font family, character spacing, character scale, kerning, and language run formatting. Unsupported run property: {property.LocalName}.");
                }
            }

            return new LegacyDocWritableFormatting(bold == true, italic == true, strike, doubleStrike, outline, shadow, emboss, imprint, hidden, noProof, false, caps, verticalPosition, underline, highlight, fontSizeHalfPoints, colorHex, fontFamily, specified, characterSpacingTwips: characterSpacingTwips, characterScalePercentage: characterScalePercentage, kerningMinimumFontSizeHalfPoints: kerningMinimumFontSizeHalfPoints, languageId: languageId, eastAsiaLanguageId: eastAsiaLanguageId);
        }

        private static bool IsEnabled(LegacyDocElement property) {
            return LegacyDocValues.IsOn(property);
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

        private static byte? ReadSupportedUnderline(LegacyDocElement underline) {
            string value = underline.Val ?? "single";
            switch (value) {
                case "none":
                    return null;
                case "single":
                    return 1;
                case "words":
                    return 2;
                case "double":
                    return 3;
                case "dotted":
                    return 4;
                case "thick":
                    return 6;
                case "dash":
                    return 7;
                case "dotDash":
                    return 8;
                case "dotDotDash":
                    return 9;
                case "wave":
                    return 10;
                case "dottedHeavy":
                    return 11;
                case "dashedHeavy":
                    return 12;
                case "dashDotHeavy":
                    return 13;
                case "dashDotDotHeavy":
                    return 14;
                case "wavyHeavy":
                    return 15;
                case "dashLong":
                    return 16;
                case "wavyDouble":
                    return 17;
                case "dashLongHeavy":
                    return 18;
            }

            throw new NotSupportedException($"Native DOC saving does not support underline style '{value}'.");
        }

        private static byte? ReadSupportedVerticalPosition(LegacyDocElement verticalTextAlignment) {
            string? value = verticalTextAlignment.Val;
            if (value == null) {
                return null;
            }

            if (value == "baseline") {
                return null;
            } else if (value == "superscript") {
                return 1;
            } else if (value == "subscript") {
                return 2;
            }

            throw new NotSupportedException($"Native DOC saving does not support vertical text alignment '{value}'.");
        }

        private static byte? ReadSupportedHighlight(LegacyDocElement highlight) {
            string? value = highlight.Val;
            if (value == null || value == "none") {
                return null;
            }

            switch (value) {
                case "black": return 1;
                case "blue": return 2;
                case "cyan": return 3;
                case "green": return 4;
                case "magenta": return 5;
                case "red": return 6;
                case "yellow": return 7;
                case "white": return 8;
                case "darkBlue": return 9;
                case "darkCyan": return 10;
                case "darkGreen": return 11;
                case "darkMagenta": return 12;
                case "darkRed": return 13;
                case "darkYellow": return 14;
                case "darkGray": return 15;
                case "lightGray": return 16;
            }

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

        private static string? ReadSupportedColorHex(LegacyDocElement color) {
            string? value = color.Val;
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            // Automatic color is retained: it overrides an inherited RGB color.
            if (string.Equals(value, "auto", StringComparison.OrdinalIgnoreCase)) {
                return "auto";
            }

            string colorValue = value!;
            string hex = colorValue.Trim().TrimStart('#').ToUpperInvariant();
            if (hex.Length != 6 || hex.Any(character => !Uri.IsHexDigit(character))) {
                throw new NotSupportedException("Native DOC saving supports text color only as a 6-digit RGB hex value or automatic color.");
            }

            return hex;
        }

        private static string? ReadSupportedRunFontFamily(LegacyDocElement runFonts) {
            return ReadSupportedRunFontFamily(ResolveNativeThemeRunFonts(runFonts));
        }

        private static string? ReadSupportedRunFontFamily(LegacyDocRunFonts runFonts) {
            string? ascii = NormalizeFontFamily(runFonts.Ascii);
            string? highAnsi = NormalizeFontFamily(runFonts.HighAnsi);
            string? eastAsia = NormalizeFontFamily(runFonts.EastAsia);
            string? complexScript = NormalizeFontFamily(runFonts.ComplexScript);

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

        private static int ReadSupportedCharacterSpacing(LegacyDocElement spacing) {
            int value = LegacyDocValues.Int32(spacing.Val) ?? 0;
            if (value < short.MinValue || value > short.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports character spacing only within the Word 97-2003 signed twip range.");
            }

            return value;
        }

        private static LegacyDocWritableLanguageIds ReadSupportedRunLanguages(LegacyDocElement languages) {
            ushort? languageId = LegacyDocLanguageMapper.TryReadLanguageId(languages.Val, "run language");
            ushort? eastAsiaLanguageId = LegacyDocLanguageMapper.TryReadLanguageId(languages.Attr("eastAsia"), "run East Asian language");
            ushort? bidiLanguageId = LegacyDocLanguageMapper.TryReadLanguageId(languages.Attr("bidi"), "run bidirectional language");

            if (bidiLanguageId != null) {
                if (languageId != null && languageId.Value != bidiLanguageId.Value) {
                    throw new NotSupportedException("Native DOC saving supports run bidirectional language only when it matches the primary run language.");
                }

                languageId ??= bidiLanguageId;
            }

            return new LegacyDocWritableLanguageIds(languageId, eastAsiaLanguageId);
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
                if (!previous.FieldLocked && previous.EndCharacter == start && previous.Formatting.HasSameEncoding(formatting)) {
                    runs[runs.Count - 1] = previous.Extend(length);
                    return;
                }
            }

            runs.Add(new LegacyDocWritableRun(start, length, formatting));
        }

        private static void WriteChpxFkp(
            byte[] stream,
            int pageOffset,
            IReadOnlyList<LegacyDocWritableSegment> segments,
            IReadOnlyDictionary<string, int> fontFamilyIndexes,
            IReadOnlyDictionary<string, int> revisionAuthorIndexes,
            int bytesPerCharacter) {
            if (segments.Count == 0 || segments.Count > byte.MaxValue) {
                throw new NotSupportedException("Native DOC saving encountered a character-format run that cannot fit in a character-format page.");
            }

            int rgbOffset = pageOffset + ((segments.Count + 1) * 4);
            int chpxOffset = AlignToEven((segments.Count + 1) * 4 + segments.Count);
            if (chpxOffset >= OleSectorSize - 1 || chpxOffset / 2 > byte.MaxValue) {
                throw new NotSupportedException("Native DOC saving encountered a character-format run that cannot fit in a character-format page.");
            }

            for (int index = 0; index < segments.Count; index++) {
                LegacyDocWritableSegment segment = segments[index];
                WriteInt32(stream, pageOffset + (index * 4), TextOffset + (segment.StartCharacter * bytesPerCharacter));
                if (segment.HasFormatting) {
                    byte[] chpx = CreateChpx(segment.Formatting, fontFamilyIndexes, revisionAuthorIndexes, segment.PictureDataOffset);
                    chpxOffset = AlignToEven(chpxOffset);
                    if (chpxOffset + chpx.Length >= OleSectorSize - 1 || chpxOffset / 2 > byte.MaxValue) {
                        throw new NotSupportedException("Native DOC saving encountered a character-format run that cannot fit in a character-format page.");
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

        private static byte[] CreateChpx(
            LegacyDocWritableFormatting formatting,
            IReadOnlyDictionary<string, int> fontFamilyIndexes,
            IReadOnlyDictionary<string, int> revisionAuthorIndexes,
            int? pictureDataOffset = null) {
            List<byte> grpprl = CreateCharacterGrpprl(formatting, fontFamilyIndexes, revisionAuthorIndexes);
            if (pictureDataOffset != null) {
                AddInt32Sprm(grpprl, SprmCPicLocation, pictureDataOffset.Value);
            }
            if (grpprl.Count > byte.MaxValue) {
                throw new NotSupportedException("Native DOC saving cannot write a character-format record larger than 255 bytes.");
            }

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

        private static List<byte> CreateCharacterGrpprl(
            LegacyDocWritableFormatting formatting,
            IReadOnlyDictionary<string, int> fontFamilyIndexes,
            IReadOnlyDictionary<string, int>? revisionAuthorIndexes = null) {
            var grpprl = new List<byte>(18);
            AddRevisionSprms(grpprl, formatting.Revision, revisionAuthorIndexes);
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

            if (formatting.LanguageId != null) {
                AddUInt16Sprm(grpprl, SprmCRgLid0, formatting.LanguageId.Value);
            }

            if (formatting.EastAsiaLanguageId != null) {
                AddUInt16Sprm(grpprl, SprmCRgLid1, formatting.EastAsiaLanguageId.Value);
            }

            if (formatting.KerningMinimumFontSizeHalfPoints.HasValue) {
                AddInt16CharacterSprm(grpprl, SprmCHpsKern, formatting.KerningMinimumFontSizeHalfPoints.Value);
            }

            if (formatting.CharacterScalePercentage.HasValue) {
                AddUInt16Sprm(grpprl, SprmCCharScale, (ushort)formatting.CharacterScalePercentage.Value);
            }

            if (formatting.CharacterSpacingTwips != null || formatting.IsSpecified(LegacyDocWritableFormattingProperties.CharacterSpacing)) {
                AddInt16CharacterSprm(grpprl, SprmCDxaSpace, formatting.CharacterSpacingTwips ?? 0);
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

        private static void AddInt32Sprm(List<byte> grpprl, ushort sprm, int operand) {
            grpprl.Add((byte)(sprm & 0xFF));
            grpprl.Add((byte)(sprm >> 8));
            grpprl.Add((byte)operand);
            grpprl.Add((byte)(operand >> 8));
            grpprl.Add((byte)(operand >> 16));
            grpprl.Add((byte)(operand >> 24));
        }

        private static void AddInt16CharacterSprm(List<byte> grpprl, ushort sprm, int operand) {
            if (operand < short.MinValue || operand > short.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports character spacing only within the Word 97-2003 signed twip range.");
            }

            short value = checked((short)operand);
            grpprl.Add((byte)(sprm & 0xFF));
            grpprl.Add((byte)(sprm >> 8));
            grpprl.Add((byte)(value & 0xFF));
            grpprl.Add((byte)(value >> 8));
        }

        private static void AddColorRefSprm(List<byte> grpprl, string colorHex) {
            grpprl.Add((byte)(SprmCCv & 0xFF));
            grpprl.Add((byte)(SprmCCv >> 8));
            if (string.Equals(colorHex, "auto", StringComparison.OrdinalIgnoreCase)) {
                // COLORREF with fAuto set (0xFF000000) is Word's automatic color.
                grpprl.AddRange(new byte[] { 0, 0, 0, 0xFF });
                return;
            }

            grpprl.Add(Convert.ToByte(colorHex.Substring(0, 2), 16));
            grpprl.Add(Convert.ToByte(colorHex.Substring(2, 2), 16));
            grpprl.Add(Convert.ToByte(colorHex.Substring(4, 2), 16));
            grpprl.Add(0);
        }

    }
}
