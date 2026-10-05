using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static void AppendSupportedRunText(StringBuilder text, List<LegacyDocWritableRun> runs, Run run, LegacyDocWritableFootnotes footnotes, LegacyDocWritableEndnotes endnotes, LegacyDocWritablePictures pictures, OpenXmlPart ownerPart) {
            AppendSupportedRunText(text, runs, run, footnotes, endnotes, LegacyDocWritableFormatting.Plain, allowHyperlinkRunStyle: false, pictures, ownerPart);
        }

        private static void AppendSupportedRunText(StringBuilder text, List<LegacyDocWritableRun> runs, Run run, LegacyDocWritableFootnotes footnotes, LegacyDocWritableEndnotes endnotes) {
            AppendSupportedRunText(text, runs, run, footnotes, endnotes, LegacyDocWritableFormatting.Plain);
        }

        private static void AppendSupportedRunText(StringBuilder text, List<LegacyDocWritableRun> runs, Run run, LegacyDocWritableFootnotes footnotes, LegacyDocWritableEndnotes endnotes, LegacyDocWritableFormatting inheritedFormatting) {
            AppendSupportedRunText(text, runs, run, footnotes, endnotes, inheritedFormatting, allowHyperlinkRunStyle: false);
        }

        private static void AppendSupportedRunText(StringBuilder text, List<LegacyDocWritableRun> runs, Run run, LegacyDocWritableFootnotes footnotes, LegacyDocWritableEndnotes endnotes, LegacyDocWritableFormatting inheritedFormatting, bool allowHyperlinkRunStyle, LegacyDocWritablePictures? pictures = null, OpenXmlPart? ownerPart = null) {
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
                    case DeletedText deletedTextNode:
                        AppendFormattedText(text, runs, deletedTextNode.Text, formatting);
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
                    case CommentReference:
                        AppendFormattedText(
                            text,
                            runs,
                            LegacyDocCommentReader.CommentReferenceCharacter.ToString(),
                            LegacyDocWritableFormatting.SpecialCharacter);
                        break;
                    case DocumentFormat.OpenXml.Wordprocessing.Drawing drawing:
                        if (pictures == null || ownerPart == null) {
                            throw new NotSupportedException(
                                "Native DOC saving does not support inline pictures inside this run context.");
                        }

                        int picturePosition = text.Length;
                        int pictureDataOffset = pictures.AddInlinePicture(drawing, ownerPart);
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

        private static LegacyDocWritableFormatting ReadSupportedParagraphMarkRunFormatting(ParagraphProperties? paragraphProperties) {
            return ReadSupportedRunFormatting(paragraphProperties?.GetFirstChild<ParagraphMarkRunProperties>());
        }

        private static void AddParagraphMarkRunFormatting(List<LegacyDocWritableRun> runs, int paragraphMarkPosition, LegacyDocWritableFormatting formatting) {
            if (formatting.HasFormatting) {
                runs.Add(new LegacyDocWritableRun(paragraphMarkPosition, 1, formatting));
            }
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
            int? characterSpacingTwips = null;
            int? kerningMinimumFontSizeHalfPoints = null;
            ushort? languageId = null;
            ushort? eastAsiaLanguageId = null;
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
                    case Kern kern:
                        specified |= LegacyDocWritableFormattingProperties.Kerning;
                        kerningMinimumFontSizeHalfPoints = ReadSupportedKerning(kern);
                        break;
                    case Spacing spacing:
                        specified |= LegacyDocWritableFormattingProperties.CharacterSpacing;
                        characterSpacingTwips = ReadSupportedCharacterSpacing(spacing);
                        break;
                    case Languages languages:
                        LegacyDocWritableLanguageIds languageIds = ReadSupportedRunLanguages(languages);
                        if (languageIds.HasAny) {
                            specified |= LegacyDocWritableFormattingProperties.Language;
                            languageId = languageIds.LanguageId;
                            eastAsiaLanguageId = languageIds.EastAsiaLanguageId;
                        }

                        break;
                    case RunStyle runStyle when allowHyperlinkRunStyle && string.Equals(runStyle.Val?.Value, "Hyperlink", StringComparison.OrdinalIgnoreCase):
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving currently supports only bold, italic, strikethrough, double-strikethrough, outline, shadow, emboss, imprint, hidden text, proofing exclusion, caps/small-caps, superscript/subscript, underline, highlight, font size, color, font family, character spacing, kerning, and language run formatting. Unsupported run property: {property.LocalName}.");
                }
            }

            return new LegacyDocWritableFormatting(bold == true, italic == true, strike, doubleStrike, outline, shadow, emboss, imprint, hidden, noProof, false, caps, verticalPosition, underline, highlight, fontSizeHalfPoints, colorHex, fontFamily, specified, characterSpacingTwips: characterSpacingTwips, kerningMinimumFontSizeHalfPoints: kerningMinimumFontSizeHalfPoints, languageId: languageId, eastAsiaLanguageId: eastAsiaLanguageId);
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
            string hex = colorValue.Trim().TrimStart('#').ToUpperInvariant();
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

        private static int ReadSupportedCharacterSpacing(Spacing spacing) {
            int value = spacing.Val?.Value ?? 0;
            if (value < short.MinValue || value > short.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports character spacing only within the Word 97-2003 signed twip range.");
            }

            return value;
        }

        private static LegacyDocWritableLanguageIds ReadSupportedRunLanguages(Languages languages) {
            ushort? languageId = LegacyDocLanguageMapper.TryReadLanguageId(languages.Val?.Value, "run language");
            ushort? eastAsiaLanguageId = LegacyDocLanguageMapper.TryReadLanguageId(languages.EastAsia?.Value, "run East Asian language");
            ushort? bidiLanguageId = LegacyDocLanguageMapper.TryReadLanguageId(languages.Bidi?.Value, "run bidirectional language");

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
                if (previous.EndCharacter == start && previous.Formatting.Equals(formatting)) {
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
            grpprl.Add(Convert.ToByte(colorHex.Substring(0, 2), 16));
            grpprl.Add(Convert.ToByte(colorHex.Substring(2, 2), 16));
            grpprl.Add(Convert.ToByte(colorHex.Substring(4, 2), 16));
            grpprl.Add(0);
        }

    }
}
