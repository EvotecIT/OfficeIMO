using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static LegacyDocWritableParagraphFormatting ReadSupportedParagraphFormatting(ParagraphProperties? paragraphProperties, IReadOnlyDictionary<string, ushort> styleIndexes) =>
            ReadSupportedParagraphFormattingCore(paragraphProperties, styleIndexes, allowParagraphStyleId: true);

        private static LegacyDocWritableParagraphFormatting ReadSupportedStyleParagraphFormatting(StyleParagraphProperties? paragraphProperties) =>
            ReadSupportedParagraphFormattingCore(paragraphProperties, EmptyStyleIndexes, allowParagraphStyleId: false);

        private static LegacyDocWritableParagraphFormatting ReadSupportedBuiltInStyleParagraphFormatting(ushort styleIndex, StyleParagraphProperties? paragraphProperties) =>
            ReadSupportedParagraphFormattingCore(paragraphProperties, EmptyStyleIndexes, allowParagraphStyleId: false, builtInStyleIndex: styleIndex);

        private static LegacyDocWritableParagraphFormatting ReadSupportedParagraphFormattingCore(OpenXmlCompositeElement? paragraphProperties, IReadOnlyDictionary<string, ushort> styleIndexes, bool allowParagraphStyleId, ushort? builtInStyleIndex = null) {
            if (paragraphProperties == null || !paragraphProperties.HasChildren) {
                return LegacyDocWritableParagraphFormatting.Plain;
            }

            byte? alignment = null;
            int? spacingBeforeTwips = null;
            int? spacingAfterTwips = null;
            int? lineSpacingTwips = null;
            int? leftIndentTwips = null;
            int? rightIndentTwips = null;
            int? firstLineIndentTwips = null;
            bool? keepLinesTogether = null;
            bool? keepWithNext = null;
            bool? pageBreakBefore = null;
            bool? avoidWidowAndOrphan = null;
            bool? suppressLineNumbers = null;
            bool? bidirectional = null;
            ushort? numberingListIndex = null;
            byte? numberingLevel = null;
            byte? verticalCharacterAlignment = null;
            LegacyDocParagraphShading? paragraphShading = null;
            LegacyDocParagraphBorders? paragraphBorders = null;
            IReadOnlyList<LegacyDocTabStop>? tabStops = null;
            ushort? styleIndex = null;
            foreach (OpenXmlElement property in paragraphProperties.ChildElements) {
                switch (property) {
                    case ParagraphStyleId paragraphStyleId:
                        if (!allowParagraphStyleId) {
                            throw new NotSupportedException("Native DOC saving does not support nested paragraph style identifiers inside custom paragraph style definitions.");
                        }

                        styleIndex = ReadSupportedParagraphStyleIndex(paragraphStyleId, styleIndexes);
                        break;
                    case Justification justification:
                        alignment = ReadSupportedParagraphAlignment(justification);
                        break;
                    case SpacingBetweenLines spacing:
                        ReadSupportedParagraphSpacing(spacing, out spacingBeforeTwips, out spacingAfterTwips, out lineSpacingTwips);
                        break;
                    case Indentation indentation:
                        ReadSupportedParagraphIndentation(indentation, out leftIndentTwips, out rightIndentTwips, out firstLineIndentTwips);
                        break;
                    case KeepLines keepLines:
                        keepLinesTogether = ReadOnOffValue(keepLines);
                        break;
                    case KeepNext keepNext:
                        keepWithNext = ReadOnOffValue(keepNext);
                        break;
                    case PageBreakBefore pageBreakBeforeProperty:
                        pageBreakBefore = ReadOnOffValue(pageBreakBeforeProperty);
                        break;
                    case WidowControl widowControl:
                        avoidWidowAndOrphan = ReadOnOffValue(widowControl);
                        break;
                    case SuppressLineNumbers lineNumbers:
                        suppressLineNumbers = ReadOnOffValue(lineNumbers);
                        break;
                    case BiDi bidi:
                        bidirectional = ReadOnOffValue(bidi);
                        break;
                    case NumberingProperties numberingProperties:
                        ReadSupportedNumberingProperties(numberingProperties, out numberingListIndex, out numberingLevel);
                        break;
                    case TextAlignment textAlignment:
                        verticalCharacterAlignment = ReadSupportedVerticalCharacterAlignment(textAlignment);
                        break;
                    case Shading shading:
                        paragraphShading = ReadSupportedParagraphShading(shading);
                        break;
                    case ParagraphBorders borders:
                        paragraphBorders = ReadSupportedParagraphBorders(borders);
                        break;
                    case Tabs tabs:
                        tabStops = ReadSupportedTabStops(tabs);
                        break;
                    case OutlineLevel outlineLevel:
                        ReadSupportedBuiltInOutlineLevel(outlineLevel, builtInStyleIndex);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving currently supports only built-in paragraph styles, simple numbering, alignment, vertical character alignment, bidirectional layout, spacing, indentation, pagination flags, tab stops, palette-backed paragraph shading, and palette-backed paragraph borders. Unsupported paragraph property: {property.LocalName}.");
                }
            }

            return new LegacyDocWritableParagraphFormatting(
                alignment,
                styleIndex,
                spacingBeforeTwips,
                spacingAfterTwips,
                lineSpacingTwips,
                leftIndentTwips,
                rightIndentTwips,
                firstLineIndentTwips,
                keepLinesTogether,
                keepWithNext,
                pageBreakBefore,
                avoidWidowAndOrphan,
                suppressLineNumbers,
                bidirectional,
                numberingListIndex,
                numberingLevel,
                verticalCharacterAlignment,
                null,
                null,
                tabStops,
                null,
                null,
                null,
                false,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                paragraphShading,
                paragraphBorders);
        }

        private static byte ReadSupportedVerticalCharacterAlignment(TextAlignment textAlignment) {
            VerticalTextAlignmentValues value = textAlignment.Val?.Value ?? VerticalTextAlignmentValues.Auto;
            if (value == VerticalTextAlignmentValues.Auto) {
                return 0;
            } else if (value == VerticalTextAlignmentValues.Baseline) {
                return 1;
            } else if (value == VerticalTextAlignmentValues.Top) {
                return 2;
            } else if (value == VerticalTextAlignmentValues.Center) {
                return 3;
            } else if (value == VerticalTextAlignmentValues.Bottom) {
                return 4;
            }

            throw new NotSupportedException($"Native DOC saving does not support paragraph vertical character alignment '{value}'.");
        }

        private static void ReadSupportedNumberingProperties(NumberingProperties numberingProperties, out ushort? numberingListIndex, out byte? numberingLevel) {
            numberingListIndex = null;
            numberingLevel = null;

            foreach (OpenXmlElement child in numberingProperties.ChildElements) {
                switch (child) {
                    case NumberingId:
                    case NumberingLevelReference:
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple paragraph numbering only. Unsupported numbering property: {child.LocalName}.");
                }
            }

            int? numberId = numberingProperties.NumberingId?.Val?.Value;
            if (numberId == null || numberId.Value == 0) {
                return;
            }

            if (numberId.Value < 0 || numberId.Value > ushort.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports paragraph numbering ids only within the Word 97-2003 list index range.");
            }

            int level = numberingProperties.NumberingLevelReference?.Val?.Value ?? 0;
            if (level < 0 || level > 8) {
                throw new NotSupportedException("Native DOC saving supports paragraph numbering levels 0 through 8.");
            }

            numberingListIndex = checked((ushort)numberId.Value);
            numberingLevel = checked((byte)level);
        }

        private static LegacyDocParagraphBorders ReadSupportedParagraphBorders(ParagraphBorders borders) {
            foreach (OpenXmlElement child in borders.ChildElements) {
                switch (child) {
                    case TopBorder:
                    case LeftBorder:
                    case BottomBorder:
                    case RightBorder:
                    case BetweenBorder:
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple paragraph borders only. Unsupported paragraph border: {child.LocalName}.");
                }
            }

            return new LegacyDocParagraphBorders(
                ReadSupportedParagraphBorder(borders.TopBorder),
                ReadSupportedParagraphBorder(borders.LeftBorder),
                ReadSupportedParagraphBorder(borders.BottomBorder),
                ReadSupportedParagraphBorder(borders.RightBorder),
                ReadSupportedParagraphBorder(borders.BetweenBorder));
        }

        private static LegacyDocParagraphBorder ReadSupportedParagraphBorder(BorderType? border) {
            if (border == null) {
                return default;
            }

            BorderValues? value = border.Val?.Value;
            if (value == null || value == BorderValues.None || value == BorderValues.Nil) {
                return default;
            }

            LegacyDocParagraphBorderStyle style = MapSupportedParagraphBorderStyle(value.Value);
            string? colorHex = border.Color?.Value;
            if (string.Equals(colorHex, "auto", StringComparison.OrdinalIgnoreCase)) {
                colorHex = null;
            }

            if (!LegacyDocColorPalette.TryGetIcoForHex(colorHex, out _)) {
                throw new NotSupportedException("Native DOC saving supports paragraph borders only with Word 97-2003 palette colors.");
            }

            int size = border.Size?.Value == null ? 4 : checked((int)border.Size.Value);
            int space = border.Space?.Value == null ? 0 : checked((int)border.Space.Value);
            if (size <= 0 || size > byte.MaxValue || space < 0 || space > byte.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports paragraph border size and spacing only within Word 97-2003 BRC80 byte ranges.");
            }

            return new LegacyDocParagraphBorder(style, colorHex, size, space);
        }

        private static LegacyDocParagraphBorderStyle MapSupportedParagraphBorderStyle(BorderValues value) {
            if (value == BorderValues.Single) {
                return LegacyDocParagraphBorderStyle.Single;
            }

            if (value == BorderValues.Double) {
                return LegacyDocParagraphBorderStyle.Double;
            }

            if (value == BorderValues.Dotted) {
                return LegacyDocParagraphBorderStyle.Dotted;
            }

            if (value == BorderValues.Dashed || value == BorderValues.DashSmallGap) {
                return LegacyDocParagraphBorderStyle.Dashed;
            }

            throw new NotSupportedException($"Native DOC saving does not support paragraph border style '{value}'.");
        }

        private static LegacyDocParagraphShading ReadSupportedParagraphShading(Shading shading) {
            ShadingPatternValues? pattern = shading.Val?.Value;
            if (pattern != null && pattern != ShadingPatternValues.Clear) {
                throw new NotSupportedException("Native DOC saving supports paragraph shading only for clear fill patterns.");
            }

            string? fillColorHex = shading.Fill?.Value;
            if (string.IsNullOrWhiteSpace(fillColorHex)
                || string.Equals(fillColorHex, "auto", StringComparison.OrdinalIgnoreCase)) {
                return default;
            }

            if (!LegacyDocColorPalette.TryGetIcoForHex(fillColorHex, out _)) {
                throw new NotSupportedException("Native DOC saving supports paragraph shading only for Word 97-2003 palette fill colors.");
            }

            return new LegacyDocParagraphShading(fillColorHex);
        }

        private static void ReadSupportedBuiltInOutlineLevel(OutlineLevel outlineLevel, ushort? builtInStyleIndex) {
            if (builtInStyleIndex == null || builtInStyleIndex.Value == 0 || builtInStyleIndex.Value > 9) {
                throw new NotSupportedException("Native DOC saving currently supports outline-level style formatting only for built-in Heading styles.");
            }

            int expectedLevel = builtInStyleIndex.Value - 1;
            int level = outlineLevel.Val?.Value ?? expectedLevel;
            if (level != expectedLevel) {
                throw new NotSupportedException($"Native DOC saving supports built-in Heading style outline levels only when they match the heading level. Expected outline level {expectedLevel}, got {level}.");
            }
        }

        private static IReadOnlyList<LegacyDocTabStop> ReadSupportedTabStops(Tabs tabs) {
            var tabStops = new List<LegacyDocTabStop>();
            foreach (TabStop tabStop in tabs.Elements<TabStop>()) {
                if (tabStop.Position == null) {
                    throw new NotSupportedException("Native DOC saving supports tab stops only when each tab stop has an explicit twip position.");
                }

                if (!TryReadSupportedTabAlignment(tabStop.Val?.Value ?? TabStopValues.Left, out LegacyDocTabStopAlignment alignment)) {
                    throw new NotSupportedException($"Native DOC saving does not support tab stop alignment '{tabStop.Val?.Value}'.");
                }

                if (!TryReadSupportedTabLeader(tabStop.Leader?.Value ?? TabStopLeaderCharValues.None, out LegacyDocTabStopLeader leader)) {
                    throw new NotSupportedException($"Native DOC saving does not support tab stop leader '{tabStop.Leader?.Value}'.");
                }

                tabStops.Add(new LegacyDocTabStop(tabStop.Position.Value, alignment, leader));
            }

            return tabStops;
        }

        private static bool TryReadSupportedTabAlignment(TabStopValues value, out LegacyDocTabStopAlignment alignment) {
            if (value == TabStopValues.Left) {
                alignment = LegacyDocTabStopAlignment.Left;
                return true;
            } else if (value == TabStopValues.Center) {
                alignment = LegacyDocTabStopAlignment.Center;
                return true;
            } else if (value == TabStopValues.Right) {
                alignment = LegacyDocTabStopAlignment.Right;
                return true;
            } else if (value == TabStopValues.Decimal) {
                alignment = LegacyDocTabStopAlignment.Decimal;
                return true;
            } else if (value == TabStopValues.Bar) {
                alignment = LegacyDocTabStopAlignment.Bar;
                return true;
            } else if (value == TabStopValues.Clear) {
                alignment = LegacyDocTabStopAlignment.Clear;
                return true;
            }

            alignment = LegacyDocTabStopAlignment.Left;
            return false;
        }

        private static bool TryReadSupportedTabLeader(TabStopLeaderCharValues value, out LegacyDocTabStopLeader leader) {
            if (value == TabStopLeaderCharValues.None) {
                leader = LegacyDocTabStopLeader.None;
                return true;
            } else if (value == TabStopLeaderCharValues.Dot) {
                leader = LegacyDocTabStopLeader.Dot;
                return true;
            } else if (value == TabStopLeaderCharValues.Hyphen) {
                leader = LegacyDocTabStopLeader.Hyphen;
                return true;
            } else if (value == TabStopLeaderCharValues.Underscore) {
                leader = LegacyDocTabStopLeader.Underscore;
                return true;
            } else if (value == TabStopLeaderCharValues.Heavy) {
                leader = LegacyDocTabStopLeader.Heavy;
                return true;
            } else if (value == TabStopLeaderCharValues.MiddleDot) {
                leader = LegacyDocTabStopLeader.MiddleDot;
                return true;
            }

            leader = LegacyDocTabStopLeader.None;
            return false;
        }

        private static ushort? ReadSupportedParagraphStyleIndex(ParagraphStyleId paragraphStyleId, IReadOnlyDictionary<string, ushort> styleIndexes) {
            string? styleId = paragraphStyleId.Val?.Value;
            if (string.IsNullOrWhiteSpace(styleId)
                || string.Equals(styleId, "Normal", StringComparison.OrdinalIgnoreCase)
                || string.Equals(styleId, "ListParagraph", StringComparison.OrdinalIgnoreCase)) {
                return null;
            }

            if (TryMapBuiltInParagraphStyleIndex(styleId!, out ushort styleIndex)) {
                return styleIndex;
            }

            if (styleIndexes.TryGetValue(styleId!, out ushort customStyleIndex)) {
                return customStyleIndex;
            }

            throw new NotSupportedException($"Native DOC saving currently supports only built-in Normal and Heading1 through Heading9 paragraph styles. Unsupported paragraph style: {styleId}.");
        }

        private static bool TryMapBuiltInParagraphStyleIndex(string styleId, out ushort styleIndex) {
            switch (styleId.Trim().ToUpperInvariant()) {
                case "HEADING1":
                    styleIndex = 1;
                    return true;
                case "HEADING2":
                    styleIndex = 2;
                    return true;
                case "HEADING3":
                    styleIndex = 3;
                    return true;
                case "HEADING4":
                    styleIndex = 4;
                    return true;
                case "HEADING5":
                    styleIndex = 5;
                    return true;
                case "HEADING6":
                    styleIndex = 6;
                    return true;
                case "HEADING7":
                    styleIndex = 7;
                    return true;
                case "HEADING8":
                    styleIndex = 8;
                    return true;
                case "HEADING9":
                    styleIndex = 9;
                    return true;
                default:
                    styleIndex = 0;
                    return false;
            }
        }

        private static byte? ReadSupportedParagraphAlignment(Justification justification) {
            JustificationValues value = justification.Val?.Value ?? JustificationValues.Left;
            if (value == JustificationValues.Left) {
                return 0;
            } else if (value == JustificationValues.Center) {
                return 1;
            } else if (value == JustificationValues.Right) {
                return 2;
            } else if (value == JustificationValues.Both) {
                return 3;
            }

            throw new NotSupportedException($"Native DOC saving does not support paragraph alignment '{value}'.");
        }

        private static void ReadSupportedParagraphSpacing(
            SpacingBetweenLines spacing,
            out int? spacingBeforeTwips,
            out int? spacingAfterTwips,
            out int? lineSpacingTwips) {
            if ((spacing.BeforeAutoSpacing?.Value ?? false) || (spacing.AfterAutoSpacing?.Value ?? false)) {
                throw new NotSupportedException("Native DOC saving currently supports paragraph spacing only as explicit twip values, not automatic spacing.");
            }

            if (spacing.BeforeLines != null || spacing.AfterLines != null) {
                throw new NotSupportedException("Native DOC saving currently supports paragraph before/after spacing only as twip values, not line-count spacing.");
            }

            LineSpacingRuleValues? lineRule = spacing.LineRule?.Value;
            if (lineRule == LineSpacingRuleValues.Auto) {
                throw new NotSupportedException("Native DOC saving currently supports exact or at-least paragraph line spacing, not automatic multiplier spacing.");
            }

            spacingBeforeTwips = ReadOptionalInt32Twips(spacing.Before?.Value, "paragraph spacing before");
            spacingAfterTwips = ReadOptionalInt32Twips(spacing.After?.Value, "paragraph spacing after");
            lineSpacingTwips = ReadOptionalInt32Twips(spacing.Line?.Value, "paragraph line spacing");
        }

        private static void ReadSupportedParagraphIndentation(
            Indentation indentation,
            out int? leftIndentTwips,
            out int? rightIndentTwips,
            out int? firstLineIndentTwips) {
            if (indentation.LeftChars != null || indentation.RightChars != null || indentation.FirstLineChars != null || indentation.HangingChars != null) {
                throw new NotSupportedException("Native DOC saving currently supports paragraph indentation only as twip values, not character-based indentation.");
            }

            if (indentation.FirstLine != null && indentation.Hanging != null) {
                throw new NotSupportedException("Native DOC saving cannot write a paragraph with both first-line and hanging indentation.");
            }

            leftIndentTwips = ReadOptionalInt32Twips(indentation.Left?.Value, "paragraph left indentation");
            rightIndentTwips = ReadOptionalInt32Twips(indentation.Right?.Value, "paragraph right indentation");
            int? firstLine = ReadOptionalInt32Twips(indentation.FirstLine?.Value, "paragraph first-line indentation");
            int? hanging = ReadOptionalInt32Twips(indentation.Hanging?.Value, "paragraph hanging indentation");
            firstLineIndentTwips = hanging != null ? -hanging.Value : firstLine;
        }

        private static int? ReadOptionalInt32Twips(string? value, string propertyName) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            if (!int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int result)) {
                throw new NotSupportedException($"Native DOC saving supports {propertyName} only when it is stored as a numeric twip value.");
            }

            if (result < short.MinValue || result > short.MaxValue) {
                throw new NotSupportedException($"Native DOC saving supports {propertyName} only within the Word 97-2003 signed twip range.");
            }

            return result;
        }

        private static bool? ReadOnOffValue(OnOffType property) {
            if (property.Val == null || property.Val.Value) {
                return true;
            }

            return null;
        }
    }
}
