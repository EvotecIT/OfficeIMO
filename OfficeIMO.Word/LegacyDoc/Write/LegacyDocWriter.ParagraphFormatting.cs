using OfficeIMO.Word.LegacyDoc.Model;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        /// <summary>
        /// Which <c>w:pPr</c> content model a paragraph property list follows. A paragraph's own properties may carry a
        /// style reference, mark run properties and section properties; a style's carry none of them; a list level's
        /// carry only a style reference. Elements outside the model are unsupported properties.
        /// </summary>
        private enum LegacyDocParagraphPropertiesKind {
            Paragraph,
            Style,
            ListLevel
        }

        /// <summary>
        /// The text an Open XML SDK enumeration value gave when interpolated into a message (its type name with empty
        /// braces), or empty when the attribute is absent. Messages keep that wording.
        /// </summary>
        private static LegacyDocWritableParagraphFormatting ReadSupportedParagraphFormatting(LegacyDocElement paragraphProperties, IReadOnlyDictionary<string, ushort> styleIndexes) =>
            ReadSupportedParagraphFormattingCore(paragraphProperties, styleIndexes, allowParagraphStyleId: true, kind: LegacyDocParagraphPropertiesKind.Paragraph);

        private static LegacyDocWritableParagraphFormatting ReadSupportedBodyParagraphFormatting(LegacyDocElement paragraphProperties, IReadOnlyDictionary<string, ushort> styleIndexes) =>
            ReadSupportedParagraphFormattingCore(paragraphProperties, styleIndexes, allowParagraphStyleId: true, kind: LegacyDocParagraphPropertiesKind.Paragraph, allowSectionProperties: true);

        private static LegacyDocWritableParagraphFormatting ReadSupportedStyleParagraphFormatting(LegacyDocElement paragraphProperties) =>
            ReadSupportedParagraphFormattingCore(paragraphProperties, EmptyStyleIndexes, allowParagraphStyleId: false, kind: LegacyDocParagraphPropertiesKind.Style);

        private static LegacyDocWritableParagraphFormatting ReadSupportedCustomParagraphStyleParagraphFormatting(LegacyDocElement paragraphProperties) =>
            ReadSupportedParagraphFormattingCore(paragraphProperties, EmptyStyleIndexes, allowParagraphStyleId: false, kind: LegacyDocParagraphPropertiesKind.Style, allowDirectOutlineLevel: true);

        private static LegacyDocWritableParagraphFormatting ReadSupportedBuiltInStyleParagraphFormatting(ushort styleIndex, LegacyDocElement paragraphProperties) =>
            ReadSupportedParagraphFormattingCore(paragraphProperties, EmptyStyleIndexes, allowParagraphStyleId: false, kind: LegacyDocParagraphPropertiesKind.Style, builtInStyleIndex: styleIndex);

        private static LegacyDocWritableParagraphFormatting ReadSupportedParagraphFormattingCore(LegacyDocElement paragraphProperties, IReadOnlyDictionary<string, ushort> styleIndexes, bool allowParagraphStyleId, LegacyDocParagraphPropertiesKind kind, ushort? builtInStyleIndex = null, bool allowSectionProperties = false, bool allowDirectOutlineLevel = false) {
            if (!paragraphProperties.HasElements) {
                return LegacyDocWritableParagraphFormatting.Plain;
            }

            byte? alignment = null;
            int? spacingBeforeTwips = null;
            int? spacingAfterTwips = null;
            int? lineSpacingTwips = null;
            bool lineSpacingIsMultiple = false;
            int? leftIndentTwips = null;
            int? rightIndentTwips = null;
            int? firstLineIndentTwips = null;
            bool? keepLinesTogether = null;
            bool? keepWithNext = null;
            bool? pageBreakBefore = null;
            bool? avoidWidowAndOrphan = null;
            bool? suppressLineNumbers = null;
            bool? suppressAutoHyphens = null;
            bool? contextualSpacing = null;
            bool? mirrorIndents = null;
            bool? kinsoku = null;
            bool? wordWrap = null;
            bool? overflowPunctuation = null;
            bool? topLinePunctuation = null;
            bool? autoSpaceDE = null;
            bool? autoSpaceDN = null;
            bool? bidirectional = null;
            ushort? numberingListIndex = null;
            byte? numberingLevel = null;
            byte? verticalCharacterAlignment = null;
            byte? paragraphOutlineLevel = null;
            LegacyDocParagraphShading? paragraphShading = null;
            LegacyDocParagraphBorders? paragraphBorders = null;
            IReadOnlyList<LegacyDocTabStop>? tabStops = null;
            ushort? styleIndex = null;
            foreach (LegacyDocElement property in paragraphProperties.Elements()) {
                // A style's properties have no style reference, mark run properties or section properties, and a
                // list level's have no mark run properties or section properties: there they are unsupported elements.
                string? propertyName = property.WName;
                if ((propertyName == "pStyle" && kind == LegacyDocParagraphPropertiesKind.Style)
                    || ((propertyName == "rPr" || propertyName == "sectPr") && kind != LegacyDocParagraphPropertiesKind.Paragraph)) {
                    propertyName = null;
                }

                switch (propertyName) {
                    case "pStyle":
                        if (!allowParagraphStyleId) {
                            throw new NotSupportedException("Native DOC saving does not support nested paragraph style identifiers inside custom paragraph style definitions.");
                        }

                        styleIndex = ReadSupportedParagraphStyleIndex(property, styleIndexes);
                        break;
                    case "jc":
                        alignment = ReadSupportedParagraphAlignment(property);
                        break;
                    case "spacing":
                        ReadSupportedParagraphSpacing(property, out spacingBeforeTwips, out spacingAfterTwips, out lineSpacingTwips, out lineSpacingIsMultiple);
                        break;
                    case "ind":
                        ReadSupportedParagraphIndentation(property, out leftIndentTwips, out rightIndentTwips, out firstLineIndentTwips);
                        break;
                    case "keepLines":
                        keepLinesTogether = ReadOnOffValue(property);
                        break;
                    case "keepNext":
                        keepWithNext = ReadOnOffValue(property);
                        break;
                    case "pageBreakBefore":
                        pageBreakBefore = ReadOnOffValue(property);
                        break;
                    case "widowControl":
                        avoidWidowAndOrphan = ReadOnOffValue(property);
                        break;
                    case "suppressLineNumbers":
                        suppressLineNumbers = ReadOnOffValue(property);
                        break;
                    case "suppressAutoHyphens":
                        suppressAutoHyphens = ReadOnOffValue(property);
                        break;
                    case "contextualSpacing":
                        contextualSpacing = ReadOnOffValue(property);
                        break;
                    case "mirrorIndents":
                        mirrorIndents = ReadOnOffValue(property);
                        break;
                    case "kinsoku":
                        kinsoku = ReadOnOffValue(property);
                        break;
                    case "wordWrap":
                        wordWrap = ReadOnOffValue(property);
                        break;
                    case "overflowPunct":
                        overflowPunctuation = ReadOnOffValue(property);
                        break;
                    case "topLinePunct":
                        topLinePunctuation = ReadOnOffValue(property);
                        break;
                    case "autoSpaceDE":
                        autoSpaceDE = ReadOnOffValue(property);
                        break;
                    case "autoSpaceDN":
                        autoSpaceDN = ReadOnOffValue(property);
                        break;
                    case "bidi":
                        bidirectional = ReadOnOffValue(property);
                        break;
                    case "numPr":
                        ReadSupportedNumberingProperties(property, out numberingListIndex, out numberingLevel);
                        break;
                    case "textAlignment":
                        verticalCharacterAlignment = ReadSupportedVerticalCharacterAlignment(property);
                        break;
                    case "shd":
                        paragraphShading = ReadSupportedParagraphShading(property);
                        break;
                    case "pBdr":
                        paragraphBorders = ReadSupportedParagraphBorders(property);
                        break;
                    case "tabs":
                        tabStops = ReadSupportedTabStops(property);
                        break;
                    case "outlineLvl":
                        paragraphOutlineLevel = ReadSupportedOutlineLevel(property, builtInStyleIndex, allowDirectOutlineLevel: allowParagraphStyleId || allowDirectOutlineLevel);
                        break;
                    case "rPr" when allowParagraphStyleId:
                        break;
                    case "rPr":
                        throw new NotSupportedException("Native DOC saving supports paragraph mark run formatting only on direct paragraph properties. Unsupported paragraph property: rPr.");
                    case "sectPr" when allowSectionProperties:
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving currently supports only built-in paragraph styles, simple numbering, alignment, vertical character alignment, East Asian typography flags, bidirectional layout, mirror indents, contextual spacing, suppressing automatic hyphenation, spacing, indentation, pagination flags, tab stops, palette-backed paragraph shading, and palette-backed paragraph borders. Unsupported paragraph property: {property.LocalName}.");
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
                suppressAutoHyphens,
                contextualSpacing,
                mirrorIndents,
                kinsoku,
                wordWrap,
                overflowPunctuation,
                topLinePunctuation,
                autoSpaceDE,
                autoSpaceDN,
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
                paragraphBorders,
                outlineLevel: paragraphOutlineLevel,
                lineSpacingIsMultiple: lineSpacingIsMultiple);
        }

        private static byte ReadSupportedVerticalCharacterAlignment(LegacyDocElement textAlignment) {
            string value = textAlignment.Val ?? "auto";
            if (value == "auto") {
                return 4;
            } else if (value == "baseline") {
                return 2;
            } else if (value == "top") {
                return 0;
            } else if (value == "center") {
                return 1;
            } else if (value == "bottom") {
                return 3;
            }

            throw new NotSupportedException($"Native DOC saving does not support paragraph vertical character alignment '{FormatSdkEnumValue(value, "VerticalTextAlignmentValues")}'.");
        }

        private static void ReadSupportedNumberingProperties(LegacyDocElement numberingProperties, out ushort? numberingListIndex, out byte? numberingLevel) {
            numberingListIndex = null;
            numberingLevel = null;

            foreach (LegacyDocElement child in numberingProperties.Elements()) {
                switch (child.WName) {
                    case "numId":
                    case "ilvl":
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple paragraph numbering only. Unsupported numbering property: {child.LocalName}.");
                }
            }

            int? numberId = LegacyDocValues.Int32(numberingProperties.Child("numId").Val);
            if (numberId == null || numberId.Value == 0) {
                return;
            }

            if (numberId.Value < 0 || numberId.Value > ushort.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports paragraph numbering ids only within the Word 97-2003 list index range.");
            }

            int level = LegacyDocValues.Int32(numberingProperties.Child("ilvl").Val) ?? 0;
            if (level < 0 || level > 8) {
                throw new NotSupportedException("Native DOC saving supports paragraph numbering levels 0 through 8.");
            }

            numberingListIndex = checked((ushort)numberId.Value);
            numberingLevel = checked((byte)level);
        }

        private static LegacyDocParagraphBorders ReadSupportedParagraphBorders(LegacyDocElement borders) {
            foreach (LegacyDocElement child in borders.Elements()) {
                switch (child.WName) {
                    case "top":
                    case "left":
                    case "bottom":
                    case "right":
                    case "between":
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple paragraph borders only. Unsupported paragraph border: {child.LocalName}.");
                }
            }

            return new LegacyDocParagraphBorders(
                ReadSupportedParagraphBorder(borders.Child("top")),
                ReadSupportedParagraphBorder(borders.Child("left")),
                ReadSupportedParagraphBorder(borders.Child("bottom")),
                ReadSupportedParagraphBorder(borders.Child("right")),
                ReadSupportedParagraphBorder(borders.Child("between")));
        }

        private static LegacyDocParagraphBorder ReadSupportedParagraphBorder(LegacyDocElement border) {
            if (border.IsNone) {
                return default;
            }

            string? value = border.Val;
            if (value == null || value == "none" || value == "nil") {
                return default;
            }

            LegacyDocParagraphBorderStyle style = MapSupportedParagraphBorderStyle(value);
            string? colorHex = border.Attr("color");
            if (string.Equals(colorHex, "auto", StringComparison.OrdinalIgnoreCase)) {
                colorHex = null;
            }

            if (!LegacyDocPaletteProjection.CanResolveColor(colorHex)) {
                throw new NotSupportedException("Native DOC saving supports paragraph borders only with Word 97-2003 palette colors.");
            }

            uint? borderSize = LegacyDocValues.UInt32(border.Attr("sz"));
            int size = borderSize == null ? 4 : checked((int)borderSize.Value);
            uint? borderSpace = LegacyDocValues.UInt32(border.Attr("space"));
            int space = borderSpace == null ? 0 : checked((int)borderSpace.Value);
            if (size <= 0 || size > byte.MaxValue || space < 0 || space > byte.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports paragraph border size and spacing only within Word 97-2003 BRC80 byte ranges.");
            }

            return new LegacyDocParagraphBorder(style, colorHex, size, space);
        }

        private static LegacyDocParagraphBorderStyle MapSupportedParagraphBorderStyle(string value) {
            if (value == "single") {
                return LegacyDocParagraphBorderStyle.Single;
            }

            if (value == "double") {
                return LegacyDocParagraphBorderStyle.Double;
            }

            if (value == "dotted") {
                return LegacyDocParagraphBorderStyle.Dotted;
            }

            if (value == "dashed" || value == "dashSmallGap") {
                return LegacyDocParagraphBorderStyle.Dashed;
            }

            throw new NotSupportedException($"Native DOC saving does not support paragraph border style '{FormatSdkEnumValue(value, "BorderValues")}'.");
        }

        private static LegacyDocParagraphShading ReadSupportedParagraphShading(LegacyDocElement shading) {
            string? pattern = shading.Val;
            if (pattern != null && pattern != "clear") {
                throw new NotSupportedException("Native DOC saving supports paragraph shading only for clear fill patterns.");
            }

            string? fillColorHex = shading.Attr("fill");
            if (string.IsNullOrWhiteSpace(fillColorHex)) {
                return default;
            }

            // An explicit automatic fill is retained because it clears inherited shading.
            var paragraphShading = new LegacyDocParagraphShading(fillColorHex);
            if (!paragraphShading.IsAutomatic && !paragraphShading.TryGetRgb(out _, out _, out _)) {
                throw new NotSupportedException("Native DOC saving supports paragraph shading only for six-digit RGB or automatic fill colors.");
            }

            return paragraphShading;
        }

        private static void ReadSupportedBuiltInOutlineLevel(LegacyDocElement outlineLevel, ushort? builtInStyleIndex) {
            _ = ReadSupportedOutlineLevel(outlineLevel, builtInStyleIndex, allowDirectOutlineLevel: false);
        }

        private static byte? ReadSupportedOutlineLevel(LegacyDocElement outlineLevel, ushort? builtInStyleIndex, bool allowDirectOutlineLevel) {
            if (builtInStyleIndex == null || builtInStyleIndex.Value == 0 || builtInStyleIndex.Value > 9) {
                if (allowDirectOutlineLevel) {
                    int directLevel = LegacyDocValues.Int32(outlineLevel.Val) ?? 9;
                    if (directLevel >= 0 && directLevel <= 9) {
                        return checked((byte)directLevel);
                    }
                }

                throw new NotSupportedException("Native DOC saving currently supports outline-level style formatting only for built-in Heading styles.");
            }

            int expectedLevel = builtInStyleIndex.Value - 1;
            int level = LegacyDocValues.Int32(outlineLevel.Val) ?? expectedLevel;
            if (level != expectedLevel) {
                throw new NotSupportedException($"Native DOC saving supports built-in Heading style outline levels only when they match the heading level. Expected outline level {expectedLevel}, got {level}.");
            }

            return null;
        }

        private static IReadOnlyList<LegacyDocTabStop> ReadSupportedTabStops(LegacyDocElement tabs) {
            var tabStops = new List<LegacyDocTabStop>();
            foreach (LegacyDocElement tabStop in tabs.Elements("tab")) {
                string? position = tabStop.Attr("pos");
                if (position == null) {
                    throw new NotSupportedException("Native DOC saving supports tab stops only when each tab stop has an explicit twip position.");
                }

                string? tabAlignment = tabStop.Val;
                if (!TryReadSupportedTabAlignment(tabAlignment ?? "left", out LegacyDocTabStopAlignment alignment)) {
                    throw new NotSupportedException($"Native DOC saving does not support tab stop alignment '{FormatSdkEnumValue(tabAlignment, "TabStopValues")}'.");
                }

                string? tabLeader = tabStop.Attr("leader");
                if (!TryReadSupportedTabLeader(tabLeader ?? "none", out LegacyDocTabStopLeader leader)) {
                    throw new NotSupportedException($"Native DOC saving does not support tab stop leader '{FormatSdkEnumValue(tabLeader, "TabStopLeaderCharValues")}'.");
                }

                tabStops.Add(new LegacyDocTabStop(LegacyDocValues.XmlInt32(position), alignment, leader));
            }

            return tabStops;
        }

        private static bool TryReadSupportedTabAlignment(string value, out LegacyDocTabStopAlignment alignment) {
            if (value == "left") {
                alignment = LegacyDocTabStopAlignment.Left;
                return true;
            } else if (value == "center") {
                alignment = LegacyDocTabStopAlignment.Center;
                return true;
            } else if (value == "right") {
                alignment = LegacyDocTabStopAlignment.Right;
                return true;
            } else if (value == "decimal") {
                alignment = LegacyDocTabStopAlignment.Decimal;
                return true;
            } else if (value == "bar") {
                alignment = LegacyDocTabStopAlignment.Bar;
                return true;
            } else if (value == "num") {
                alignment = LegacyDocTabStopAlignment.Number;
                return true;
            } else if (value == "clear") {
                alignment = LegacyDocTabStopAlignment.Clear;
                return true;
            }

            alignment = LegacyDocTabStopAlignment.Left;
            return false;
        }

        private static bool TryReadSupportedTabLeader(string value, out LegacyDocTabStopLeader leader) {
            if (value == "none") {
                leader = LegacyDocTabStopLeader.None;
                return true;
            } else if (value == "dot") {
                leader = LegacyDocTabStopLeader.Dot;
                return true;
            } else if (value == "hyphen") {
                leader = LegacyDocTabStopLeader.Hyphen;
                return true;
            } else if (value == "underscore") {
                leader = LegacyDocTabStopLeader.Underscore;
                return true;
            } else if (value == "heavy") {
                leader = LegacyDocTabStopLeader.Heavy;
                return true;
            } else if (value == "middleDot") {
                leader = LegacyDocTabStopLeader.MiddleDot;
                return true;
            }

            leader = LegacyDocTabStopLeader.None;
            return false;
        }

        private static ushort? ReadSupportedParagraphStyleIndex(LegacyDocElement paragraphStyleId, IReadOnlyDictionary<string, ushort> styleIndexes) {
            string? styleId = paragraphStyleId.Val;
            if (string.IsNullOrWhiteSpace(styleId)
                || string.Equals(styleId, "Normal", StringComparison.Ordinal)) {
                return null;
            }

            if (TryMapBuiltInParagraphStyleIndex(styleId!, out ushort styleIndex)) {
                return styleIndex;
            }

            if (styleIndexes.TryGetValue(styleId!, out ushort customStyleIndex)) {
                return customStyleIndex;
            }

            throw new NotSupportedException($"Native DOC saving cannot resolve paragraph style '{styleId}' in the document style sheet.");
        }

        private static bool TryMapBuiltInParagraphStyleIndex(string styleId, out ushort styleIndex) {
            switch (styleId) {
                case "Heading1":
                    styleIndex = 1;
                    return true;
                case "Heading2":
                    styleIndex = 2;
                    return true;
                case "Heading3":
                    styleIndex = 3;
                    return true;
                case "Heading4":
                    styleIndex = 4;
                    return true;
                case "Heading5":
                    styleIndex = 5;
                    return true;
                case "Heading6":
                    styleIndex = 6;
                    return true;
                case "Heading7":
                    styleIndex = 7;
                    return true;
                case "Heading8":
                    styleIndex = 8;
                    return true;
                case "Heading9":
                    styleIndex = 9;
                    return true;
                default:
                    styleIndex = 0;
                    return false;
            }
        }

        private static byte? ReadSupportedParagraphAlignment(LegacyDocElement justification) {
            string value = justification.Val ?? "left";
            if (value == "left") {
                return 0;
            } else if (value == "center") {
                return 1;
            } else if (value == "right") {
                return 2;
            } else if (value == "both") {
                return 3;
            }

            throw new NotSupportedException($"Native DOC saving does not support paragraph alignment '{FormatSdkEnumValue(value, "JustificationValues")}'.");
        }

        private static void ReadSupportedParagraphSpacing(
            LegacyDocElement spacing,
            out int? spacingBeforeTwips,
            out int? spacingAfterTwips,
            out int? lineSpacingTwips,
            out bool lineSpacingIsMultiple) {
            if ((LegacyDocValues.OnOff(spacing.Attr("beforeAutospacing")) ?? false) || (LegacyDocValues.OnOff(spacing.Attr("afterAutospacing")) ?? false)) {
                throw new NotSupportedException("Native DOC saving currently supports paragraph spacing only as explicit twip values, not automatic spacing.");
            }

            if (spacing.Attr("beforeLines") != null || spacing.Attr("afterLines") != null) {
                throw new NotSupportedException("Native DOC saving currently supports paragraph before/after spacing only as twip values, not line-count spacing.");
            }

            spacingBeforeTwips = ReadOptionalInt32Twips(spacing.Attr("before"), "paragraph spacing before");
            spacingAfterTwips = ReadOptionalInt32Twips(spacing.Attr("after"), "paragraph spacing after");
            lineSpacingTwips = ReadOptionalInt32Twips(spacing.Attr("line"), "paragraph line spacing");
            string lineRule = spacing.Attr("lineRule") ?? "auto";
            lineSpacingIsMultiple = lineSpacingTwips.HasValue && lineRule == "auto";
            if (lineSpacingTwips.HasValue) {
                if (lineSpacingTwips < 0 || lineSpacingTwips > 31680)
                    throw new NotSupportedException("Native DOC line spacing must be between 0 and 31680.");
                if (lineRule == "exact") {
                    // DOC represents exact spacing with a negative LSPD value;
                    // zero would instead mean at-least single spacing.
                    if (lineSpacingTwips.Value == 0)
                        throw new NotSupportedException("Native DOC saving cannot represent exact zero line spacing. Use a positive exact spacing value.");
                    lineSpacingTwips = -lineSpacingTwips.Value;
                }
            }
        }

        private static void ReadSupportedParagraphIndentation(
            LegacyDocElement indentation,
            out int? leftIndentTwips,
            out int? rightIndentTwips,
            out int? firstLineIndentTwips) {
            if (indentation.Attr("leftChars") != null || indentation.Attr("rightChars") != null || indentation.Attr("firstLineChars") != null || indentation.Attr("hangingChars") != null) {
                throw new NotSupportedException("Native DOC saving currently supports paragraph indentation only as twip values, not character-based indentation.");
            }

            if (indentation.Attr("firstLine") != null && indentation.Attr("hanging") != null) {
                throw new NotSupportedException("Native DOC saving cannot write a paragraph with both first-line and hanging indentation.");
            }

            leftIndentTwips = ReadOptionalInt32Twips(indentation.Attr("left"), "paragraph left indentation");
            rightIndentTwips = ReadOptionalInt32Twips(indentation.Attr("right"), "paragraph right indentation");
            int? firstLine = ReadOptionalInt32Twips(indentation.Attr("firstLine"), "paragraph first-line indentation");
            int? hanging = ReadOptionalInt32Twips(indentation.Attr("hanging"), "paragraph hanging indentation");
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

        private static bool? ReadOnOffValue(LegacyDocElement property) => property.Val is { } value ? LegacyDocValues.ParseOnOff(value) : true;
    }
}
