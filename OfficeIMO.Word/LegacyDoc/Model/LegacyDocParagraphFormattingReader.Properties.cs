namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static partial class LegacyDocParagraphFormattingReader {
        internal static LegacyDocParagraphFormat ReadGrpprl(byte[] bytes, int offset, int count, ushort? baseStyleIndex = null, bool requireComplete = false) {
            int end = offset + count;
            LegacyDocParagraphAlignment? alignment = null;
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
            byte? outlineLevel = null;
            LegacyDocParagraphShading? paragraphShading = null;
            LegacyDocParagraphBorder paragraphTopBorder = default;
            LegacyDocParagraphBorder paragraphLeftBorder = default;
            LegacyDocParagraphBorder paragraphBottomBorder = default;
            LegacyDocParagraphBorder paragraphRightBorder = default;
            LegacyDocParagraphBorder paragraphBetweenBorder = default;
            bool? isInTable = null;
            bool? isTableTerminatingParagraph = null;
            int tableDepth = 0;
            int maximumTableDepth = 0;
            bool hasNestedTable = false;
            bool hasInnerTableCellMarker = false;
            bool hasInnerTableTerminatingParagraphMarker = false;
            var tabStops = new List<LegacyDocTabStop>();
            IReadOnlyList<int>? tableCellWidthsTwips = null;
            int? tableLeftIndentTwips = null;
            int? tableOriginTwips = null;
            int? tableGapHalfTwips = null;
            IReadOnlyList<LegacyDocTableCellHorizontalMerge>? tableCellHorizontalMerges = null;
            IReadOnlyList<LegacyDocTableCellVerticalMerge>? tableCellVerticalMerges = null;
            IReadOnlyList<LegacyDocTableCellVerticalAlignment>? tableCellVerticalAlignments = null;
            IReadOnlyList<LegacyDocTableCellTextDirection>? tableCellTextDirections = null;
            IReadOnlyList<bool>? tableCellFitTexts = null;
            IReadOnlyList<bool>? tableCellNoWraps = null;
            IReadOnlyList<bool>? tableCellHideMarks = null;
            IReadOnlyList<LegacyDocTableCellMargins>? tableCellMargins = null;
            IReadOnlyList<LegacyDocTableCellShading>? tableCellShadings = null;
            IReadOnlyList<LegacyDocTableCellBorders>? tableCellBorders = null;
            LegacyDocTableCellMargins? defaultTableCellMargins = null;
            int? defaultTableCellSpacingTwips = null;
            int? tableRowHeightTwips = null;
            bool tableRowHeightIsExact = false;
            bool? tableRowCantSplit = null;
            bool? tableRowIsHeader = null;
            LegacyDocTableAlignment? tableAlignment = null;
            LegacyDocTablePreferredWidth? tablePreferredWidth = null;
            bool? tableAutofit = null;
            bool hasMergedTableCells = false;
            ushort? styleIndex = baseStyleIndex;
            ushort? tableStyleIndex = null;
            LegacyDocTableBorders tableBorders = default;
            while (offset + 2 <= end) {
                ushort sprm = LegacyDocFib.ReadUInt16(bytes, offset);
                if (sprm == SprmTInsert || sprm == SprmTDelete) {
                    if (!TryChangeTableCellCount(bytes, offset, end, sprm,
                        ref tableCellWidthsTwips, ref tableCellHorizontalMerges, ref tableCellVerticalMerges,
                        ref tableCellVerticalAlignments, ref tableCellTextDirections, ref tableCellFitTexts,
                        ref tableCellNoWraps, ref tableCellHideMarks, ref tableCellMargins,
                        ref tableCellShadings, ref tableCellBorders)) break;
                    offset += sprm == SprmTInsert ? 6 : 4;
                    continue;
                }
                if (sprm == SprmTDxaCol) {
                    if (!TryChangeTableColumnWidths(bytes, offset, end, ref tableCellWidthsTwips)) break;
                    offset += 6;
                    continue;
                }
                if (sprm == SprmTDxaLeft || sprm == SprmTDxaGapHalf) {
                    if (end - offset < 4) break;
                    int value = ReadInt16(bytes, offset + 2);
                    if (sprm == SprmTDxaLeft) tableOriginTwips = value;
                    else if (value >= 0) tableGapHalfTwips = value;
                    else break;
                    tableLeftIndentTwips = (tableOriginTwips ?? 0) - (tableGapHalfTwips ?? 0);
                    offset += 4;
                    continue;
                }
                if (sprm == SprmTIstd) {
                    if (offset + 4 > end) break;
                    tableStyleIndex = LegacyDocFib.ReadUInt16(bytes, offset + 2);
                    tableBorders = default;
                    offset += 4;
                    continue;
                }
                if (sprm == SprmTTableBorders || sprm == SprmTTableBorders80) {
                    if (!TryReadTableBorders(bytes, offset, end, sprm == SprmTTableBorders80, out tableBorders)) break;
                    offset += 3 + bytes[offset + 2];
                    continue;
                }
                if (sprm >= SprmTBrcTopCv && sprm <= SprmTBrcRightCv) {
                    if (!TryReadTableCellBorderColors(bytes, offset, end, sprm, ref tableCellBorders, out int colorLength)) break;
                    offset += 2 + colorLength;
                    continue;
                }
                if (sprm == SprmTSetBrc) {
                    if (!TryReadTableCellBorderRange(bytes, offset, end, tableCellWidthsTwips?.Count ?? 0,
                        ref tableCellBorders)) break;
                    offset += 14;
                    continue;
                }
                if (sprm == SprmPIstd) {
                    if (offset + 4 > end) {
                        break;
                    }

                    styleIndex = LegacyDocFib.ReadUInt16(bytes, offset + 2);
                    offset += 4;
                    continue;
                }

                if (sprm == SprmPFKeep
                    || sprm == SprmPFKeepFollow
                    || sprm == SprmPFPageBreakBefore
                    || sprm == SprmPFNoLineNumb
                    || sprm == SprmPFNoAutoHyph
                    || sprm == SprmPFKinsoku
                    || sprm == SprmPFWordWrap
                    || sprm == SprmPFOverflowPunct
                    || sprm == SprmPFTopLinePunct
                    || sprm == SprmPFAutoSpaceDE
                    || sprm == SprmPFAutoSpaceDN
                    || sprm == SprmPFContextualSpacing
                    || sprm == SprmPFMirrorIndents
                    || sprm == SprmPFWidowControl
                    || sprm == SprmPFBiDi
                    || sprm == SprmPFInTable
                    || sprm == SprmPFTtp
                    || sprm == SprmPFInnerTableCell
                    || sprm == SprmPFInnerTtp) {
                    if (offset + 3 > end) {
                        break;
                    }

                    bool? value = ReadBoolOperand(bytes[offset + 2]);
                    switch (sprm) {
                        case SprmPFKeep:
                            keepLinesTogether = value;
                            break;
                        case SprmPFKeepFollow:
                            keepWithNext = value;
                            break;
                        case SprmPFPageBreakBefore:
                            pageBreakBefore = value;
                            break;
                        case SprmPFWidowControl:
                            avoidWidowAndOrphan = value;
                            break;
                        case SprmPFNoLineNumb:
                            suppressLineNumbers = value;
                            break;
                        case SprmPFNoAutoHyph:
                            suppressAutoHyphens = value;
                            break;
                        case SprmPFKinsoku:
                            kinsoku = value;
                            break;
                        case SprmPFWordWrap:
                            wordWrap = value;
                            break;
                        case SprmPFOverflowPunct:
                            overflowPunctuation = value;
                            break;
                        case SprmPFTopLinePunct:
                            topLinePunctuation = value;
                            break;
                        case SprmPFAutoSpaceDE:
                            autoSpaceDE = value;
                            break;
                        case SprmPFAutoSpaceDN:
                            autoSpaceDN = value;
                            break;
                        case SprmPFContextualSpacing:
                            contextualSpacing = value;
                            break;
                        case SprmPFMirrorIndents:
                            mirrorIndents = value;
                            break;
                        case SprmPFBiDi:
                            bidirectional = value;
                            break;
                        case SprmPFInTable:
                            isInTable = value;
                            break;
                        case SprmPFTtp:
                            isTableTerminatingParagraph = value;
                            break;
                        case SprmPFInnerTableCell:
                            hasInnerTableCellMarker |= value == true;
                            hasNestedTable |= value == true;
                            break;
                        case SprmPFInnerTtp:
                            hasInnerTableTerminatingParagraphMarker |= value == true;
                            hasNestedTable |= value == true;
                            break;
                    }

                    offset += 3;
                    continue;
                }

                if (sprm == SprmPItap || sprm == SprmPDtap) {
                    if (offset + 6 > end) {
                        break;
                    }

                    int operand = LegacyDocFib.ReadInt32(bytes, offset + 2);
                    tableDepth = sprm == SprmPItap
                        ? operand
                        : tableDepth + operand;
                    if (tableDepth > maximumTableDepth) {
                        maximumTableDepth = tableDepth;
                    }

                    hasNestedTable |= tableDepth > 1;
                    offset += 6;
                    continue;
                }

                if (sprm == SprmPJc || sprm == SprmPJc80) {
                    if (offset + 3 > end) {
                        break;
                    }

                    alignment = MapAlignment(bytes[offset + 2]);
                    offset += 3;
                    continue;
                }

                if (sprm == SprmPIlvl) {
                    if (offset + 3 > end) {
                        break;
                    }

                    byte level = bytes[offset + 2];
                    if (level <= 8) {
                        numberingLevel = level;
                    }

                    offset += 3;
                    continue;
                }

                if (sprm == SprmPOutLvl) {
                    if (offset + 3 > end) {
                        break;
                    }

                    byte level = bytes[offset + 2];
                    if (level <= 9) {
                        outlineLevel = level;
                    }

                    offset += 3;
                    continue;
                }

                if (sprm == SprmPIlfo) {
                    if (offset + 4 > end) {
                        break;
                    }

                    ushort ilfo = LegacyDocFib.ReadUInt16(bytes, offset + 2);
                    if (ilfo > 0) {
                        numberingListIndex = ilfo;
                    }

                    offset += 4;
                    continue;
                }

                if (sprm == SprmPWAlignFont) {
                    if (offset + 4 > end) {
                        break;
                    }

                    ushort verticalAlignment = LegacyDocFib.ReadUInt16(bytes, offset + 2);
                    if (verticalAlignment <= 4) {
                        verticalCharacterAlignment = (byte)verticalAlignment;
                    }

                    offset += 4;
                    continue;
                }

                if (sprm == SprmTFCantSplit || sprm == SprmTFCantSplit90 || sprm == SprmTTableHeader) {
                    if (offset + 3 > end) {
                        break;
                    }

                    bool? value = ReadBoolOperand(bytes[offset + 2]);
                    if (sprm == SprmTTableHeader) {
                        tableRowIsHeader = value;
                    } else {
                        tableRowCantSplit = value;
                    }

                    offset += 3;
                    continue;
                }

                if (sprm == SprmTFAutofit) {
                    if (offset + 3 > end) {
                        break;
                    }

                    tableAutofit = bytes[offset + 2] != 0;
                    offset += 3;
                    continue;
                }

                if (sprm == SprmTJc) {
                    if (offset + 4 > end) {
                        break;
                    }

                    tableAlignment = MapTableAlignment(LegacyDocFib.ReadUInt16(bytes, offset + 2));
                    offset += 4;
                    continue;
                }

                if (sprm == SprmTTableWidth) {
                    if (offset + 5 > end) {
                        break;
                    }

                    tablePreferredWidth = ReadTablePreferredWidth(bytes[offset + 2], LegacyDocFib.ReadUInt16(bytes, offset + 3));
                    offset += 5;
                    continue;
                }

                if (sprm == SprmPDxaLeft || sprm == SprmPDxaRight || sprm == SprmPDxaLeft1 || sprm == SprmPDyaBefore || sprm == SprmPDyaAfter) {
                    if (offset + 4 > end) {
                        break;
                    }

                    int value = ReadInt16(bytes, offset + 2);
                    switch (sprm) {
                        case SprmPDxaLeft:
                            leftIndentTwips = value;
                            break;
                        case SprmPDxaRight:
                            rightIndentTwips = value;
                            break;
                        case SprmPDxaLeft1:
                            firstLineIndentTwips = value;
                            break;
                        case SprmPDyaBefore:
                            spacingBeforeTwips = value;
                            break;
                        case SprmPDyaAfter:
                            spacingAfterTwips = value;
                            break;
                    }

                    offset += 4;
                    continue;
                }

                if (sprm == SprmPDyaLine) {
                    if (offset + 6 > end) {
                        break;
                    }

                    int dyaLine = ReadInt16(bytes, offset + 2);
                    int fMultLinespace = ReadInt16(bytes, offset + 4);
                    if ((fMultLinespace == 0 || fMultLinespace == 1) && dyaLine >= -31680 && dyaLine <= 31680) {
                        lineSpacingTwips = dyaLine;
                        lineSpacingIsMultiple = fMultLinespace == 1 && dyaLine >= 0;
                    }

                    offset += 6;
                    continue;
                }

                if (sprm == SprmPShd80) {
                    if (offset + 4 > end) {
                        break;
                    }

                    paragraphShading = ReadParagraphShading(LegacyDocFib.ReadUInt16(bytes, offset + 2));
                    offset += 4;
                    continue;
                }

                if (sprm == SprmPBrcTop80
                    || sprm == SprmPBrcLeft80
                    || sprm == SprmPBrcBottom80
                    || sprm == SprmPBrcRight80
                    || sprm == SprmPBrcBetween80) {
                    if (offset + 6 > end) {
                        break;
                    }

                    LegacyDocParagraphBorder border = ReadParagraphBorder(bytes, offset + 2);
                    switch (sprm) {
                        case SprmPBrcTop80:
                            paragraphTopBorder = border;
                            break;
                        case SprmPBrcLeft80:
                            paragraphLeftBorder = border;
                            break;
                        case SprmPBrcBottom80:
                            paragraphBottomBorder = border;
                            break;
                        case SprmPBrcRight80:
                            paragraphRightBorder = border;
                            break;
                        case SprmPBrcBetween80:
                            paragraphBetweenBorder = border;
                            break;
                    }

                    offset += 6;
                    continue;
                }

                if (sprm == SprmPChgTabsPapx || sprm == SprmPChgTabs) {
                    if (offset + 3 > end) {
                        break;
                    }

                    int tabOperandLength = bytes[offset + 2];
                    if (offset + 3 + tabOperandLength > end) {
                        break;
                    }

                    // Word also stores list-level additions in sprmPChgTabs. With no
                    // deletions its payload matches sprmPChgTabsPapx. Range deletions
                    // need inherited-tab resolution and remain outside this projection.
                    if (sprm == SprmPChgTabsPapx || (tabOperandLength >= 2 && bytes[offset + 3] == 0)) {
                        ReadTabChanges(bytes, offset + 3, offset + 3 + tabOperandLength, tabStops);
                    }
                    offset += 3 + tabOperandLength;
                    continue;
                }

                if (sprm == SprmTDyaRowHeight) {
                    if (offset + 4 > end) {
                        break;
                    }

                    int rowHeight = ReadInt16(bytes, offset + 2);
                    if (rowHeight < 0) {
                        tableRowHeightTwips = -rowHeight;
                        tableRowHeightIsExact = true;
                    } else if (rowHeight > 0) {
                        tableRowHeightTwips = rowHeight;
                        tableRowHeightIsExact = false;
                    }

                    offset += 4;
                    continue;
                }

                if (sprm == SprmTDefTable) {
                    if (!TryReadTableDefinition(
                        bytes,
                        offset,
                        end,
                        out tableCellWidthsTwips,
                        out tableLeftIndentTwips,
                        out tableCellHorizontalMerges,
                        out tableCellVerticalMerges,
                        out tableCellVerticalAlignments,
                        out tableCellTextDirections,
                        out tableCellFitTexts,
                        out tableCellNoWraps,
                        out tableCellHideMarks,
                        out tableCellBorders,
                        out bool tableDefinitionHasUnsupportedMergedCells,
                        out int tableDefinitionOperandLength)) {
                        break;
                    }

                    hasMergedTableCells |= tableDefinitionHasUnsupportedMergedCells;
                    offset += 2 + tableDefinitionOperandLength;
                    continue;
                }

                if (sprm == SprmTDefTableShd80) {
                    if (!TryReadTableCellShadings(
                        bytes,
                        offset,
                        end,
                        out tableCellShadings,
                        out int tableCellShadingOperandLength)) {
                        break;
                    }

                    offset += 2 + tableCellShadingOperandLength;
                    continue;
                }

                if (sprm == SprmTCellPadding || sprm == SprmTCellPaddingDefault) {
                    if (!TryReadTableCellPadding(
                        bytes,
                        offset,
                        end,
                        sprm == SprmTCellPaddingDefault,
                        ref tableCellMargins,
                        ref defaultTableCellMargins,
                        out int tableCellPaddingOperandLength)) {
                        break;
                    }

                    offset += 2 + tableCellPaddingOperandLength;
                    continue;
                }

                if (sprm == SprmTCellSpacingDefault) {
                    if (!TryReadTableCellSpacing(
                        bytes,
                        offset,
                        end,
                        ref defaultTableCellSpacingTwips,
                        out int tableCellSpacingOperandLength)) {
                        break;
                    }

                    offset += 2 + tableCellSpacingOperandLength;
                    continue;
                }

                if (!TryGetSprmOperandLength(bytes, offset, end, out int operandLength)) {
                    break;
                }

                offset += 2 + operandLength;
            }

            if (requireComplete && offset != end)
                throw new InvalidDataException("Truncated native list paragraph formatting operand.");

            return new LegacyDocParagraphFormat(
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
                isInTable,
                isTableTerminatingParagraph,
                tabStops,
                tableCellWidthsTwips,
                tableLeftIndentTwips,
                tableRowHeightTwips,
                tableRowHeightIsExact,
                tableRowCantSplit,
                tableRowIsHeader,
                tableAlignment,
                tablePreferredWidth,
                tableAutofit,
                tableCellHorizontalMerges,
                tableCellVerticalMerges,
                tableCellVerticalAlignments,
                tableCellTextDirections,
                tableCellFitTexts,
                tableCellNoWraps,
                tableCellHideMarks,
                tableCellMargins,
                tableCellShadings,
                tableCellBorders,
                defaultTableCellMargins,
                defaultTableCellSpacingTwips,
                hasMergedTableCells,
                hasNestedTable,
                maximumTableDepth,
                hasInnerTableCellMarker,
                hasInnerTableTerminatingParagraphMarker,
                paragraphShading,
                new LegacyDocParagraphBorders(
                    paragraphTopBorder,
                    paragraphLeftBorder,
                    paragraphBottomBorder,
                    paragraphRightBorder,
                    paragraphBetweenBorder),
                outlineLevel,
                lineSpacingIsMultiple: lineSpacingIsMultiple,
                tableStyleIndex: tableStyleIndex,
                tableBorders: tableBorders);
        }

    }
}
