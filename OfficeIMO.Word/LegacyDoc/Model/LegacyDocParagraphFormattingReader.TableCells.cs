namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static partial class LegacyDocParagraphFormattingReader {
        private static bool TryReadTableDefinition(
            byte[] bytes,
            int sprmOffset,
            int end,
            out IReadOnlyList<int>? tableCellWidthsTwips,
            out int? tableLeftIndentTwips,
            out IReadOnlyList<LegacyDocTableCellHorizontalMerge>? tableCellHorizontalMerges,
            out IReadOnlyList<LegacyDocTableCellVerticalMerge>? tableCellVerticalMerges,
            out IReadOnlyList<LegacyDocTableCellVerticalAlignment>? tableCellVerticalAlignments,
            out IReadOnlyList<LegacyDocTableCellTextDirection>? tableCellTextDirections,
            out IReadOnlyList<bool>? tableCellFitTexts,
            out IReadOnlyList<bool>? tableCellNoWraps,
            out IReadOnlyList<bool>? tableCellHideMarks,
            out IReadOnlyList<LegacyDocTableCellBorders>? tableCellBorders,
            out bool hasUnsupportedMergedTableCells,
            out int operandLength) {
            tableCellWidthsTwips = null;
            tableLeftIndentTwips = null;
            tableCellHorizontalMerges = null;
            tableCellVerticalMerges = null;
            tableCellVerticalAlignments = null;
            tableCellTextDirections = null;
            tableCellFitTexts = null;
            tableCellNoWraps = null;
            tableCellHideMarks = null;
            tableCellBorders = null;
            hasUnsupportedMergedTableCells = false;
            operandLength = 0;
            if (sprmOffset + 5 > end) {
                return false;
            }

            ushort cb = LegacyDocFib.ReadUInt16(bytes, sprmOffset + 2);
            operandLength = cb + 1;
            int operandOffset = sprmOffset + 2;
            int operandEnd = operandOffset + operandLength;
            if (operandEnd > end || cb < 4) {
                return false;
            }

            int columnCount = bytes[sprmOffset + 4];
            if (columnCount > 63) return false;
            if (columnCount <= 0) {
                tableCellWidthsTwips = Array.Empty<int>();
                tableLeftIndentTwips = null;
                tableCellHorizontalMerges = Array.Empty<LegacyDocTableCellHorizontalMerge>();
                tableCellVerticalMerges = Array.Empty<LegacyDocTableCellVerticalMerge>();
                tableCellVerticalAlignments = Array.Empty<LegacyDocTableCellVerticalAlignment>();
                tableCellTextDirections = Array.Empty<LegacyDocTableCellTextDirection>();
                tableCellFitTexts = Array.Empty<bool>();
                tableCellNoWraps = Array.Empty<bool>();
                tableCellHideMarks = Array.Empty<bool>();
                tableCellBorders = Array.Empty<LegacyDocTableCellBorders>();
                return true;
            }

            int edgesOffset = sprmOffset + 5;
            int tc80Offset = edgesOffset + ((columnCount + 1) * 2);
            if (tc80Offset > operandEnd || (operandEnd - tc80Offset) % Tc80Length != 0) {
                return false;
            }
            int definedCellCount = Math.Min(columnCount, (operandEnd - tc80Offset) / Tc80Length);

            var widths = new int[columnCount];
            var horizontalMerges = new LegacyDocTableCellHorizontalMerge[columnCount];
            var verticalMerges = new LegacyDocTableCellVerticalMerge[columnCount];
            var verticalAlignments = new LegacyDocTableCellVerticalAlignment[columnCount];
            var textDirections = new LegacyDocTableCellTextDirection[columnCount];
            var fitTexts = new bool[columnCount];
            var noWraps = new bool[columnCount];
            var hideMarks = new bool[columnCount];
            var borders = new LegacyDocTableCellBorders[columnCount];
            int previousEdge = ReadInt16(bytes, edgesOffset);
            if (previousEdge != 0) {
                tableLeftIndentTwips = previousEdge;
            }

            for (int index = 0; index < columnCount; index++) {
                int nextEdge = ReadInt16(bytes, edgesOffset + ((index + 1) * 2));
                int width = nextEdge - previousEdge;
                if (width < 0) return false;
                widths[index] = width;
                previousEdge = nextEdge;
            }

            for (int index = 0; index < definedCellCount; index++) {
                ushort tcgrf = LegacyDocFib.ReadUInt16(bytes, tc80Offset + (index * Tc80Length));
                switch (tcgrf & TcgrfHorizontalMergeMask) {
                    case 0:
                        horizontalMerges[index] = LegacyDocTableCellHorizontalMerge.None;
                        break;
                    case 0x0001:
                        horizontalMerges[index] = LegacyDocTableCellHorizontalMerge.Restart;
                        break;
                    case 0x0002:
                        horizontalMerges[index] = LegacyDocTableCellHorizontalMerge.Continue;
                        break;
                    default:
                        horizontalMerges[index] = LegacyDocTableCellHorizontalMerge.None;
                        hasUnsupportedMergedTableCells = true;
                        break;
                }

                switch (tcgrf & TcgrfVerticalMergeMask) {
                    case 0:
                        verticalMerges[index] = LegacyDocTableCellVerticalMerge.None;
                        break;
                    case 0x0020:
                        verticalMerges[index] = LegacyDocTableCellVerticalMerge.Restart;
                        break;
                    case 0x0040:
                        verticalMerges[index] = LegacyDocTableCellVerticalMerge.Continue;
                        break;
                    default:
                        verticalMerges[index] = LegacyDocTableCellVerticalMerge.None;
                        hasUnsupportedMergedTableCells = true;
                        break;
                }

                switch (tcgrf & TcgrfVerticalAlignmentMask) {
                    case 0:
                        verticalAlignments[index] = LegacyDocTableCellVerticalAlignment.Top;
                        break;
                    case 0x0080:
                        verticalAlignments[index] = LegacyDocTableCellVerticalAlignment.Center;
                        break;
                    case 0x0100:
                        verticalAlignments[index] = LegacyDocTableCellVerticalAlignment.Bottom;
                        break;
                    default:
                        verticalAlignments[index] = LegacyDocTableCellVerticalAlignment.Top;
                        break;
                }

                switch ((tcgrf & TcgrfTextFlowMask) >> 2) {
                    case 0:
                        textDirections[index] = LegacyDocTableCellTextDirection.LeftToRightTopToBottom;
                        break;
                    case 1:
                        textDirections[index] = LegacyDocTableCellTextDirection.TopToBottomRightToLeft;
                        break;
                    case 3:
                        textDirections[index] = LegacyDocTableCellTextDirection.BottomToTopLeftToRight;
                        break;
                    case 4:
                        textDirections[index] = LegacyDocTableCellTextDirection.LeftToRightTopToBottomRotated;
                        break;
                    case 5:
                        textDirections[index] = LegacyDocTableCellTextDirection.TopToBottomRightToLeftRotated;
                        break;
                    default:
                        textDirections[index] = LegacyDocTableCellTextDirection.LeftToRightTopToBottom;
                        break;
                }

                fitTexts[index] = (tcgrf & TcgrfFitTextMask) != 0;
                noWraps[index] = (tcgrf & TcgrfNoWrapMask) != 0;
                hideMarks[index] = (tcgrf & TcgrfHideMarkMask) != 0;
                borders[index] = ReadTableCellBorders(bytes, tc80Offset + (index * Tc80Length));
            }

            tableCellWidthsTwips = widths;
            tableCellHorizontalMerges = horizontalMerges.Any(merge => merge != LegacyDocTableCellHorizontalMerge.None)
                ? horizontalMerges
                : Array.Empty<LegacyDocTableCellHorizontalMerge>();
            tableCellVerticalMerges = verticalMerges.Any(merge => merge != LegacyDocTableCellVerticalMerge.None)
                ? verticalMerges
                : Array.Empty<LegacyDocTableCellVerticalMerge>();
            tableCellVerticalAlignments = verticalAlignments.Any(alignment => alignment != LegacyDocTableCellVerticalAlignment.Top)
                ? verticalAlignments
                : Array.Empty<LegacyDocTableCellVerticalAlignment>();
            tableCellTextDirections = textDirections.Any(textDirection => textDirection != LegacyDocTableCellTextDirection.LeftToRightTopToBottom)
                ? textDirections
                : Array.Empty<LegacyDocTableCellTextDirection>();
            tableCellFitTexts = fitTexts.Any(fitText => fitText)
                ? fitTexts
                : Array.Empty<bool>();
            tableCellNoWraps = noWraps.Any(noWrap => noWrap)
                ? noWraps
                : Array.Empty<bool>();
            tableCellHideMarks = hideMarks.Any(hideMark => hideMark)
                ? hideMarks
                : Array.Empty<bool>();
            tableCellBorders = borders.Any(border => border.HasAny)
                ? borders
                : Array.Empty<LegacyDocTableCellBorders>();
            return true;
        }

        private static LegacyDocTableCellBorders ReadTableCellBorders(byte[] bytes, int tc80Offset) {
            return new LegacyDocTableCellBorders(
                ReadBrc80(bytes, tc80Offset + 4),
                ReadBrc80(bytes, tc80Offset + 8),
                ReadBrc80(bytes, tc80Offset + 12),
                ReadBrc80(bytes, tc80Offset + 16));
        }

        private static LegacyDocTableCellBorder ReadBrc80(byte[] bytes, int offset) {
            if (offset + 4 > bytes.Length) {
                return default;
            }

            if (bytes[offset] == 0xFF
                && bytes[offset + 1] == 0xFF
                && bytes[offset + 2] == 0xFF
                && bytes[offset + 3] == 0xFF) {
                return new LegacyDocTableCellBorder(LegacyDocTableCellBorderStyle.ExplicitNone, null, 0, 0);
            }

            byte sizeEighthPoints = bytes[offset];
            byte borderType = bytes[offset + 1];
            byte colorIndex = bytes[offset + 2];
            byte spacePoints = bytes[offset + 3];
            LegacyDocTableCellBorderStyle style = MapBrc80BorderStyle(borderType);
            if (style == LegacyDocTableCellBorderStyle.None) {
                return default;
            }

            string? colorHex = LegacyDocColorPalette.GetHexForIco(colorIndex);
            return new LegacyDocTableCellBorder(style, colorHex, sizeEighthPoints, spacePoints);
        }

        private static LegacyDocTableCellBorderStyle MapBrc80BorderStyle(byte borderType) {
            switch (borderType) {
                case 0x01:
                    return LegacyDocTableCellBorderStyle.Single;
                case 0x03:
                    return LegacyDocTableCellBorderStyle.Double;
                case 0x06:
                    return LegacyDocTableCellBorderStyle.Dotted;
                case 0x07:
                    return LegacyDocTableCellBorderStyle.Dashed;
                default:
                    return LegacyDocTableCellBorderStyle.None;
            }
        }

        private static bool TryReadTableCellPadding(
            byte[] bytes,
            int sprmOffset,
            int end,
            bool isDefault,
            ref IReadOnlyList<LegacyDocTableCellMargins>? tableCellMargins,
            ref LegacyDocTableCellMargins? defaultTableCellMargins,
            out int operandLength) {
            operandLength = 0;
            if (sprmOffset + 9 > end) {
                return false;
            }

            int cb = bytes[sprmOffset + 2];
            operandLength = 1 + cb;
            if (cb != 6 || sprmOffset + 2 + operandLength > end) {
                return false;
            }

            int itcFirst = bytes[sprmOffset + 3];
            int itcLim = bytes[sprmOffset + 4];
            byte grfbrc = bytes[sprmOffset + 5];
            byte ftsWidth = bytes[sprmOffset + 6];
            int width = LegacyDocFib.ReadUInt16(bytes, sprmOffset + 7);
            if (ftsWidth != FtsDxa || width < 0 || width > 31680) {
                return true;
            }

            LegacyDocTableCellMargins margins = CreateTableCellMargins(grfbrc, width);
            if (!margins.HasAny) {
                return true;
            }

            if (isDefault) {
                defaultTableCellMargins = (defaultTableCellMargins ?? default).Merge(margins);
                return true;
            }

            if (itcFirst >= itcLim) {
                return true;
            }

            LegacyDocTableCellMargins[] marginArray;
            if (tableCellMargins == null || tableCellMargins.Count < itcLim) {
                marginArray = new LegacyDocTableCellMargins[itcLim];
                if (tableCellMargins != null) {
                    for (int index = 0; index < tableCellMargins.Count; index++) {
                        marginArray[index] = tableCellMargins[index];
                    }
                }
            } else {
                marginArray = tableCellMargins.ToArray();
            }
            for (int index = itcFirst; index < itcLim; index++) {
                marginArray[index] = marginArray[index].Merge(margins);
            }

            tableCellMargins = marginArray.Any(margin => margin.HasAny)
                ? marginArray
                : Array.Empty<LegacyDocTableCellMargins>();
            return true;
        }

        private static bool TryReadTableCellSpacing(
            byte[] bytes,
            int sprmOffset,
            int end,
            ref int? defaultTableCellSpacingTwips,
            out int operandLength) {
            operandLength = 0;
            if (sprmOffset + 9 > end) {
                return false;
            }

            int cb = bytes[sprmOffset + 2];
            operandLength = 1 + cb;
            if (cb != 6 || sprmOffset + 2 + operandLength > end) {
                return false;
            }

            byte ftsWidth = bytes[sprmOffset + 6];
            int width = LegacyDocFib.ReadUInt16(bytes, sprmOffset + 7);
            if (ftsWidth == FtsDxa && width >= 0 && width <= 31680) {
                defaultTableCellSpacingTwips = width;
            }

            return true;
        }

        private static bool TryReadTableCellShadings(
            byte[] bytes,
            int sprmOffset,
            int end,
            out IReadOnlyList<LegacyDocTableCellShading>? tableCellShadings,
            out int operandLength) {
            tableCellShadings = null;
            operandLength = 0;
            if (sprmOffset + 3 > end) {
                return false;
            }

            int cb = bytes[sprmOffset + 2];
            operandLength = 1 + cb;
            if (cb % 2 != 0 || sprmOffset + 2 + operandLength > end) {
                return false;
            }

            int cellCount = cb / 2;
            var shadings = new LegacyDocTableCellShading[cellCount];
            for (int index = 0; index < cellCount; index++) {
                ushort shd80 = LegacyDocFib.ReadUInt16(bytes, sprmOffset + 3 + (index * 2));
                shadings[index] = ReadTableCellShading(shd80);
            }

            tableCellShadings = shadings.Any(shading => shading.HasAny)
                ? shadings
                : Array.Empty<LegacyDocTableCellShading>();
            return true;
        }

        private static LegacyDocTableCellShading ReadTableCellShading(ushort shd80) {
            if (shd80 == 0 || shd80 == Shd80Nil) {
                return default;
            }

            byte backgroundIco = (byte)((shd80 >> 5) & 0x1F);
            string? fillColorHex = LegacyDocColorPalette.GetHexForIco(backgroundIco);
            return string.IsNullOrEmpty(fillColorHex)
                ? default
                : new LegacyDocTableCellShading(fillColorHex);
        }

        private static LegacyDocTableCellMargins CreateTableCellMargins(byte sideMask, int widthTwips) {
            return new LegacyDocTableCellMargins(
                (sideMask & FbrcTop) != 0 ? widthTwips : null,
                (sideMask & FbrcRight) != 0 ? widthTwips : null,
                (sideMask & FbrcBottom) != 0 ? widthTwips : null,
                (sideMask & FbrcLeft) != 0 ? widthTwips : null);
        }

    }
}
