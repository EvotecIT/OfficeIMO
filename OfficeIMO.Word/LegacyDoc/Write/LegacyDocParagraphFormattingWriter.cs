namespace OfficeIMO.Word.LegacyDoc.Write {
    using OfficeIMO.Word.LegacyDoc.Model;

    internal static class LegacyDocParagraphFormattingWriter {
        private const int PapxFkpBxLength = 13;
        private const ushort SprmPFKeep = 0x2405;
        private const ushort SprmPFKeepFollow = 0x2406;
        private const ushort SprmPFPageBreakBefore = 0x2407;
        private const ushort SprmPFInTable = 0x2416;
        private const ushort SprmPFTtp = 0x2417;
        private const ushort SprmPJc = 0x2461;
        private const ushort SprmPDxaRight = 0x840E;
        private const ushort SprmPDxaLeft = 0x840F;
        private const ushort SprmPDxaLeft1 = 0x8411;
        private const ushort SprmPDyaLine = 0x6412;
        private const ushort SprmPDyaBefore = 0xA413;
        private const ushort SprmPDyaAfter = 0xA414;
        private const ushort SprmPBrcTop80 = 0x6424;
        private const ushort SprmPBrcLeft80 = 0x6425;
        private const ushort SprmPBrcBottom80 = 0x6426;
        private const ushort SprmPBrcRight80 = 0x6427;
        private const ushort SprmPBrcBetween80 = 0x6428;
        private const ushort SprmPShd80 = 0x442D;
        private const ushort SprmPFWidowControl = 0x2431;
        private const ushort SprmPChgTabsPapx = 0xC60D;
        private const ushort SprmPIlvl = 0x260A;
        private const ushort SprmPIlfo = 0x460B;
        private const ushort SprmTTableHeader = 0x3404;
        private const ushort SprmTFCantSplit90 = 0x3466;
        private const ushort SprmTJc = 0x548A;
        private const ushort SprmTDyaRowHeight = 0x9407;
        private const ushort SprmTFAutofit = 0x3615;
        private const ushort SprmTDefTable = 0xD608;
        private const ushort SprmTDefTableShd80 = 0xD609;
        private const ushort SprmTCellPadding = 0xD632;
        private const ushort SprmTCellSpacingDefault = 0xD633;
        private const ushort SprmTCellPaddingDefault = 0xD634;
        private const ushort SprmTTableWidth = 0xF614;
        private const ushort Shd80Nil = 0xFFFF;
        private const int Tc80Length = 20;
        private const byte FbrcTop = 0x01;
        private const byte FbrcLeft = 0x02;
        private const byte FbrcBottom = 0x04;
        private const byte FbrcRight = 0x08;
        private const byte FtsPercent = 0x02;
        private const byte FtsDxa = 0x03;

        internal static void WritePapxFkp(byte[] stream, int pageOffset, int textOffset, int oleSectorSize, IReadOnlyList<LegacyDocWritableParagraphSegment> segments, int bytesPerCharacter) {
            if (segments.Count == 0 || segments.Count > byte.MaxValue) {
                throw new NotSupportedException("Native DOC saving currently supports paragraph formatting only when it fits in one paragraph-format page.");
            }

            int rgbxOffset = pageOffset + ((segments.Count + 1) * 4);
            int papxOffset = oleSectorSize - 1;

            if (rgbxOffset + (segments.Count * PapxFkpBxLength) > pageOffset + papxOffset) {
                throw new NotSupportedException("Native DOC saving currently supports paragraph formatting only when it fits in one paragraph-format page.");
            }

            for (int index = 0; index < segments.Count; index++) {
                LegacyDocWritableParagraphSegment segment = segments[index];
                WriteInt32(stream, pageOffset + (index * 4), textOffset + (segment.StartCharacter * bytesPerCharacter));
                if (!segment.Formatting.HasFormatting) {
                    continue;
                }

                byte[]? papx = segment.PapxOverride;
                if (papx == null && segment.Formatting.HasFormatting) {
                    papx = CreatePapx(segment.Formatting);
                }

                if (papx != null) {
                    papxOffset -= papx.Length;
                    papxOffset = papxOffset % 2 == 0 ? papxOffset : papxOffset - 1;
                    if (pageOffset + papxOffset <= (rgbxOffset + (segments.Count * PapxFkpBxLength)) || papxOffset / 2 > byte.MaxValue) {
                        throw new NotSupportedException("Native DOC saving currently supports paragraph formatting only when it fits in one paragraph-format page.");
                    }

                    Buffer.BlockCopy(papx, 0, stream, pageOffset + papxOffset, papx.Length);
                    stream[rgbxOffset + (index * PapxFkpBxLength)] = (byte)(papxOffset / 2);
                }
            }

            LegacyDocWritableParagraphSegment lastSegment = segments[segments.Count - 1];
            WriteInt32(stream, pageOffset + (segments.Count * 4), textOffset + (lastSegment.EndCharacter * bytesPerCharacter));
            stream[pageOffset + oleSectorSize - 1] = (byte)segments.Count;
        }

        private static byte[] CreatePapx(LegacyDocWritableParagraphFormatting formatting) {
            var grpprl = new List<byte>(6) {
                (byte)((formatting.StyleIndex ?? 0) & 0xFF),
                (byte)((formatting.StyleIndex ?? 0) >> 8)
            };

            if (formatting.Alignment != null) {
                AddSingleByteSprm(grpprl, SprmPJc, formatting.Alignment.Value);
            }

            if (formatting.KeepLinesTogether == true) {
                AddSingleByteSprm(grpprl, SprmPFKeep, 1);
            }

            if (formatting.KeepWithNext == true) {
                AddSingleByteSprm(grpprl, SprmPFKeepFollow, 1);
            }

            if (formatting.PageBreakBefore == true) {
                AddSingleByteSprm(grpprl, SprmPFPageBreakBefore, 1);
            }

            if (formatting.AvoidWidowAndOrphan == true) {
                AddSingleByteSprm(grpprl, SprmPFWidowControl, 1);
            }

            if (formatting.NumberingLevel != null) {
                AddSingleByteSprm(grpprl, SprmPIlvl, formatting.NumberingLevel.Value);
            }

            if (formatting.NumberingListIndex != null) {
                AddUInt16Sprm(grpprl, SprmPIlfo, formatting.NumberingListIndex.Value);
            }

            if (formatting.IsInTable == true) {
                AddSingleByteSprm(grpprl, SprmPFInTable, 1);
            }

            if (formatting.IsTableTerminatingParagraph == true) {
                AddSingleByteSprm(grpprl, SprmPFTtp, 1);
            }

            if (formatting.TableAlignment != null) {
                AddInt16Sprm(grpprl, SprmTJc, MapTableAlignment(formatting.TableAlignment.Value));
            }

            if (formatting.TablePreferredWidth != null) {
                AddTablePreferredWidthSprm(grpprl, formatting.TablePreferredWidth.Value);
            }

            if (formatting.TableAutofit != null) {
                AddSingleByteSprm(grpprl, SprmTFAutofit, formatting.TableAutofit.Value ? (byte)1 : (byte)0);
            }

            if (formatting.LeftIndentTwips != null) {
                AddInt16Sprm(grpprl, SprmPDxaLeft, formatting.LeftIndentTwips.Value);
            }

            if (formatting.RightIndentTwips != null) {
                AddInt16Sprm(grpprl, SprmPDxaRight, formatting.RightIndentTwips.Value);
            }

            if (formatting.FirstLineIndentTwips != null) {
                AddInt16Sprm(grpprl, SprmPDxaLeft1, formatting.FirstLineIndentTwips.Value);
            }

            if (formatting.SpacingBeforeTwips != null) {
                AddInt16Sprm(grpprl, SprmPDyaBefore, formatting.SpacingBeforeTwips.Value);
            }

            if (formatting.SpacingAfterTwips != null) {
                AddInt16Sprm(grpprl, SprmPDyaAfter, formatting.SpacingAfterTwips.Value);
            }

            if (formatting.LineSpacingTwips != null) {
                AddLineSpacingSprm(grpprl, formatting.LineSpacingTwips.Value);
            }

            if (formatting.TabStops.Count > 0) {
                AddTabStopsSprm(grpprl, formatting.TabStops);
            }

            if (formatting.ParagraphShading != null && formatting.ParagraphShading.Value.HasAny) {
                AddParagraphShadingSprm(grpprl, formatting.ParagraphShading.Value);
            }

            if (formatting.ParagraphBorders != null && formatting.ParagraphBorders.Value.HasAny) {
                AddParagraphBorderSprms(grpprl, formatting.ParagraphBorders.Value);
            }

            if (formatting.TableCellWidthsTwips.Count > 0) {
                AddTableDefinitionSprm(
                    grpprl,
                    formatting.TableCellWidthsTwips,
                    formatting.TableLeftIndentTwips,
                    formatting.TableCellHorizontalMerges,
                    formatting.TableCellVerticalMerges,
                    formatting.TableCellVerticalAlignments,
                    formatting.TableCellTextDirections,
                    formatting.TableCellFitTexts,
                    formatting.TableCellNoWraps,
                    formatting.TableCellHideMarks,
                    formatting.TableCellBorders);
            }

            if (formatting.DefaultTableCellMargins != null) {
                AddDefaultTableCellPaddingSprms(grpprl, formatting.DefaultTableCellMargins.Value);
            }

            if (formatting.DefaultTableCellSpacingTwips != null) {
                AddDefaultTableCellSpacingSprm(grpprl, formatting.DefaultTableCellSpacingTwips.Value);
            }

            if (formatting.TableCellMargins.Count > 0) {
                AddTableCellPaddingSprms(grpprl, formatting.TableCellMargins);
            }

            if (formatting.TableCellShadings.Count > 0) {
                AddTableCellShadingSprm(grpprl, formatting.TableCellShadings);
            }

            if (formatting.TableRowHeightTwips != null) {
                AddTableRowHeightSprm(grpprl, formatting.TableRowHeightTwips.Value, formatting.TableRowHeightIsExact);
            }

            if (formatting.TableRowCantSplit == true) {
                AddSingleByteSprm(grpprl, SprmTFCantSplit90, 1);
            }

            if (formatting.TableRowIsHeader == true) {
                AddSingleByteSprm(grpprl, SprmTTableHeader, 1);
            }

            if (grpprl.Count % 2 != 0) {
                grpprl.Add(0);
            }

            int cb = grpprl.Count / 2;
            if (cb > byte.MaxValue) {
                throw new NotSupportedException("Native DOC saving cannot write paragraph formatting because the PAPX record is too large.");
            }

            var papx = new byte[grpprl.Count + 2];
            papx[0] = 0;
            papx[1] = (byte)cb;
            grpprl.CopyTo(papx, 2);
            return papx;
        }

        internal static byte[] CreateStyleParagraphUpx(LegacyDocWritableParagraphFormatting formatting) {
            if (!formatting.HasFormatting) {
                return Array.Empty<byte>();
            }

            byte[] papx = CreatePapx(formatting);
            var upx = new byte[papx.Length - 2];
            Buffer.BlockCopy(papx, 2, upx, 0, upx.Length);
            return upx;
        }

        private static void AddSingleByteSprm(List<byte> grpprl, ushort sprm, byte operand) {
            grpprl.Add((byte)(sprm & 0xFF));
            grpprl.Add((byte)(sprm >> 8));
            grpprl.Add(operand);
        }

        private static void AddInt16Sprm(List<byte> grpprl, ushort sprm, int operand) {
            if (operand < short.MinValue || operand > short.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports paragraph spacing and indentation only within the Word 97-2003 signed twip range.");
            }

            short value = checked((short)operand);
            grpprl.Add((byte)(sprm & 0xFF));
            grpprl.Add((byte)(sprm >> 8));
            grpprl.Add((byte)(value & 0xFF));
            grpprl.Add((byte)(value >> 8));
        }

        private static void AddUInt16Sprm(List<byte> grpprl, ushort sprm, ushort operand) {
            grpprl.Add((byte)(sprm & 0xFF));
            grpprl.Add((byte)(sprm >> 8));
            grpprl.Add((byte)(operand & 0xFF));
            grpprl.Add((byte)(operand >> 8));
        }

        private static void AddLineSpacingSprm(List<byte> grpprl, int lineSpacingTwips) {
            AddInt16Sprm(grpprl, SprmPDyaLine, lineSpacingTwips);
            grpprl.Add(0);
            grpprl.Add(0);
        }

        private static void AddParagraphShadingSprm(List<byte> grpprl, LegacyDocParagraphShading shading) {
            if (!LegacyDocColorPalette.TryGetIcoForHex(shading.FillColorHex, out byte backgroundIco) || backgroundIco == 0) {
                throw new NotSupportedException("Native DOC saving supports paragraph shading only for Word 97-2003 palette fill colors.");
            }

            ushort shd80 = (ushort)(backgroundIco << 5);
            grpprl.Add((byte)(SprmPShd80 & 0xFF));
            grpprl.Add((byte)(SprmPShd80 >> 8));
            grpprl.Add((byte)(shd80 & 0xFF));
            grpprl.Add((byte)(shd80 >> 8));
        }

        private static void AddParagraphBorderSprms(List<byte> grpprl, LegacyDocParagraphBorders borders) {
            AddParagraphBorderSprmIfPresent(grpprl, SprmPBrcTop80, borders.Top);
            AddParagraphBorderSprmIfPresent(grpprl, SprmPBrcLeft80, borders.Left);
            AddParagraphBorderSprmIfPresent(grpprl, SprmPBrcBottom80, borders.Bottom);
            AddParagraphBorderSprmIfPresent(grpprl, SprmPBrcRight80, borders.Right);
            AddParagraphBorderSprmIfPresent(grpprl, SprmPBrcBetween80, borders.Between);
        }

        private static void AddParagraphBorderSprmIfPresent(List<byte> grpprl, ushort sprm, LegacyDocParagraphBorder border) {
            if (!border.HasAny) {
                return;
            }

            AddParagraphBorderSprm(grpprl, sprm, border);
        }

        private static void AddParagraphBorderSprm(List<byte> grpprl, ushort sprm, LegacyDocParagraphBorder border) {
            if (border.SizeEighthPoints <= 0 || border.SizeEighthPoints > byte.MaxValue || border.SpacePoints < 0 || border.SpacePoints > byte.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports paragraph border size and spacing only within Word 97-2003 BRC80 byte ranges.");
            }

            if (!TryMapParagraphBorderStyle(border.Style, out byte borderType)) {
                throw new NotSupportedException($"Native DOC saving does not support paragraph border style '{border.Style}'.");
            }

            if (!LegacyDocColorPalette.TryGetIcoForHex(border.ColorHex, out byte colorIndex)) {
                throw new NotSupportedException("Native DOC saving supports paragraph borders only with Word 97-2003 palette colors.");
            }

            grpprl.Add((byte)(sprm & 0xFF));
            grpprl.Add((byte)(sprm >> 8));
            grpprl.Add(checked((byte)border.SizeEighthPoints));
            grpprl.Add(borderType);
            grpprl.Add(colorIndex);
            grpprl.Add(checked((byte)border.SpacePoints));
        }

        private static bool TryMapParagraphBorderStyle(LegacyDocParagraphBorderStyle style, out byte borderType) {
            switch (style) {
                case LegacyDocParagraphBorderStyle.Single:
                    borderType = 0x01;
                    return true;
                case LegacyDocParagraphBorderStyle.Double:
                    borderType = 0x03;
                    return true;
                case LegacyDocParagraphBorderStyle.Dotted:
                    borderType = 0x06;
                    return true;
                case LegacyDocParagraphBorderStyle.Dashed:
                    borderType = 0x07;
                    return true;
                default:
                    borderType = 0;
                    return false;
            }
        }

        private static void AddTableRowHeightSprm(List<byte> grpprl, int rowHeightTwips, bool isExact) {
            if (rowHeightTwips <= 0 || rowHeightTwips > short.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports table row heights only as positive twip values within the Word 97-2003 signed twip range.");
            }

            int operand = isExact ? -rowHeightTwips : rowHeightTwips;
            AddInt16Sprm(grpprl, SprmTDyaRowHeight, operand);
        }

        private static void AddTablePreferredWidthSprm(List<byte> grpprl, LegacyDocTablePreferredWidth preferredWidth) {
            byte ftsWidth;
            switch (preferredWidth.Unit) {
                case LegacyDocTablePreferredWidthUnit.Percent:
                    ftsWidth = FtsPercent;
                    break;
                case LegacyDocTablePreferredWidthUnit.Dxa:
                    ftsWidth = FtsDxa;
                    break;
                case LegacyDocTablePreferredWidthUnit.Auto:
                    return;
                default:
                    throw new NotSupportedException($"Native DOC saving does not support table preferred width unit '{preferredWidth.Unit}'.");
            }

            if (preferredWidth.Value <= 0 || preferredWidth.Value > short.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports table preferred width only as a positive Word 97-2003 signed width value.");
            }

            grpprl.Add((byte)(SprmTTableWidth & 0xFF));
            grpprl.Add((byte)(SprmTTableWidth >> 8));
            grpprl.Add(ftsWidth);
            AddInt16(grpprl, preferredWidth.Value, "table preferred width");
        }

        private static int MapTableAlignment(LegacyDocTableAlignment alignment) {
            switch (alignment) {
                case LegacyDocTableAlignment.Left:
                    return 0;
                case LegacyDocTableAlignment.Center:
                    return 1;
                case LegacyDocTableAlignment.Right:
                    return 2;
                default:
                    throw new NotSupportedException($"Native DOC saving does not support table alignment '{alignment}'.");
            }
        }

        private static void AddTabStopsSprm(List<byte> grpprl, IReadOnlyList<LegacyDocTabStop> tabStops) {
            var clearTabStops = tabStops
                .Where(tabStop => tabStop.Alignment == LegacyDocTabStopAlignment.Clear)
                .ToArray();
            var addedTabStops = tabStops
                .Where(tabStop => tabStop.Alignment != LegacyDocTabStopAlignment.Clear)
                .ToArray();

            if (clearTabStops.Length > byte.MaxValue || addedTabStops.Length > byte.MaxValue) {
                throw new NotSupportedException("Native DOC saving cannot write more than 255 tab stops in one paragraph.");
            }

            var operand = new List<byte>(2 + (clearTabStops.Length * 2) + (addedTabStops.Length * 3));
            operand.Add((byte)clearTabStops.Length);
            foreach (LegacyDocTabStop tabStop in clearTabStops) {
                AddInt16(operand, tabStop.PositionTwips, "tab stop clear position");
            }

            operand.Add((byte)addedTabStops.Length);
            foreach (LegacyDocTabStop tabStop in addedTabStops) {
                AddInt16(operand, tabStop.PositionTwips, "tab stop position");
            }

            foreach (LegacyDocTabStop tabStop in addedTabStops) {
                if (!TryMapTabAlignment(tabStop.Alignment, out byte alignment)
                    || !TryMapTabLeader(tabStop.Leader, out byte leader)) {
                    throw new NotSupportedException("Native DOC saving encountered an unsupported tab stop alignment or leader.");
                }

                operand.Add((byte)(alignment | (leader << 3)));
            }

            AddVariableSprm(grpprl, SprmPChgTabsPapx, operand);
        }

        private static bool TryMapTabAlignment(LegacyDocTabStopAlignment alignment, out byte value) {
            switch (alignment) {
                case LegacyDocTabStopAlignment.Left:
                    value = 0;
                    return true;
                case LegacyDocTabStopAlignment.Center:
                    value = 1;
                    return true;
                case LegacyDocTabStopAlignment.Right:
                    value = 2;
                    return true;
                case LegacyDocTabStopAlignment.Decimal:
                    value = 3;
                    return true;
                case LegacyDocTabStopAlignment.Bar:
                    value = 4;
                    return true;
                default:
                    value = 0;
                    return false;
            }
        }

        private static bool TryMapTabLeader(LegacyDocTabStopLeader leader, out byte value) {
            switch (leader) {
                case LegacyDocTabStopLeader.None:
                    value = 0;
                    return true;
                case LegacyDocTabStopLeader.Dot:
                    value = 1;
                    return true;
                case LegacyDocTabStopLeader.Hyphen:
                    value = 2;
                    return true;
                case LegacyDocTabStopLeader.Underscore:
                    value = 3;
                    return true;
                case LegacyDocTabStopLeader.Heavy:
                    value = 4;
                    return true;
                case LegacyDocTabStopLeader.MiddleDot:
                    value = 5;
                    return true;
                default:
                    value = 0;
                    return false;
            }
        }

        private static void AddVariableSprm(List<byte> grpprl, ushort sprm, List<byte> operand) {
            if (operand.Count > byte.MaxValue) {
                throw new NotSupportedException("Native DOC saving cannot write a variable-size paragraph formatting record because the DOC record is too large.");
            }

            grpprl.Add((byte)(sprm & 0xFF));
            grpprl.Add((byte)(sprm >> 8));
            grpprl.Add((byte)operand.Count);
            grpprl.AddRange(operand);
        }

        private static void AddDefaultTableCellPaddingSprms(List<byte> grpprl, LegacyDocTableCellMargins margins) {
            AddTableCellPaddingSprmIfPresent(grpprl, SprmTCellPaddingDefault, 0, FbrcTop, margins.TopTwips);
            AddTableCellPaddingSprmIfPresent(grpprl, SprmTCellPaddingDefault, 0, FbrcRight, margins.RightTwips);
            AddTableCellPaddingSprmIfPresent(grpprl, SprmTCellPaddingDefault, 0, FbrcBottom, margins.BottomTwips);
            AddTableCellPaddingSprmIfPresent(grpprl, SprmTCellPaddingDefault, 0, FbrcLeft, margins.LeftTwips);
        }

        private static void AddDefaultTableCellSpacingSprm(List<byte> grpprl, int widthTwips) {
            AddTableCellCssaSprm(grpprl, SprmTCellSpacingDefault, 0, FbrcTop | FbrcRight | FbrcBottom | FbrcLeft, widthTwips, "table cell spacing");
        }

        private static void AddTableCellPaddingSprms(List<byte> grpprl, IReadOnlyList<LegacyDocTableCellMargins> margins) {
            if (margins.Count > byte.MaxValue) {
                throw new NotSupportedException("Native DOC saving cannot write more than 255 table cells in one row.");
            }

            for (int index = 0; index < margins.Count; index++) {
                LegacyDocTableCellMargins cellMargins = margins[index];
                AddTableCellPaddingSprmIfPresent(grpprl, SprmTCellPadding, index, FbrcTop, cellMargins.TopTwips);
                AddTableCellPaddingSprmIfPresent(grpprl, SprmTCellPadding, index, FbrcRight, cellMargins.RightTwips);
                AddTableCellPaddingSprmIfPresent(grpprl, SprmTCellPadding, index, FbrcBottom, cellMargins.BottomTwips);
                AddTableCellPaddingSprmIfPresent(grpprl, SprmTCellPadding, index, FbrcLeft, cellMargins.LeftTwips);
            }
        }

        private static void AddTableCellPaddingSprmIfPresent(List<byte> grpprl, ushort sprm, int cellIndex, byte sideMask, int? widthTwips) {
            if (widthTwips == null) {
                return;
            }

            if (widthTwips.Value < 0 || widthTwips.Value > 31680) {
                throw new NotSupportedException("Native DOC saving supports table cell margins only as nonnegative DXA twip values within the Word 97-2003 limit.");
            }

            AddTableCellCssaSprm(grpprl, sprm, cellIndex, sideMask, widthTwips.Value, "table cell margins");
        }

        private static void AddTableCellCssaSprm(List<byte> grpprl, ushort sprm, int cellIndex, byte sideMask, int widthTwips, string propertyName) {
            if (widthTwips < 0 || widthTwips > 31680) {
                throw new NotSupportedException($"Native DOC saving supports {propertyName} only as nonnegative DXA twip values within the Word 97-2003 limit.");
            }

            var operand = new List<byte>(6) {
                checked((byte)cellIndex),
                checked((byte)(cellIndex + 1)),
                sideMask,
                FtsDxa,
                (byte)(widthTwips & 0xFF),
                (byte)(widthTwips >> 8)
            };
            AddVariableSprm(grpprl, sprm, operand);
        }

        private static void AddTableCellShadingSprm(List<byte> grpprl, IReadOnlyList<LegacyDocTableCellShading> shadings) {
            if (shadings.Count * 2 > byte.MaxValue) {
                throw new NotSupportedException("Native DOC saving cannot write table cell shading for more than 127 cells in one row.");
            }

            var operand = new List<byte>(shadings.Count * 2);
            foreach (LegacyDocTableCellShading shading in shadings) {
                if (!shading.HasAny) {
                    operand.Add((byte)(Shd80Nil & 0xFF));
                    operand.Add((byte)(Shd80Nil >> 8));
                    continue;
                }

                if (!LegacyDocColorPalette.TryGetIcoForHex(shading.FillColorHex, out byte backgroundIco) || backgroundIco == 0) {
                    throw new NotSupportedException("Native DOC saving supports table cell shading only for Word 97-2003 palette fill colors.");
                }

                ushort shd80 = (ushort)(backgroundIco << 5);
                operand.Add((byte)(shd80 & 0xFF));
                operand.Add((byte)(shd80 >> 8));
            }

            AddVariableSprm(grpprl, SprmTDefTableShd80, operand);
        }

        private static void AddTableDefinitionSprm(
            List<byte> grpprl,
            IReadOnlyList<int> cellWidthsTwips,
            int? tableLeftIndentTwips,
            IReadOnlyList<LegacyDocTableCellHorizontalMerge> cellHorizontalMerges,
            IReadOnlyList<LegacyDocTableCellVerticalMerge> cellVerticalMerges,
            IReadOnlyList<LegacyDocTableCellVerticalAlignment> cellVerticalAlignments,
            IReadOnlyList<LegacyDocTableCellTextDirection> cellTextDirections,
            IReadOnlyList<bool> cellFitTexts,
            IReadOnlyList<bool> cellNoWraps,
            IReadOnlyList<bool> cellHideMarks,
            IReadOnlyList<LegacyDocTableCellBorders> cellBorders) {
            if (cellWidthsTwips.Count > byte.MaxValue) {
                throw new NotSupportedException("Native DOC saving cannot write more than 255 table cells in one row.");
            }

            var remainder = new List<byte>(1 + ((cellWidthsTwips.Count + 1) * 2) + (cellWidthsTwips.Count * Tc80Length));
            remainder.Add(checked((byte)cellWidthsTwips.Count));
            int edge = tableLeftIndentTwips ?? 0;
            AddInt16(remainder, edge, "table left edge");
            foreach (int width in cellWidthsTwips) {
                if (width <= 0) {
                    throw new NotSupportedException("Native DOC saving supports table cell widths only as positive twip values.");
                }

                edge = checked(edge + width);
                AddInt16(remainder, edge, "table cell edge");
            }

            for (int index = 0; index < cellWidthsTwips.Count; index++) {
                ushort flags = GetTableCellFormattingFlags(cellHorizontalMerges, cellVerticalMerges, cellVerticalAlignments, cellTextDirections, cellFitTexts, cellNoWraps, cellHideMarks, index);
                remainder.Add((byte)(flags & 0xFF));
                remainder.Add((byte)(flags >> 8));
                AddInt16(remainder, 0, "table cell preferred width");
                AddTableCellBorder(remainder, GetTableCellBorders(cellBorders, index).Top);
                AddTableCellBorder(remainder, GetTableCellBorders(cellBorders, index).Left);
                AddTableCellBorder(remainder, GetTableCellBorders(cellBorders, index).Bottom);
                AddTableCellBorder(remainder, GetTableCellBorders(cellBorders, index).Right);
            }

            int cb = checked(remainder.Count + 1);
            if (cb > ushort.MaxValue) {
                throw new NotSupportedException("Native DOC saving cannot write table row definitions because the DOC table definition is too large.");
            }

            grpprl.Add((byte)(SprmTDefTable & 0xFF));
            grpprl.Add((byte)(SprmTDefTable >> 8));
            grpprl.Add((byte)(cb & 0xFF));
            grpprl.Add((byte)(cb >> 8));
            grpprl.AddRange(remainder);
        }

        private static LegacyDocTableCellBorders GetTableCellBorders(IReadOnlyList<LegacyDocTableCellBorders> cellBorders, int index) {
            return index < cellBorders.Count ? cellBorders[index] : default;
        }

        private static void AddTableCellBorder(List<byte> bytes, LegacyDocTableCellBorder border) {
            if (!border.HasAny) {
                bytes.Add(0);
                bytes.Add(0);
                bytes.Add(0);
                bytes.Add(0);
                return;
            }

            if (border.SizeEighthPoints <= 0 || border.SizeEighthPoints > byte.MaxValue || border.SpacePoints < 0 || border.SpacePoints > byte.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports table cell border size and spacing only within Word 97-2003 BRC80 byte ranges.");
            }

            if (!TryMapTableCellBorderStyle(border.Style, out byte borderType)) {
                throw new NotSupportedException($"Native DOC saving does not support table cell border style '{border.Style}'.");
            }

            if (!LegacyDocColorPalette.TryGetIcoForHex(border.ColorHex, out byte colorIndex)) {
                throw new NotSupportedException("Native DOC saving supports table cell borders only with Word 97-2003 palette colors.");
            }

            bytes.Add(checked((byte)border.SizeEighthPoints));
            bytes.Add(borderType);
            bytes.Add(colorIndex);
            bytes.Add(checked((byte)border.SpacePoints));
        }

        private static bool TryMapTableCellBorderStyle(LegacyDocTableCellBorderStyle style, out byte borderType) {
            switch (style) {
                case LegacyDocTableCellBorderStyle.Single:
                    borderType = 0x01;
                    return true;
                case LegacyDocTableCellBorderStyle.Double:
                    borderType = 0x03;
                    return true;
                case LegacyDocTableCellBorderStyle.Dotted:
                    borderType = 0x06;
                    return true;
                case LegacyDocTableCellBorderStyle.Dashed:
                    borderType = 0x07;
                    return true;
                default:
                    borderType = 0;
                    return false;
            }
        }

        private static ushort GetTableCellFormattingFlags(
            IReadOnlyList<LegacyDocTableCellHorizontalMerge> cellHorizontalMerges,
            IReadOnlyList<LegacyDocTableCellVerticalMerge> cellVerticalMerges,
            IReadOnlyList<LegacyDocTableCellVerticalAlignment> cellVerticalAlignments,
            IReadOnlyList<LegacyDocTableCellTextDirection> cellTextDirections,
            IReadOnlyList<bool> cellFitTexts,
            IReadOnlyList<bool> cellNoWraps,
            IReadOnlyList<bool> cellHideMarks,
            int index) {
            return (ushort)(GetTableCellHorizontalMergeFlags(cellHorizontalMerges, index)
                | GetTableCellVerticalMergeFlags(cellVerticalMerges, index)
                | GetTableCellVerticalAlignmentFlags(cellVerticalAlignments, index)
                | GetTableCellTextDirectionFlags(cellTextDirections, index)
                | GetTableCellFitTextFlags(cellFitTexts, index)
                | GetTableCellNoWrapFlags(cellNoWraps, index)
                | GetTableCellHideMarkFlags(cellHideMarks, index));
        }

        private static ushort GetTableCellHorizontalMergeFlags(IReadOnlyList<LegacyDocTableCellHorizontalMerge> cellHorizontalMerges, int index) {
            if (index >= cellHorizontalMerges.Count) {
                return 0;
            }

            switch (cellHorizontalMerges[index]) {
                case LegacyDocTableCellHorizontalMerge.Restart:
                    return 0x0001;
                case LegacyDocTableCellHorizontalMerge.Continue:
                    return 0x0002;
                default:
                    return 0;
            }
        }

        private static ushort GetTableCellVerticalMergeFlags(IReadOnlyList<LegacyDocTableCellVerticalMerge> cellVerticalMerges, int index) {
            if (index >= cellVerticalMerges.Count) {
                return 0;
            }

            switch (cellVerticalMerges[index]) {
                case LegacyDocTableCellVerticalMerge.Restart:
                    return 0x0020;
                case LegacyDocTableCellVerticalMerge.Continue:
                    return 0x0040;
                default:
                    return 0;
            }
        }

        private static ushort GetTableCellVerticalAlignmentFlags(IReadOnlyList<LegacyDocTableCellVerticalAlignment> cellVerticalAlignments, int index) {
            if (index >= cellVerticalAlignments.Count) {
                return 0;
            }

            switch (cellVerticalAlignments[index]) {
                case LegacyDocTableCellVerticalAlignment.Center:
                    return 0x0080;
                case LegacyDocTableCellVerticalAlignment.Bottom:
                    return 0x0100;
                default:
                    return 0;
            }
        }

        private static ushort GetTableCellTextDirectionFlags(IReadOnlyList<LegacyDocTableCellTextDirection> cellTextDirections, int index) {
            if (index >= cellTextDirections.Count) {
                return 0;
            }

            switch (cellTextDirections[index]) {
                case LegacyDocTableCellTextDirection.TopToBottomRightToLeft:
                    return 0x0004;
                case LegacyDocTableCellTextDirection.BottomToTopLeftToRight:
                    return 0x000C;
                case LegacyDocTableCellTextDirection.LeftToRightTopToBottomRotated:
                    return 0x0010;
                case LegacyDocTableCellTextDirection.TopToBottomRightToLeftRotated:
                    return 0x0014;
                default:
                    return 0;
            }
        }

        private static ushort GetTableCellFitTextFlags(IReadOnlyList<bool> cellFitTexts, int index) {
            return index < cellFitTexts.Count && cellFitTexts[index] ? (ushort)0x1000 : (ushort)0;
        }

        private static ushort GetTableCellNoWrapFlags(IReadOnlyList<bool> cellNoWraps, int index) {
            return index < cellNoWraps.Count && cellNoWraps[index] ? (ushort)0x2000 : (ushort)0;
        }

        private static ushort GetTableCellHideMarkFlags(IReadOnlyList<bool> cellHideMarks, int index) {
            return index < cellHideMarks.Count && cellHideMarks[index] ? (ushort)0x4000 : (ushort)0;
        }

        private static void AddInt16(List<byte> bytes, int operand, string propertyName) {
            if (operand < short.MinValue || operand > short.MaxValue) {
                throw new NotSupportedException($"Native DOC saving supports {propertyName} only within the Word 97-2003 signed twip range.");
            }

            short value = checked((short)operand);
            bytes.Add((byte)(value & 0xFF));
            bytes.Add((byte)(value >> 8));
        }

        private static void WriteInt32(byte[] bytes, int offset, int value) {
            bytes[offset] = (byte)value;
            bytes[offset + 1] = (byte)(value >> 8);
            bytes[offset + 2] = (byte)(value >> 16);
            bytes[offset + 3] = (byte)(value >> 24);
        }
    }

    internal readonly struct LegacyDocWritableParagraphFormatting : IEquatable<LegacyDocWritableParagraphFormatting> {
        internal static readonly LegacyDocWritableParagraphFormatting Plain = new LegacyDocWritableParagraphFormatting(null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, false, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null);

        internal LegacyDocWritableParagraphFormatting(
            byte? alignment,
            ushort? styleIndex,
            int? spacingBeforeTwips,
            int? spacingAfterTwips,
            int? lineSpacingTwips,
            int? leftIndentTwips,
            int? rightIndentTwips,
            int? firstLineIndentTwips,
            bool? keepLinesTogether,
            bool? keepWithNext,
            bool? pageBreakBefore,
            bool? avoidWidowAndOrphan,
            ushort? numberingListIndex,
            byte? numberingLevel,
            bool? isInTable,
            bool? isTableTerminatingParagraph,
            IReadOnlyList<LegacyDocTabStop>? tabStops,
            IReadOnlyList<int>? tableCellWidthsTwips,
            int? tableLeftIndentTwips,
            int? tableRowHeightTwips,
            bool tableRowHeightIsExact,
            bool? tableRowCantSplit,
            bool? tableRowIsHeader,
            LegacyDocTableAlignment? tableAlignment,
            LegacyDocTablePreferredWidth? tablePreferredWidth,
            bool? tableAutofit,
            IReadOnlyList<LegacyDocTableCellHorizontalMerge>? tableCellHorizontalMerges,
            IReadOnlyList<LegacyDocTableCellVerticalMerge>? tableCellVerticalMerges,
            IReadOnlyList<LegacyDocTableCellVerticalAlignment>? tableCellVerticalAlignments,
            IReadOnlyList<LegacyDocTableCellTextDirection>? tableCellTextDirections,
            IReadOnlyList<bool>? tableCellFitTexts,
            IReadOnlyList<bool>? tableCellNoWraps,
            IReadOnlyList<bool>? tableCellHideMarks,
            IReadOnlyList<LegacyDocTableCellMargins>? tableCellMargins,
            IReadOnlyList<LegacyDocTableCellShading>? tableCellShadings,
            IReadOnlyList<LegacyDocTableCellBorders>? tableCellBorders,
            LegacyDocParagraphShading? paragraphShading = null,
            LegacyDocParagraphBorders? paragraphBorders = null,
            LegacyDocTableCellMargins? defaultTableCellMargins = null,
            int? defaultTableCellSpacingTwips = null) {
            Alignment = alignment;
            StyleIndex = styleIndex;
            SpacingBeforeTwips = spacingBeforeTwips;
            SpacingAfterTwips = spacingAfterTwips;
            LineSpacingTwips = lineSpacingTwips;
            LeftIndentTwips = leftIndentTwips;
            RightIndentTwips = rightIndentTwips;
            FirstLineIndentTwips = firstLineIndentTwips;
            KeepLinesTogether = keepLinesTogether;
            KeepWithNext = keepWithNext;
            PageBreakBefore = pageBreakBefore;
            AvoidWidowAndOrphan = avoidWidowAndOrphan;
            NumberingListIndex = numberingListIndex.HasValue && numberingListIndex.Value > 0
                ? numberingListIndex
                : null;
            NumberingLevel = numberingLevel.HasValue && numberingLevel.Value <= 8
                ? numberingLevel
                : null;
            IsInTable = isInTable;
            IsTableTerminatingParagraph = isTableTerminatingParagraph;
            TabStops = tabStops == null || tabStops.Count == 0
                ? Array.Empty<LegacyDocTabStop>()
                : tabStops.ToArray();
            TableCellWidthsTwips = tableCellWidthsTwips == null || tableCellWidthsTwips.Count == 0
                ? Array.Empty<int>()
                : tableCellWidthsTwips.ToArray();
            TableLeftIndentTwips = tableLeftIndentTwips.HasValue && tableLeftIndentTwips.Value > 0 && tableLeftIndentTwips.Value <= short.MaxValue
                ? tableLeftIndentTwips
                : null;
            TableCellHorizontalMerges = tableCellHorizontalMerges == null || tableCellHorizontalMerges.Count == 0
                ? Array.Empty<LegacyDocTableCellHorizontalMerge>()
                : tableCellHorizontalMerges.ToArray();
            TableCellVerticalMerges = tableCellVerticalMerges == null || tableCellVerticalMerges.Count == 0
                ? Array.Empty<LegacyDocTableCellVerticalMerge>()
                : tableCellVerticalMerges.ToArray();
            TableCellVerticalAlignments = tableCellVerticalAlignments == null || tableCellVerticalAlignments.Count == 0
                ? Array.Empty<LegacyDocTableCellVerticalAlignment>()
                : tableCellVerticalAlignments.ToArray();
            TableCellTextDirections = tableCellTextDirections == null || tableCellTextDirections.Count == 0
                ? Array.Empty<LegacyDocTableCellTextDirection>()
                : tableCellTextDirections.ToArray();
            TableCellFitTexts = tableCellFitTexts == null || tableCellFitTexts.Count == 0
                ? Array.Empty<bool>()
                : tableCellFitTexts.ToArray();
            TableCellNoWraps = tableCellNoWraps == null || tableCellNoWraps.Count == 0
                ? Array.Empty<bool>()
                : tableCellNoWraps.ToArray();
            TableCellHideMarks = tableCellHideMarks == null || tableCellHideMarks.Count == 0
                ? Array.Empty<bool>()
                : tableCellHideMarks.ToArray();
            TableCellMargins = tableCellMargins == null || tableCellMargins.Count == 0
                ? Array.Empty<LegacyDocTableCellMargins>()
                : tableCellMargins.ToArray();
            TableCellShadings = tableCellShadings == null || tableCellShadings.Count == 0
                ? Array.Empty<LegacyDocTableCellShading>()
                : tableCellShadings.ToArray();
            TableCellBorders = tableCellBorders == null || tableCellBorders.Count == 0
                ? Array.Empty<LegacyDocTableCellBorders>()
                : tableCellBorders.ToArray();
            DefaultTableCellMargins = defaultTableCellMargins.HasValue && defaultTableCellMargins.Value.HasAny
                ? defaultTableCellMargins
                : null;
            DefaultTableCellSpacingTwips = defaultTableCellSpacingTwips.HasValue && defaultTableCellSpacingTwips.Value >= 0 && defaultTableCellSpacingTwips.Value <= 31680
                ? defaultTableCellSpacingTwips
                : null;
            TableRowHeightTwips = tableRowHeightTwips;
            TableRowHeightIsExact = tableRowHeightIsExact;
            TableRowCantSplit = tableRowCantSplit;
            TableRowIsHeader = tableRowIsHeader;
            TableAlignment = tableAlignment;
            TablePreferredWidth = tablePreferredWidth;
            TableAutofit = tableAutofit;
            ParagraphShading = paragraphShading.HasValue && paragraphShading.Value.HasAny
                ? paragraphShading
                : null;
            ParagraphBorders = paragraphBorders.HasValue && paragraphBorders.Value.HasAny
                ? paragraphBorders
                : null;
        }

        internal byte? Alignment { get; }

        internal ushort? StyleIndex { get; }

        internal int? SpacingBeforeTwips { get; }

        internal int? SpacingAfterTwips { get; }

        internal int? LineSpacingTwips { get; }

        internal int? LeftIndentTwips { get; }

        internal int? RightIndentTwips { get; }

        internal int? FirstLineIndentTwips { get; }

        internal bool? KeepLinesTogether { get; }

        internal bool? KeepWithNext { get; }

        internal bool? PageBreakBefore { get; }

        internal bool? AvoidWidowAndOrphan { get; }

        internal ushort? NumberingListIndex { get; }

        internal byte? NumberingLevel { get; }

        internal bool? IsInTable { get; }

        internal bool? IsTableTerminatingParagraph { get; }

        internal IReadOnlyList<LegacyDocTabStop> TabStops { get; }

        internal IReadOnlyList<int> TableCellWidthsTwips { get; }

        internal int? TableLeftIndentTwips { get; }

        internal IReadOnlyList<LegacyDocTableCellHorizontalMerge> TableCellHorizontalMerges { get; }

        internal IReadOnlyList<LegacyDocTableCellVerticalMerge> TableCellVerticalMerges { get; }

        internal IReadOnlyList<LegacyDocTableCellVerticalAlignment> TableCellVerticalAlignments { get; }

        internal IReadOnlyList<LegacyDocTableCellTextDirection> TableCellTextDirections { get; }

        internal IReadOnlyList<bool> TableCellFitTexts { get; }

        internal IReadOnlyList<bool> TableCellNoWraps { get; }

        internal IReadOnlyList<bool> TableCellHideMarks { get; }

        internal IReadOnlyList<LegacyDocTableCellMargins> TableCellMargins { get; }

        internal LegacyDocTableCellMargins? DefaultTableCellMargins { get; }

        internal int? DefaultTableCellSpacingTwips { get; }

        internal IReadOnlyList<LegacyDocTableCellShading> TableCellShadings { get; }

        internal IReadOnlyList<LegacyDocTableCellBorders> TableCellBorders { get; }

        internal int? TableRowHeightTwips { get; }

        internal bool TableRowHeightIsExact { get; }

        internal bool? TableRowCantSplit { get; }

        internal bool? TableRowIsHeader { get; }

        internal LegacyDocTableAlignment? TableAlignment { get; }

        internal LegacyDocTablePreferredWidth? TablePreferredWidth { get; }

        internal bool? TableAutofit { get; }

        internal LegacyDocParagraphShading? ParagraphShading { get; }

        internal LegacyDocParagraphBorders? ParagraphBorders { get; }

        internal bool HasFormatting => Alignment != null
            || StyleIndex != null
            || SpacingBeforeTwips != null
            || SpacingAfterTwips != null
            || LineSpacingTwips != null
            || LeftIndentTwips != null
            || RightIndentTwips != null
            || FirstLineIndentTwips != null
            || KeepLinesTogether != null
            || KeepWithNext != null
            || PageBreakBefore != null
            || AvoidWidowAndOrphan != null
            || NumberingListIndex != null
            || NumberingLevel != null
            || IsInTable != null
            || IsTableTerminatingParagraph != null
            || TabStops.Count > 0
            || TableCellWidthsTwips.Count > 0
            || TableLeftIndentTwips != null
            || TableCellHorizontalMerges.Count > 0
            || TableCellVerticalMerges.Count > 0
            || TableCellVerticalAlignments.Count > 0
            || TableCellTextDirections.Count > 0
            || TableCellFitTexts.Count > 0
            || TableCellNoWraps.Count > 0
            || TableCellHideMarks.Count > 0
            || DefaultTableCellMargins != null
            || DefaultTableCellSpacingTwips != null
            || TableCellMargins.Count > 0
            || TableCellShadings.Count > 0
            || TableCellBorders.Count > 0
            || TableRowHeightTwips != null
            || TableRowCantSplit != null
            || TableRowIsHeader != null
            || TableAlignment != null
            || TablePreferredWidth != null
            || TableAutofit != null
            || ParagraphShading != null
            || ParagraphBorders != null;

        internal LegacyDocWritableParagraphFormatting WithStyleIndex(ushort styleIndex) {
            return new LegacyDocWritableParagraphFormatting(
                Alignment,
                styleIndex,
                SpacingBeforeTwips,
                SpacingAfterTwips,
                LineSpacingTwips,
                LeftIndentTwips,
                RightIndentTwips,
                FirstLineIndentTwips,
                KeepLinesTogether,
                KeepWithNext,
                PageBreakBefore,
                AvoidWidowAndOrphan,
                NumberingListIndex,
                NumberingLevel,
                IsInTable,
                IsTableTerminatingParagraph,
                TabStops,
                TableCellWidthsTwips,
                TableLeftIndentTwips,
                TableRowHeightTwips,
                TableRowHeightIsExact,
                TableRowCantSplit,
                TableRowIsHeader,
                TableAlignment,
                TablePreferredWidth,
                TableAutofit,
                TableCellHorizontalMerges,
                TableCellVerticalMerges,
                TableCellVerticalAlignments,
                TableCellTextDirections,
                TableCellFitTexts,
                TableCellNoWraps,
                TableCellHideMarks,
                TableCellMargins,
                TableCellShadings,
                TableCellBorders,
                ParagraphShading,
                ParagraphBorders,
                DefaultTableCellMargins,
                DefaultTableCellSpacingTwips);
        }

        internal LegacyDocWritableParagraphFormatting WithInheritedParagraphFormatting(LegacyDocWritableParagraphFormatting inherited) {
            if (!inherited.HasFormatting) {
                return this;
            }

            return new LegacyDocWritableParagraphFormatting(
                Alignment ?? inherited.Alignment,
                StyleIndex ?? inherited.StyleIndex,
                SpacingBeforeTwips ?? inherited.SpacingBeforeTwips,
                SpacingAfterTwips ?? inherited.SpacingAfterTwips,
                LineSpacingTwips ?? inherited.LineSpacingTwips,
                LeftIndentTwips ?? inherited.LeftIndentTwips,
                RightIndentTwips ?? inherited.RightIndentTwips,
                FirstLineIndentTwips ?? inherited.FirstLineIndentTwips,
                KeepLinesTogether ?? inherited.KeepLinesTogether,
                KeepWithNext ?? inherited.KeepWithNext,
                PageBreakBefore ?? inherited.PageBreakBefore,
                AvoidWidowAndOrphan ?? inherited.AvoidWidowAndOrphan,
                NumberingListIndex ?? inherited.NumberingListIndex,
                NumberingLevel ?? inherited.NumberingLevel,
                IsInTable ?? inherited.IsInTable,
                IsTableTerminatingParagraph ?? inherited.IsTableTerminatingParagraph,
                TabStops.Count > 0 ? TabStops : inherited.TabStops,
                TableCellWidthsTwips.Count > 0 ? TableCellWidthsTwips : inherited.TableCellWidthsTwips,
                TableLeftIndentTwips ?? inherited.TableLeftIndentTwips,
                TableRowHeightTwips ?? inherited.TableRowHeightTwips,
                TableRowHeightTwips != null ? TableRowHeightIsExact : inherited.TableRowHeightIsExact,
                TableRowCantSplit ?? inherited.TableRowCantSplit,
                TableRowIsHeader ?? inherited.TableRowIsHeader,
                TableAlignment ?? inherited.TableAlignment,
                TablePreferredWidth ?? inherited.TablePreferredWidth,
                TableAutofit ?? inherited.TableAutofit,
                TableCellHorizontalMerges.Count > 0 ? TableCellHorizontalMerges : inherited.TableCellHorizontalMerges,
                TableCellVerticalMerges.Count > 0 ? TableCellVerticalMerges : inherited.TableCellVerticalMerges,
                TableCellVerticalAlignments.Count > 0 ? TableCellVerticalAlignments : inherited.TableCellVerticalAlignments,
                TableCellTextDirections.Count > 0 ? TableCellTextDirections : inherited.TableCellTextDirections,
                TableCellFitTexts.Count > 0 ? TableCellFitTexts : inherited.TableCellFitTexts,
                TableCellNoWraps.Count > 0 ? TableCellNoWraps : inherited.TableCellNoWraps,
                TableCellHideMarks.Count > 0 ? TableCellHideMarks : inherited.TableCellHideMarks,
                TableCellMargins.Count > 0 ? TableCellMargins : inherited.TableCellMargins,
                TableCellShadings.Count > 0 ? TableCellShadings : inherited.TableCellShadings,
                TableCellBorders.Count > 0 ? TableCellBorders : inherited.TableCellBorders,
                ParagraphShading ?? inherited.ParagraphShading,
                ParagraphBorders ?? inherited.ParagraphBorders,
                DefaultTableCellMargins ?? inherited.DefaultTableCellMargins,
                DefaultTableCellSpacingTwips ?? inherited.DefaultTableCellSpacingTwips);
        }

        internal LegacyDocWritableParagraphFormatting WithTableMarkers(
            bool isTableTerminatingParagraph,
            IReadOnlyList<int>? tableCellWidthsTwips = null,
            int? tableLeftIndentTwips = null,
            int? tableRowHeightTwips = null,
            bool tableRowHeightIsExact = false,
            bool? tableRowCantSplit = null,
            bool? tableRowIsHeader = null,
            LegacyDocTableAlignment? tableAlignment = null,
            LegacyDocTablePreferredWidth? tablePreferredWidth = null,
            bool? tableAutofit = null,
            IReadOnlyList<LegacyDocTableCellHorizontalMerge>? tableCellHorizontalMerges = null,
            IReadOnlyList<LegacyDocTableCellVerticalMerge>? tableCellVerticalMerges = null,
            IReadOnlyList<LegacyDocTableCellVerticalAlignment>? tableCellVerticalAlignments = null,
            IReadOnlyList<LegacyDocTableCellTextDirection>? tableCellTextDirections = null,
            IReadOnlyList<bool>? tableCellFitTexts = null,
            IReadOnlyList<bool>? tableCellNoWraps = null,
            IReadOnlyList<bool>? tableCellHideMarks = null,
            IReadOnlyList<LegacyDocTableCellMargins>? tableCellMargins = null,
            IReadOnlyList<LegacyDocTableCellShading>? tableCellShadings = null,
            IReadOnlyList<LegacyDocTableCellBorders>? tableCellBorders = null,
            LegacyDocTableCellMargins? defaultTableCellMargins = null,
            int? defaultTableCellSpacingTwips = null) {
            return new LegacyDocWritableParagraphFormatting(
                Alignment,
                StyleIndex,
                SpacingBeforeTwips,
                SpacingAfterTwips,
                LineSpacingTwips,
                LeftIndentTwips,
                RightIndentTwips,
                FirstLineIndentTwips,
                KeepLinesTogether,
                KeepWithNext,
                PageBreakBefore,
                AvoidWidowAndOrphan,
                NumberingListIndex,
                NumberingLevel,
                true,
                isTableTerminatingParagraph ? true : null,
                TabStops,
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
                ParagraphShading,
                ParagraphBorders,
                defaultTableCellMargins,
                defaultTableCellSpacingTwips);
        }

        public bool Equals(LegacyDocWritableParagraphFormatting other) {
            return Alignment == other.Alignment
                && StyleIndex == other.StyleIndex
                && SpacingBeforeTwips == other.SpacingBeforeTwips
                && SpacingAfterTwips == other.SpacingAfterTwips
                && LineSpacingTwips == other.LineSpacingTwips
                && LeftIndentTwips == other.LeftIndentTwips
                && RightIndentTwips == other.RightIndentTwips
                && FirstLineIndentTwips == other.FirstLineIndentTwips
                && KeepLinesTogether == other.KeepLinesTogether
                && KeepWithNext == other.KeepWithNext
                && PageBreakBefore == other.PageBreakBefore
                && AvoidWidowAndOrphan == other.AvoidWidowAndOrphan
                && NumberingListIndex == other.NumberingListIndex
                && NumberingLevel == other.NumberingLevel
                && IsInTable == other.IsInTable
                && IsTableTerminatingParagraph == other.IsTableTerminatingParagraph
                && TabStopsEqual(TabStops, other.TabStops)
                && TableCellWidthsEqual(TableCellWidthsTwips, other.TableCellWidthsTwips)
                && TableLeftIndentTwips == other.TableLeftIndentTwips
                && TableCellHorizontalMergesEqual(TableCellHorizontalMerges, other.TableCellHorizontalMerges)
                && TableCellVerticalMergesEqual(TableCellVerticalMerges, other.TableCellVerticalMerges)
                && TableCellVerticalAlignmentsEqual(TableCellVerticalAlignments, other.TableCellVerticalAlignments)
                && TableCellTextDirectionsEqual(TableCellTextDirections, other.TableCellTextDirections)
                && TableCellBooleansEqual(TableCellFitTexts, other.TableCellFitTexts)
                && TableCellBooleansEqual(TableCellNoWraps, other.TableCellNoWraps)
                && TableCellBooleansEqual(TableCellHideMarks, other.TableCellHideMarks)
                && DefaultTableCellMargins.Equals(other.DefaultTableCellMargins)
                && DefaultTableCellSpacingTwips == other.DefaultTableCellSpacingTwips
                && TableCellMarginsEqual(TableCellMargins, other.TableCellMargins)
                && TableCellShadingsEqual(TableCellShadings, other.TableCellShadings)
                && TableCellBordersEqual(TableCellBorders, other.TableCellBorders)
                && TableRowHeightTwips == other.TableRowHeightTwips
                && TableRowHeightIsExact == other.TableRowHeightIsExact
                && TableRowCantSplit == other.TableRowCantSplit
                && TableRowIsHeader == other.TableRowIsHeader
                && TableAlignment == other.TableAlignment
                && TablePreferredWidth.Equals(other.TablePreferredWidth)
                && TableAutofit == other.TableAutofit
                && ParagraphShading.Equals(other.ParagraphShading)
                && ParagraphBorders.Equals(other.ParagraphBorders);
        }

        public override bool Equals(object? obj) {
            return obj is LegacyDocWritableParagraphFormatting other && Equals(other);
        }

        public override int GetHashCode() {
            int hash = 17;
            hash = (hash * 31) + Alignment.GetHashCode();
            hash = (hash * 31) + StyleIndex.GetHashCode();
            hash = (hash * 31) + SpacingBeforeTwips.GetHashCode();
            hash = (hash * 31) + SpacingAfterTwips.GetHashCode();
            hash = (hash * 31) + LineSpacingTwips.GetHashCode();
            hash = (hash * 31) + LeftIndentTwips.GetHashCode();
            hash = (hash * 31) + RightIndentTwips.GetHashCode();
            hash = (hash * 31) + FirstLineIndentTwips.GetHashCode();
            hash = (hash * 31) + KeepLinesTogether.GetHashCode();
            hash = (hash * 31) + KeepWithNext.GetHashCode();
            hash = (hash * 31) + PageBreakBefore.GetHashCode();
            hash = (hash * 31) + AvoidWidowAndOrphan.GetHashCode();
            hash = (hash * 31) + NumberingListIndex.GetHashCode();
            hash = (hash * 31) + NumberingLevel.GetHashCode();
            hash = (hash * 31) + IsInTable.GetHashCode();
            hash = (hash * 31) + IsTableTerminatingParagraph.GetHashCode();
            hash = (hash * 31) + TableLeftIndentTwips.GetHashCode();
            hash = (hash * 31) + TableRowHeightTwips.GetHashCode();
            hash = (hash * 31) + TableRowHeightIsExact.GetHashCode();
            hash = (hash * 31) + TableRowCantSplit.GetHashCode();
            hash = (hash * 31) + TableRowIsHeader.GetHashCode();
            hash = (hash * 31) + TableAlignment.GetHashCode();
            hash = (hash * 31) + TablePreferredWidth.GetHashCode();
            hash = (hash * 31) + TableAutofit.GetHashCode();
            hash = (hash * 31) + ParagraphShading.GetHashCode();
            hash = (hash * 31) + ParagraphBorders.GetHashCode();
            hash = (hash * 31) + DefaultTableCellMargins.GetHashCode();
            hash = (hash * 31) + DefaultTableCellSpacingTwips.GetHashCode();
            foreach (LegacyDocTableCellHorizontalMerge merge in TableCellHorizontalMerges) {
                hash = (hash * 31) + merge.GetHashCode();
            }

            foreach (LegacyDocTableCellVerticalMerge merge in TableCellVerticalMerges) {
                hash = (hash * 31) + merge.GetHashCode();
            }

            foreach (LegacyDocTableCellVerticalAlignment alignment in TableCellVerticalAlignments) {
                hash = (hash * 31) + alignment.GetHashCode();
            }

            foreach (LegacyDocTableCellTextDirection textDirection in TableCellTextDirections) {
                hash = (hash * 31) + textDirection.GetHashCode();
            }

            foreach (bool fitText in TableCellFitTexts) {
                hash = (hash * 31) + fitText.GetHashCode();
            }

            foreach (bool noWrap in TableCellNoWraps) {
                hash = (hash * 31) + noWrap.GetHashCode();
            }

            foreach (bool hideMark in TableCellHideMarks) {
                hash = (hash * 31) + hideMark.GetHashCode();
            }

            foreach (LegacyDocTableCellMargins margins in TableCellMargins) {
                hash = (hash * 31) + margins.GetHashCode();
            }

            foreach (LegacyDocTableCellShading shading in TableCellShadings) {
                hash = (hash * 31) + shading.GetHashCode();
            }

            foreach (LegacyDocTableCellBorders borders in TableCellBorders) {
                hash = (hash * 31) + borders.GetHashCode();
            }

            foreach (LegacyDocTabStop tabStop in TabStops) {
                hash = (hash * 31) + tabStop.GetHashCode();
            }

            foreach (int width in TableCellWidthsTwips) {
                hash = (hash * 31) + width.GetHashCode();
            }

            return hash;
        }

        private static bool TableCellWidthsEqual(IReadOnlyList<int> first, IReadOnlyList<int> second) {
            if (first.Count != second.Count) {
                return false;
            }

            for (int index = 0; index < first.Count; index++) {
                if (first[index] != second[index]) {
                    return false;
                }
            }

            return true;
        }

        private static bool TableCellHorizontalMergesEqual(IReadOnlyList<LegacyDocTableCellHorizontalMerge> first, IReadOnlyList<LegacyDocTableCellHorizontalMerge> second) {
            if (first.Count != second.Count) {
                return false;
            }

            for (int index = 0; index < first.Count; index++) {
                if (first[index] != second[index]) {
                    return false;
                }
            }

            return true;
        }

        private static bool TableCellVerticalMergesEqual(IReadOnlyList<LegacyDocTableCellVerticalMerge> first, IReadOnlyList<LegacyDocTableCellVerticalMerge> second) {
            if (first.Count != second.Count) {
                return false;
            }

            for (int index = 0; index < first.Count; index++) {
                if (first[index] != second[index]) {
                    return false;
                }
            }

            return true;
        }

        private static bool TableCellVerticalAlignmentsEqual(IReadOnlyList<LegacyDocTableCellVerticalAlignment> first, IReadOnlyList<LegacyDocTableCellVerticalAlignment> second) {
            if (first.Count != second.Count) {
                return false;
            }

            for (int index = 0; index < first.Count; index++) {
                if (first[index] != second[index]) {
                    return false;
                }
            }

            return true;
        }

        private static bool TableCellTextDirectionsEqual(IReadOnlyList<LegacyDocTableCellTextDirection> first, IReadOnlyList<LegacyDocTableCellTextDirection> second) {
            if (first.Count != second.Count) {
                return false;
            }

            for (int index = 0; index < first.Count; index++) {
                if (first[index] != second[index]) {
                    return false;
                }
            }

            return true;
        }

        private static bool TableCellBooleansEqual(IReadOnlyList<bool> first, IReadOnlyList<bool> second) {
            if (first.Count != second.Count) {
                return false;
            }

            for (int index = 0; index < first.Count; index++) {
                if (first[index] != second[index]) {
                    return false;
                }
            }

            return true;
        }

        private static bool TableCellMarginsEqual(IReadOnlyList<LegacyDocTableCellMargins> first, IReadOnlyList<LegacyDocTableCellMargins> second) {
            if (first.Count != second.Count) {
                return false;
            }

            for (int index = 0; index < first.Count; index++) {
                if (!first[index].Equals(second[index])) {
                    return false;
                }
            }

            return true;
        }

        private static bool TableCellShadingsEqual(IReadOnlyList<LegacyDocTableCellShading> first, IReadOnlyList<LegacyDocTableCellShading> second) {
            if (first.Count != second.Count) {
                return false;
            }

            for (int index = 0; index < first.Count; index++) {
                if (!first[index].Equals(second[index])) {
                    return false;
                }
            }

            return true;
        }

        private static bool TableCellBordersEqual(IReadOnlyList<LegacyDocTableCellBorders> first, IReadOnlyList<LegacyDocTableCellBorders> second) {
            if (first.Count != second.Count) {
                return false;
            }

            for (int index = 0; index < first.Count; index++) {
                if (!first[index].Equals(second[index])) {
                    return false;
                }
            }

            return true;
        }

        private static bool TabStopsEqual(IReadOnlyList<LegacyDocTabStop> first, IReadOnlyList<LegacyDocTabStop> second) {
            if (first.Count != second.Count) {
                return false;
            }

            for (int index = 0; index < first.Count; index++) {
                if (!first[index].Equals(second[index])) {
                    return false;
                }
            }

            return true;
        }
    }

    internal readonly struct LegacyDocWritableParagraph {
        internal LegacyDocWritableParagraph(int startCharacter, int length, LegacyDocWritableParagraphFormatting formatting) {
            StartCharacter = startCharacter;
            Length = length;
            Formatting = formatting;
        }

        internal int StartCharacter { get; }

        internal int Length { get; }

        internal int EndCharacter => StartCharacter + Length;

        internal LegacyDocWritableParagraphFormatting Formatting { get; }
    }

    internal readonly struct LegacyDocWritableParagraphSegment {
        internal LegacyDocWritableParagraphSegment(int startCharacter, int length, LegacyDocWritableParagraphFormatting formatting) {
            StartCharacter = startCharacter;
            Length = length;
            Formatting = formatting;
            PapxOverride = null;
        }

        internal LegacyDocWritableParagraphSegment(int startCharacter, int length, byte[] papxOverride) {
            StartCharacter = startCharacter;
            Length = length;
            Formatting = LegacyDocWritableParagraphFormatting.Plain;
            PapxOverride = papxOverride;
        }

        internal int StartCharacter { get; }

        internal int Length { get; }

        internal int EndCharacter => StartCharacter + Length;

        internal LegacyDocWritableParagraphFormatting Formatting { get; }

        internal byte[]? PapxOverride { get; }

        internal LegacyDocWritableParagraphSegment Extend(int additionalLength) {
            return PapxOverride == null
                ? new LegacyDocWritableParagraphSegment(StartCharacter, Length + additionalLength, Formatting)
                : new LegacyDocWritableParagraphSegment(StartCharacter, Length + additionalLength, PapxOverride);
        }

        internal bool CanMergeWith(LegacyDocWritableParagraphFormatting formatting) {
            return PapxOverride == null && Formatting.Equals(formatting);
        }
    }
}
