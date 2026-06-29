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
        private const ushort SprmPFWidowControl = 0x2431;
        private const ushort SprmPChgTabsPapx = 0xC60D;
        private const ushort SprmTDyaRowHeight = 0x9407;
        private const ushort SprmTDefTable = 0xD608;
        private const int Tc80Length = 20;

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

                byte[] papx = CreatePapx(segment.Formatting);
                papxOffset -= papx.Length;
                papxOffset = papxOffset % 2 == 0 ? papxOffset : papxOffset - 1;
                if (pageOffset + papxOffset <= (rgbxOffset + (segments.Count * PapxFkpBxLength)) || papxOffset / 2 > byte.MaxValue) {
                    throw new NotSupportedException("Native DOC saving currently supports paragraph formatting only when it fits in one paragraph-format page.");
                }

                Buffer.BlockCopy(papx, 0, stream, pageOffset + papxOffset, papx.Length);
                stream[rgbxOffset + (index * PapxFkpBxLength)] = (byte)(papxOffset / 2);
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

            if (formatting.IsInTable == true) {
                AddSingleByteSprm(grpprl, SprmPFInTable, 1);
            }

            if (formatting.IsTableTerminatingParagraph == true) {
                AddSingleByteSprm(grpprl, SprmPFTtp, 1);
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

            if (formatting.TableCellWidthsTwips.Count > 0) {
                AddTableDefinitionSprm(grpprl, formatting.TableCellWidthsTwips);
            }

            if (formatting.TableRowHeightTwips != null) {
                AddTableRowHeightSprm(grpprl, formatting.TableRowHeightTwips.Value, formatting.TableRowHeightIsExact);
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

        private static void AddLineSpacingSprm(List<byte> grpprl, int lineSpacingTwips) {
            AddInt16Sprm(grpprl, SprmPDyaLine, lineSpacingTwips);
            grpprl.Add(0);
            grpprl.Add(0);
        }

        private static void AddTableRowHeightSprm(List<byte> grpprl, int rowHeightTwips, bool isExact) {
            if (rowHeightTwips <= 0 || rowHeightTwips > short.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports table row heights only as positive twip values within the Word 97-2003 signed twip range.");
            }

            int operand = isExact ? -rowHeightTwips : rowHeightTwips;
            AddInt16Sprm(grpprl, SprmTDyaRowHeight, operand);
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
                throw new NotSupportedException("Native DOC saving cannot write paragraph tab stops because the DOC tab-stop record is too large.");
            }

            grpprl.Add((byte)(sprm & 0xFF));
            grpprl.Add((byte)(sprm >> 8));
            grpprl.Add((byte)operand.Count);
            grpprl.AddRange(operand);
        }

        private static void AddTableDefinitionSprm(List<byte> grpprl, IReadOnlyList<int> cellWidthsTwips) {
            if (cellWidthsTwips.Count > byte.MaxValue) {
                throw new NotSupportedException("Native DOC saving cannot write more than 255 table cells in one row.");
            }

            var remainder = new List<byte>(1 + ((cellWidthsTwips.Count + 1) * 2) + (cellWidthsTwips.Count * Tc80Length));
            remainder.Add(checked((byte)cellWidthsTwips.Count));
            AddInt16(remainder, 0, "table left edge");
            int edge = 0;
            foreach (int width in cellWidthsTwips) {
                if (width <= 0) {
                    throw new NotSupportedException("Native DOC saving supports table cell widths only as positive twip values.");
                }

                edge = checked(edge + width);
                AddInt16(remainder, edge, "table cell edge");
            }

            for (int index = 0; index < cellWidthsTwips.Count; index++) {
                for (int byteIndex = 0; byteIndex < Tc80Length; byteIndex++) {
                    remainder.Add(0);
                }
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
        internal static readonly LegacyDocWritableParagraphFormatting Plain = new LegacyDocWritableParagraphFormatting(null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, false);

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
            bool? isInTable,
            bool? isTableTerminatingParagraph,
            IReadOnlyList<LegacyDocTabStop>? tabStops,
            IReadOnlyList<int>? tableCellWidthsTwips,
            int? tableRowHeightTwips,
            bool tableRowHeightIsExact) {
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
            IsInTable = isInTable;
            IsTableTerminatingParagraph = isTableTerminatingParagraph;
            TabStops = tabStops == null || tabStops.Count == 0
                ? Array.Empty<LegacyDocTabStop>()
                : tabStops.ToArray();
            TableCellWidthsTwips = tableCellWidthsTwips == null || tableCellWidthsTwips.Count == 0
                ? Array.Empty<int>()
                : tableCellWidthsTwips.ToArray();
            TableRowHeightTwips = tableRowHeightTwips;
            TableRowHeightIsExact = tableRowHeightIsExact;
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

        internal bool? IsInTable { get; }

        internal bool? IsTableTerminatingParagraph { get; }

        internal IReadOnlyList<LegacyDocTabStop> TabStops { get; }

        internal IReadOnlyList<int> TableCellWidthsTwips { get; }

        internal int? TableRowHeightTwips { get; }

        internal bool TableRowHeightIsExact { get; }

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
            || IsInTable != null
            || IsTableTerminatingParagraph != null
            || TabStops.Count > 0
            || TableCellWidthsTwips.Count > 0
            || TableRowHeightTwips != null;

        internal LegacyDocWritableParagraphFormatting WithTableMarkers(bool isTableTerminatingParagraph, IReadOnlyList<int>? tableCellWidthsTwips = null, int? tableRowHeightTwips = null, bool tableRowHeightIsExact = false) {
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
                true,
                isTableTerminatingParagraph ? true : null,
                TabStops,
                tableCellWidthsTwips,
                tableRowHeightTwips,
                tableRowHeightIsExact);
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
                && IsInTable == other.IsInTable
                && IsTableTerminatingParagraph == other.IsTableTerminatingParagraph
                && TabStopsEqual(TabStops, other.TabStops)
                && TableCellWidthsEqual(TableCellWidthsTwips, other.TableCellWidthsTwips)
                && TableRowHeightTwips == other.TableRowHeightTwips
                && TableRowHeightIsExact == other.TableRowHeightIsExact;
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
            hash = (hash * 31) + IsInTable.GetHashCode();
            hash = (hash * 31) + IsTableTerminatingParagraph.GetHashCode();
            hash = (hash * 31) + TableRowHeightTwips.GetHashCode();
            hash = (hash * 31) + TableRowHeightIsExact.GetHashCode();
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
        }

        internal int StartCharacter { get; }

        internal int Length { get; }

        internal int EndCharacter => StartCharacter + Length;

        internal LegacyDocWritableParagraphFormatting Formatting { get; }

        internal LegacyDocWritableParagraphSegment Extend(int additionalLength) {
            return new LegacyDocWritableParagraphSegment(StartCharacter, Length + additionalLength, Formatting);
        }
    }
}
