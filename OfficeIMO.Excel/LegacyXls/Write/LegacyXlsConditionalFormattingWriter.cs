using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static class LegacyXlsConditionalFormattingWriter {
        internal static bool SupportsWorksheetConditionalFormatting(
            ExcelSheet sheet,
            int sheetIndex,
            LegacyXlsFormulaNameIndex formulaNameIndex,
            out string? reason) {
            reason = null;
            foreach (ConditionalFormatting conditionalFormatting in GetWorksheetConditionalFormattings(sheet)) {
                if (!TryCreateBlock(conditionalFormatting, headerId: 0, sheetIndex, formulaNameIndex, out _, out _, out _, out reason)) {
                    return false;
                }
            }

            return true;
        }

        internal static IReadOnlyList<ConditionalFormattingBlock> CreateBlocks(
            ExcelSheet sheet,
            int sheetIndex,
            LegacyXlsFormulaNameIndex formulaNameIndex) {
            var blocks = new List<ConditionalFormattingBlock>();
            ushort headerId = 1;
            foreach (ConditionalFormatting conditionalFormatting in GetWorksheetConditionalFormattings(sheet)) {
                if (TryCreateBlock(conditionalFormatting, headerId, sheetIndex, formulaNameIndex, out byte[]? headerPayload, out IReadOnlyList<byte[]> rulePayloads, out IReadOnlyList<byte[]> extensionPayloads, out _)) {
                    blocks.Add(new ConditionalFormattingBlock(headerPayload!, rulePayloads, extensionPayloads));
                    headerId++;
                }
            }

            return blocks;
        }

        private static IReadOnlyList<ConditionalFormatting> GetWorksheetConditionalFormattings(ExcelSheet sheet) {
            Worksheet? worksheet = sheet.WorksheetPart.Worksheet;
            return worksheet == null
                ? Array.Empty<ConditionalFormatting>()
                : worksheet.Elements<ConditionalFormatting>().ToArray();
        }

        private static bool TryCreateBlock(
            ConditionalFormatting conditionalFormatting,
            ushort headerId,
            int sheetIndex,
            LegacyXlsFormulaNameIndex formulaNameIndex,
            out byte[]? headerPayload,
            out IReadOnlyList<byte[]> rulePayloads,
            out IReadOnlyList<byte[]> extensionPayloads,
            out string? reason) {
            headerPayload = null;
            rulePayloads = Array.Empty<byte[]>();
            extensionPayloads = Array.Empty<byte[]>();
            reason = null;

            if (conditionalFormatting.Pivot?.Value == true) {
                reason = "conditional formatting pivot metadata";
                return false;
            }

            if (HasExtensionMetadata(conditionalFormatting)) {
                reason = "conditional formatting extension metadata";
                return false;
            }

            if (!SupportsConditionalFormattingMetadata(conditionalFormatting)) {
                reason = "conditional formatting metadata";
                return false;
            }

            if (!TryParseRanges(conditionalFormatting.SequenceOfReferences?.InnerText, out IReadOnlyList<CellRange> ranges, out reason)) {
                return false;
            }

            var payloads = new List<byte[]>();
            var extensionPayloadList = new List<byte[]>();
            ushort ruleIndex = 0;
            foreach (ConditionalFormattingRule rule in conditionalFormatting.Elements<ConditionalFormattingRule>()) {
                if (!TryCreateRulePayload(rule, ranges, headerId, ruleIndex, sheetIndex, formulaNameIndex, out byte[]? rulePayload, out byte[]? extensionPayload, out reason)) {
                    return false;
                }

                payloads.Add(rulePayload!);
                if (extensionPayload != null) {
                    extensionPayloadList.Add(extensionPayload);
                }

                ruleIndex++;
            }

            if (payloads.Count == 0 || payloads.Count > ushort.MaxValue) {
                reason = "conditional formatting rule counts outside BIFF8 limits";
                return false;
            }

            CellRange enclosingRange = GetEnclosingRange(ranges);
            using var stream = new MemoryStream();
            WriteUInt16(stream, checked((ushort)payloads.Count));
            WriteUInt16(stream, checked((ushort)(headerId << 1)));
            WriteCellRange(stream, enclosingRange);
            WriteUInt16(stream, checked((ushort)ranges.Count));
            foreach (CellRange range in ranges) {
                WriteCellRange(stream, range);
            }

            headerPayload = stream.ToArray();
            rulePayloads = payloads;
            extensionPayloads = extensionPayloadList.Count == 0 ? Array.Empty<byte[]>() : extensionPayloadList;
            return true;
        }

        private static bool TryCreateRulePayload(
            ConditionalFormattingRule rule,
            IReadOnlyList<CellRange> ranges,
            ushort headerId,
            ushort ruleIndex,
            int sheetIndex,
            LegacyXlsFormulaNameIndex formulaNameIndex,
            out byte[]? payload,
            out byte[]? extensionPayload,
            out string? reason) {
            payload = null;
            extensionPayload = null;
            reason = null;

            if (rule.FormatId?.Value != null) {
                reason = "conditional formatting differential formats";
                return false;
            }

            if (!SupportsConditionalFormattingRuleMetadata(rule)) {
                reason = "conditional formatting rule metadata";
                return false;
            }

            if (rule.Elements<OpenXmlElement>().Any(element => element is not Formula)) {
                reason = "conditional formatting visual or extension payloads";
                return false;
            }

            ConditionalFormatValues type = rule.Type?.Value ?? ConditionalFormatValues.CellIs;
            foreach (Formula formula in rule.Elements<Formula>()) {
                if (!SupportsConditionalFormattingFormulaMetadata(formula)) {
                    reason = "conditional formatting formula metadata";
                    return false;
                }
            }

            List<string> formulas = rule.Elements<Formula>().Select(formula => formula.Text ?? string.Empty).ToList();
            byte conditionType;
            byte comparisonOperator;
            bool requiresSecondFormula = false;
            if (type == ConditionalFormatValues.CellIs) {
                conditionType = 0x01;
                if (!TryGetOperatorCode(rule.Operator?.Value, out comparisonOperator, out requiresSecondFormula)) {
                    reason = "conditional formatting operators outside the BIFF8 classic rule subset";
                    return false;
                }

                if (formulas.Count < 1 || formulas.Count > 2 || (requiresSecondFormula && formulas.Count != 2)) {
                    reason = "conditional formatting formula counts";
                    return false;
                }
            } else if (type == ConditionalFormatValues.Expression || IsFormulaBackedRuleType(type)) {
                conditionType = 0x02;
                comparisonOperator = 0;
                if (formulas.Count == 0 && !TryCreateFormulaBackedRuleFormula(type, rule, ranges, out formulas, out reason)) {
                    return false;
                }

                if (formulas.Count != 1) {
                    reason = "conditional formatting expression formula counts";
                    return false;
                }
            } else {
                reason = "conditional formatting rule types outside the BIFF8 classic rule subset";
                return false;
            }

            CellRange anchorRange = ranges[0];
            if (!TryEncodeFormula(formulas[0], sheetIndex, formulaNameIndex, anchorRange, out byte[] formula1Tokens, out reason)) {
                return false;
            }

            byte[] formula2Tokens = Array.Empty<byte>();
            if (formulas.Count > 1) {
                if (!TryEncodeFormula(formulas[1], sheetIndex, formulaNameIndex, anchorRange, out formula2Tokens, out reason)) {
                    return false;
                }
            }

            using var stream = new MemoryStream();
            stream.WriteByte(conditionType);
            stream.WriteByte(comparisonOperator);
            WriteUInt16(stream, checked((ushort)formula1Tokens.Length));
            WriteUInt16(stream, checked((ushort)formula2Tokens.Length));
            stream.Write(formula1Tokens, 0, formula1Tokens.Length);
            stream.Write(formula2Tokens, 0, formula2Tokens.Length);
            payload = stream.ToArray();

            if (rule.StopIfTrue?.Value == true || rule.Priority?.Value != null) {
                if (!TryCreateExtensionPayload(rule, headerId, ruleIndex, out extensionPayload, out reason)) {
                    return false;
                }
            }

            return true;
        }

        private static bool IsFormulaBackedRuleType(ConditionalFormatValues type) {
            return type == ConditionalFormatValues.ContainsText
                || type == ConditionalFormatValues.NotContainsText
                || type == ConditionalFormatValues.BeginsWith
                || type == ConditionalFormatValues.EndsWith
                || type == ConditionalFormatValues.ContainsBlanks
                || type == ConditionalFormatValues.NotContainsBlanks
                || type == ConditionalFormatValues.ContainsErrors
                || type == ConditionalFormatValues.NotContainsErrors
                || type == ConditionalFormatValues.TimePeriod
                || type == ConditionalFormatValues.DuplicateValues
                || type == ConditionalFormatValues.UniqueValues
                || type == ConditionalFormatValues.AboveAverage
                || type == ConditionalFormatValues.Top10;
        }

        private static bool TryCreateFormulaBackedRuleFormula(
            ConditionalFormatValues type,
            ConditionalFormattingRule rule,
            IReadOnlyList<CellRange> ranges,
            out List<string> formulas,
            out string? reason) {
            formulas = new List<string>();
            reason = null;
            if (ranges.Count != 1) {
                reason = "conditional formatting formula-backed rule ranges outside the native XLS subset";
                return false;
            }

            CellRange range = ranges[0];
            string rangeReference = FormatRange(range, absolute: true);
            string firstCell = A1.CellReference(range.FirstRow + 1, range.FirstColumn + 1);
            if (type == ConditionalFormatValues.DuplicateValues) {
                formulas.Add($"COUNTIF({rangeReference},{firstCell})>1");
                return true;
            }

            if (type == ConditionalFormatValues.UniqueValues) {
                formulas.Add($"COUNTIF({rangeReference},{firstCell})=1");
                return true;
            }

            if (type == ConditionalFormatValues.AboveAverage) {
                bool aboveAverage = rule.AboveAverage?.Value ?? true;
                bool equalAverage = rule.EqualAverage?.Value ?? false;
                int? standardDeviation = rule.StdDev?.Value;
                if (standardDeviation < 0) {
                    reason = "conditional formatting standard-deviation rules outside the native XLS subset";
                    return false;
                }

                string average = $"AVERAGE({rangeReference})";
                string threshold = standardDeviation.HasValue && standardDeviation.Value > 0
                    ? aboveAverage
                        ? $"{average}+{standardDeviation.Value}*STDEV({rangeReference})"
                        : $"{average}-{standardDeviation.Value}*STDEV({rangeReference})"
                    : average;
                string comparison = aboveAverage
                    ? equalAverage ? ">=" : ">"
                    : equalAverage ? "<=" : "<";
                formulas.Add($"{firstCell}{comparison}{threshold}");
                return true;
            }

            if (type == ConditionalFormatValues.Top10) {
                uint rank = rule.Rank?.Value ?? 10U;
                bool percent = rule.Percent?.Value == true;
                if (rank == 0 || (!percent && rank > ushort.MaxValue) || (percent && rank > 100U)) {
                    reason = "conditional formatting top/bottom ranks outside the native XLS subset";
                    return false;
                }

                bool bottom = rule.Bottom?.Value == true;
                string functionName = bottom ? "SMALL" : "LARGE";
                string comparison = bottom ? "<=" : ">=";
                string rankExpression = percent
                    ? $"ROUNDUP(COUNT({rangeReference})*{rank}/100,0)"
                    : rank.ToString(System.Globalization.CultureInfo.InvariantCulture);
                formulas.Add($"{firstCell}{comparison}{functionName}({rangeReference},{rankExpression})");
                return true;
            }

            reason = "conditional formatting rule types outside the BIFF8 classic rule subset";
            return false;
        }

        private static string FormatRange(CellRange range, bool absolute = false) {
            string start = FormatCell(range.FirstRow + 1, range.FirstColumn + 1, absolute);
            string end = FormatCell(range.LastRow + 1, range.LastColumn + 1, absolute);
            return start == end ? start : start + ":" + end;
        }

        private static string FormatCell(int row, int column, bool absolute) {
            return absolute
                ? A1.AbsoluteCellReference(row, column)
                : A1.CellReference(row, column);
        }

        private static bool TryCreateExtensionPayload(
            ConditionalFormattingRule rule,
            ushort headerId,
            ushort ruleIndex,
            out byte[]? payload,
            out string? reason) {
            payload = null;
            reason = null;

            int priorityValue = rule.Priority?.Value ?? 0;
            if (priorityValue < 0) {
                reason = "conditional formatting priorities outside BIFF8 limits";
                return false;
            }

            uint priority = (uint)priorityValue;
            if (priority > ushort.MaxValue) {
                reason = "conditional formatting priorities outside BIFF8 limits";
                return false;
            }

            using var stream = new MemoryStream();
            WriteUInt16(stream, 0x087b);
            WriteUInt16(stream, 0);
            WriteUInt16(stream, 0);
            WriteUInt16(stream, 0);
            WriteUInt16(stream, 0);
            WriteUInt16(stream, 0);
            WriteUInt32(stream, 0);
            WriteUInt16(stream, headerId);
            WriteUInt16(stream, ruleIndex);
            stream.WriteByte(0x05);
            stream.WriteByte(0);
            WriteUInt16(stream, checked((ushort)priority));
            stream.WriteByte((byte)(rule.StopIfTrue?.Value == true ? 0x03 : 0x01));
            stream.WriteByte(0);
            stream.WriteByte(16);
            for (int i = 0; i < 16; i++) {
                stream.WriteByte(0);
            }

            payload = stream.ToArray();
            return true;
        }

        private static bool TryEncodeFormula(
            string formula,
            int sheetIndex,
            LegacyXlsFormulaNameIndex formulaNameIndex,
            CellRange anchorRange,
            out byte[] tokens,
            out string? reason) {
            tokens = Array.Empty<byte>();
            reason = null;
            if (string.IsNullOrWhiteSpace(formula)) {
                reason = "conditional formatting formulas";
                return false;
            }

            if (!LegacyXlsFormulaEncoder.TryEncodeWithRelativeReferenceAnchor(
                formula,
                formulaNameIndex,
                sheetIndex,
                anchorRange.FirstRow,
                anchorRange.FirstColumn,
                out tokens,
                out string? formulaReason)) {
                reason = "conditional formatting formulas outside the native XLS formula subset: " + formulaReason;
                return false;
            }

            if (tokens.Length > ushort.MaxValue) {
                reason = "conditional formatting formula token payload lengths outside BIFF8 limits";
                return false;
            }

            return true;
        }

        private static bool TryGetOperatorCode(ConditionalFormattingOperatorValues? value, out byte code, out bool requiresSecondFormula) {
            requiresSecondFormula = false;
            if (!value.HasValue || value.Value == ConditionalFormattingOperatorValues.Between) {
                code = 0x01;
                requiresSecondFormula = true;
                return true;
            }

            if (value.Value == ConditionalFormattingOperatorValues.NotBetween) {
                code = 0x02;
                requiresSecondFormula = true;
                return true;
            }

            if (value.Value == ConditionalFormattingOperatorValues.Equal) code = 0x03;
            else if (value.Value == ConditionalFormattingOperatorValues.NotEqual) code = 0x04;
            else if (value.Value == ConditionalFormattingOperatorValues.GreaterThan) code = 0x05;
            else if (value.Value == ConditionalFormattingOperatorValues.LessThan) code = 0x06;
            else if (value.Value == ConditionalFormattingOperatorValues.GreaterThanOrEqual) code = 0x07;
            else if (value.Value == ConditionalFormattingOperatorValues.LessThanOrEqual) code = 0x08;
            else {
                code = 0;
                return false;
            }

            return true;
        }

        private static bool TryParseRanges(string? sequenceOfReferences, out IReadOnlyList<CellRange> ranges, out string? reason) {
            ranges = Array.Empty<CellRange>();
            reason = null;
            if (string.IsNullOrWhiteSpace(sequenceOfReferences)) {
                reason = "conditional formatting ranges";
                return false;
            }

            string[] parts = sequenceOfReferences!.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0 || parts.Length > 8192) {
                reason = "conditional formatting range counts outside BIFF8 limits";
                return false;
            }

            var parsed = new List<CellRange>(parts.Length);
            foreach (string part in parts) {
                string rangeText = part.Replace("$", string.Empty);
                if (!A1.TryParseRangeCoordinates(rangeText,
                        out int firstRow, out int firstColumn,
                        out int lastRow, out int lastColumn)) {
                    if (!A1.TryParseCellReferenceCoordinates(rangeText,
                            out firstRow, out firstColumn)) {
                        reason = "conditional formatting ranges";
                        return false;
                    }

                    lastRow = firstRow;
                    lastColumn = firstColumn;
                }

                if (firstRow < 1 || firstColumn < 1 || lastRow > 65536 || lastColumn > 256) {
                    reason = "conditional formatting ranges outside BIFF8 worksheet limits";
                    return false;
                }

                parsed.Add(new CellRange(
                    checked((ushort)(firstRow - 1)),
                    checked((ushort)(lastRow - 1)),
                    checked((ushort)(firstColumn - 1)),
                    checked((ushort)(lastColumn - 1))));
            }

            ranges = parsed;
            return true;
        }

        private static bool HasExtensionMetadata(OpenXmlElement element) {
            return element.Elements<ExtensionList>().Any(extensionList => extensionList.Elements<Extension>().Any());
        }

        private static bool SupportsConditionalFormattingMetadata(ConditionalFormatting conditionalFormatting) {
            foreach (OpenXmlAttribute attribute in conditionalFormatting.GetAttributes()) {
                if (!string.IsNullOrEmpty(attribute.NamespaceUri)) {
                    return false;
                }

                if (!string.Equals(attribute.LocalName, "sqref", StringComparison.Ordinal)
                    && !string.Equals(attribute.LocalName, "pivot", StringComparison.Ordinal)) {
                    return false;
                }
            }

            return true;
        }

        private static bool SupportsConditionalFormattingRuleMetadata(ConditionalFormattingRule rule) {
            foreach (OpenXmlAttribute attribute in rule.GetAttributes()) {
                if (!string.IsNullOrEmpty(attribute.NamespaceUri)) {
                    return false;
                }

                switch (attribute.LocalName) {
                    case "type":
                    case "dxfId":
                    case "priority":
                    case "stopIfTrue":
                    case "aboveAverage":
                    case "percent":
                    case "bottom":
                    case "operator":
                    case "text":
                    case "timePeriod":
                    case "rank":
                    case "stdDev":
                    case "equalAverage":
                        break;
                    default:
                        return false;
                }
            }

            return true;
        }

        private static bool SupportsConditionalFormattingFormulaMetadata(Formula formula) {
            return !formula.HasChildren && !formula.GetAttributes().Any();
        }

        private static CellRange GetEnclosingRange(IReadOnlyList<CellRange> ranges) {
            ushort firstRow = ushort.MaxValue;
            ushort lastRow = 0;
            ushort firstColumn = ushort.MaxValue;
            ushort lastColumn = 0;
            foreach (CellRange range in ranges) {
                if (range.FirstRow < firstRow) firstRow = range.FirstRow;
                if (range.LastRow > lastRow) lastRow = range.LastRow;
                if (range.FirstColumn < firstColumn) firstColumn = range.FirstColumn;
                if (range.LastColumn > lastColumn) lastColumn = range.LastColumn;
            }

            return new CellRange(firstRow, lastRow, firstColumn, lastColumn);
        }

        private static void WriteCellRange(Stream stream, CellRange range) {
            WriteUInt16(stream, range.FirstRow);
            WriteUInt16(stream, range.LastRow);
            WriteUInt16(stream, range.FirstColumn);
            WriteUInt16(stream, range.LastColumn);
        }

        private static void WriteUInt16(Stream stream, ushort value) {
            stream.WriteByte((byte)(value & 0xff));
            stream.WriteByte((byte)((value >> 8) & 0xff));
        }

        private static void WriteUInt32(Stream stream, uint value) {
            stream.WriteByte((byte)(value & 0xff));
            stream.WriteByte((byte)((value >> 8) & 0xff));
            stream.WriteByte((byte)((value >> 16) & 0xff));
            stream.WriteByte((byte)((value >> 24) & 0xff));
        }

        internal readonly struct ConditionalFormattingBlock {
            internal ConditionalFormattingBlock(byte[] headerPayload, IReadOnlyList<byte[]> rulePayloads, IReadOnlyList<byte[]> extensionPayloads) {
                HeaderPayload = headerPayload;
                RulePayloads = rulePayloads;
                ExtensionPayloads = extensionPayloads;
            }

            internal byte[] HeaderPayload { get; }
            internal IReadOnlyList<byte[]> RulePayloads { get; }
            internal IReadOnlyList<byte[]> ExtensionPayloads { get; }
        }

        private readonly struct CellRange {
            internal CellRange(ushort firstRow, ushort lastRow, ushort firstColumn, ushort lastColumn) {
                FirstRow = firstRow;
                LastRow = lastRow;
                FirstColumn = firstColumn;
                LastColumn = lastColumn;
            }

            internal ushort FirstRow { get; }
            internal ushort LastRow { get; }
            internal ushort FirstColumn { get; }
            internal ushort LastColumn { get; }
        }
    }
}
