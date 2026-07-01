using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static class LegacyXlsScenarioWriter {
        private const int MaximumScenarioItems = 32;

        internal static bool SupportsWorksheetScenarios(ExcelSheet sheet, out string? reason) {
            reason = null;
            if (!TryCreateWorksheetScenarios(sheet, out _, out reason)) {
                return false;
            }

            return true;
        }

        internal static bool TryCreateScenarioManagerPayload(ExcelSheet sheet, out byte[]? payload) {
            payload = null;
            if (!TryCreateWorksheetScenarios(sheet, out WorksheetScenarios? scenarios, out _) || scenarios!.Scenarios.Count == 0) {
                return false;
            }

            payload = BuildScenarioManagerPayload(scenarios);
            return true;
        }

        internal static IReadOnlyList<byte[]> CreateScenarioPayloads(ExcelSheet sheet) {
            if (!TryCreateWorksheetScenarios(sheet, out WorksheetScenarios? scenarios, out _) || scenarios!.Scenarios.Count == 0) {
                return Array.Empty<byte[]>();
            }

            var payloads = new List<byte[]>(scenarios.Scenarios.Count);
            foreach (ScenarioData scenario in scenarios.Scenarios) {
                payloads.Add(BuildScenarioPayload(scenario));
            }

            return payloads;
        }

        private static bool TryCreateWorksheetScenarios(ExcelSheet sheet, out WorksheetScenarios? scenarios, out string? reason) {
            scenarios = null;
            reason = null;
            Scenarios? openXmlScenarios = sheet.WorksheetPart.Worksheet?.Elements<Scenarios>().FirstOrDefault();
            if (openXmlScenarios == null) {
                scenarios = new WorksheetScenarios(Array.Empty<ScenarioData>(), null, null, Array.Empty<CellRange>());
                return true;
            }

            if (openXmlScenarios.HasChildren && openXmlScenarios.ChildElements.Any(child => child is not Scenario)) {
                reason = "worksheet scenarios with unsupported metadata";
                return false;
            }

            if (!SupportsKnownScenariosAttributes(openXmlScenarios)) {
                reason = "worksheet scenarios with unsupported metadata";
                return false;
            }

            var parsedScenarios = new List<ScenarioData>();
            foreach (Scenario scenario in openXmlScenarios.Elements<Scenario>()) {
                if (!TryCreateScenario(scenario, out ScenarioData? parsedScenario, out reason)) {
                    return false;
                }

                parsedScenarios.Add(parsedScenario!);
            }

            if (parsedScenarios.Count == 0) {
                scenarios = new WorksheetScenarios(parsedScenarios, null, null, Array.Empty<CellRange>());
                return true;
            }

            if (parsedScenarios.Count > ushort.MaxValue) {
                reason = "worksheet scenario counts outside BIFF8 limits";
                return false;
            }

            ushort? current = null;
            if (openXmlScenarios.Current?.Value is uint currentValue) {
                if (currentValue >= parsedScenarios.Count || currentValue > short.MaxValue) {
                    reason = "current worksheet scenario indexes outside BIFF8 limits";
                    return false;
                }

                current = checked((ushort)currentValue);
            }

            ushort? shown = null;
            if (openXmlScenarios.Show?.Value is uint shownValue) {
                if (shownValue >= parsedScenarios.Count || shownValue > short.MaxValue) {
                    reason = "shown worksheet scenario indexes outside BIFF8 limits";
                    return false;
                }

                shown = checked((ushort)shownValue);
            }

            if (!TryParseResultRanges(openXmlScenarios.SequenceOfReferences?.InnerText, out IReadOnlyList<CellRange> resultRanges, out reason)) {
                return false;
            }

            scenarios = new WorksheetScenarios(parsedScenarios, current, shown, resultRanges);
            return true;
        }

        private static bool TryCreateScenario(Scenario scenario, out ScenarioData? parsedScenario, out string? reason) {
            parsedScenario = null;
            reason = null;

            if (!SupportsKnownScenarioAttributes(scenario)) {
                reason = "worksheet scenarios with unsupported metadata";
                return false;
            }

            string name = scenario.Name?.Value ?? string.Empty;
            if (string.IsNullOrWhiteSpace(name) || name.Length > byte.MaxValue) {
                reason = "worksheet scenario names outside BIFF8 limits";
                return false;
            }

            string user = scenario.User?.Value ?? string.Empty;
            if (user.Length > byte.MaxValue) {
                reason = "worksheet scenario user names outside BIFF8 limits";
                return false;
            }

            string comment = scenario.Comment?.Value ?? string.Empty;
            if (comment.Length > byte.MaxValue) {
                reason = "worksheet scenario comments outside BIFF8 limits";
                return false;
            }

            if (scenario.ChildElements.Any(child => child is not InputCells)) {
                reason = "worksheet scenarios with unsupported metadata";
                return false;
            }

            var inputCells = new List<ScenarioInputCell>();
            foreach (InputCells inputCell in scenario.Elements<InputCells>()) {
                if (!TryCreateInputCell(inputCell, out ScenarioInputCell parsedInputCell, out reason)) {
                    return false;
                }

                inputCells.Add(parsedInputCell);
            }

            uint? declaredCount = scenario.Count?.Value;
            if (declaredCount.HasValue && declaredCount.Value != inputCells.Count) {
                reason = "worksheet scenario cell counts";
                return false;
            }

            if (inputCells.Count == 0 || inputCells.Count > MaximumScenarioItems) {
                reason = "worksheet scenario cell counts outside BIFF8 limits";
                return false;
            }

            if (!SupportsScenarioPayloadLength(name, user, comment, inputCells, out reason)) {
                return false;
            }

            parsedScenario = new ScenarioData(
                name,
                scenario.Locked?.Value == true,
                scenario.Hidden?.Value == true,
                user,
                comment,
                inputCells);
            return true;
        }

        private static bool TryCreateInputCell(InputCells inputCell, out ScenarioInputCell parsedInputCell, out string? reason) {
            parsedInputCell = default;
            reason = null;

            if (inputCell.HasChildren || !SupportsKnownInputCellAttributes(inputCell)) {
                reason = "worksheet scenario cells with unsupported metadata";
                return false;
            }

            if (inputCell.Undone?.Value == true || inputCell.NumberFormatId?.Value is uint) {
                reason = "worksheet scenario cells with unsupported metadata";
                return false;
            }

            string? reference = inputCell.CellReference?.Value;
            if (string.IsNullOrWhiteSpace(reference)) {
                reason = "worksheet scenario cells without references";
                return false;
            }

            string normalizedReference = reference!.Replace("$", string.Empty);
            if (normalizedReference.Contains("!", StringComparison.Ordinal)
                || normalizedReference.Contains(":", StringComparison.Ordinal)
                || !A1.TryParseCellReferenceFast(normalizedReference, out int row, out int column)) {
                reason = "worksheet scenario cell references";
                return false;
            }

            if (row < 1 || row > 65536 || column < 1 || column > 256) {
                reason = "worksheet scenario cell references outside BIFF8 worksheet limits";
                return false;
            }

            string value = inputCell.Val?.Value ?? string.Empty;
            if (value.Length > ushort.MaxValue) {
                reason = "worksheet scenario values outside BIFF8 limits";
                return false;
            }

            parsedInputCell = new ScenarioInputCell(
                checked((ushort)(row - 1)),
                checked((ushort)(column - 1)),
                inputCell.Deleted?.Value == true,
                value);
            return true;
        }

        private static bool SupportsScenarioPayloadLength(
            string name,
            string user,
            string comment,
            IReadOnlyList<ScenarioInputCell> inputCells,
            out string? reason) {
            reason = null;
            long payloadLength = 8L
                + (4L * inputCells.Count)
                + GetUnicodeStringNoCchPayloadLength(name)
                + (user.Length > 0 ? GetUnicodeStringNoCchPayloadLength(user) : 0L)
                + (comment.Length > 0 ? GetUnicodeStringNoCchPayloadLength(comment) : 0L);

            foreach (ScenarioInputCell inputCell in inputCells) {
                payloadLength += 2L + GetUnicodeStringNoCchPayloadLength(inputCell.Value);
            }

            if (payloadLength > ushort.MaxValue) {
                reason = "worksheet scenario payload lengths outside BIFF8 limits";
                return false;
            }

            return true;
        }

        private static long GetUnicodeStringNoCchPayloadLength(string text) {
            return 1L + (2L * text.Length);
        }

        private static bool TryParseResultRanges(string? sequenceOfReferences, out IReadOnlyList<CellRange> ranges, out string? reason) {
            ranges = Array.Empty<CellRange>();
            reason = null;
            if (string.IsNullOrWhiteSpace(sequenceOfReferences)) {
                return true;
            }

            string[] parts = sequenceOfReferences!.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length > MaximumScenarioItems) {
                reason = "worksheet scenario result-range counts outside BIFF8 limits";
                return false;
            }

            var parsed = new List<CellRange>(parts.Length);
            foreach (string part in parts) {
                string reference = part.Replace("$", string.Empty);
                if (!TryParseReference(reference, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                    reason = "worksheet scenario result ranges";
                    return false;
                }

                if (firstRow < 1 || firstColumn < 1 || lastRow > 65536 || lastColumn > 256) {
                    reason = "worksheet scenario result ranges outside BIFF8 worksheet limits";
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

        private static bool TryParseReference(string reference, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn) {
            if (A1.TryParseRange(reference, out firstRow, out firstColumn, out lastRow, out lastColumn)) {
                return true;
            }

            if (!A1.TryParseCellReferenceFast(reference, out firstRow, out firstColumn)) {
                lastRow = lastColumn = 0;
                return false;
            }

            lastRow = firstRow;
            lastColumn = firstColumn;
            return true;
        }

        private static byte[] BuildScenarioManagerPayload(WorksheetScenarios scenarios) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, scenarios.Scenarios.Count);
            WriteInt16(stream, scenarios.CurrentScenarioIndex.HasValue ? checked((short)scenarios.CurrentScenarioIndex.Value) : (short)-1);
            WriteInt16(stream, scenarios.ShownScenarioIndex.HasValue ? checked((short)scenarios.ShownScenarioIndex.Value) : (short)-1);
            WriteUInt16(stream, scenarios.ResultRanges.Count);
            foreach (CellRange range in scenarios.ResultRanges) {
                WriteCellRange(stream, range);
            }

            return stream.ToArray();
        }

        private static byte[] BuildScenarioPayload(ScenarioData scenario) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, scenario.InputCells.Count);
            WriteUInt16(stream, BuildScenarioFlags(scenario));
            stream.WriteByte(checked((byte)scenario.Name.Length));
            stream.WriteByte(checked((byte)scenario.User.Length));
            stream.WriteByte(checked((byte)scenario.Comment.Length));
            stream.WriteByte(0);
            foreach (ScenarioInputCell inputCell in scenario.InputCells) {
                WriteCellReference(stream, inputCell);
            }

            WriteUnicodeStringNoCch(stream, scenario.Name);
            if (scenario.User.Length > 0) {
                WriteUnicodeStringNoCch(stream, scenario.User);
            }

            if (scenario.Comment.Length > 0) {
                WriteUnicodeStringNoCch(stream, scenario.Comment);
            }

            foreach (ScenarioInputCell inputCell in scenario.InputCells) {
                WriteUnicodeString(stream, inputCell.Value);
            }

            return stream.ToArray();
        }

        private static ushort BuildScenarioFlags(ScenarioData scenario) {
            ushort flags = 0;
            if (scenario.Locked) flags |= 0x0001;
            if (scenario.Hidden) flags |= 0x0002;
            return flags;
        }

        private static void WriteCellReference(Stream stream, ScenarioInputCell inputCell) {
            WriteUInt16(stream, inputCell.Row);
            ushort columnBits = inputCell.Column;
            if (inputCell.Deleted) {
                columnBits |= 0x4000;
            }

            WriteUInt16(stream, columnBits);
        }

        private static void WriteCellRange(Stream stream, CellRange range) {
            WriteUInt16(stream, range.FirstRow);
            WriteUInt16(stream, range.LastRow);
            WriteUInt16(stream, range.FirstColumn);
            WriteUInt16(stream, range.LastColumn);
        }

        private static bool SupportsKnownScenariosAttributes(OpenXmlElement element) {
            foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                if (!string.IsNullOrEmpty(attribute.NamespaceUri)) {
                    return false;
                }

                switch (attribute.LocalName) {
                    case "current":
                    case "show":
                    case "sqref":
                        break;
                    default:
                        return false;
                }
            }

            return true;
        }

        private static bool SupportsKnownScenarioAttributes(OpenXmlElement element) {
            foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                if (!string.IsNullOrEmpty(attribute.NamespaceUri)) {
                    return false;
                }

                switch (attribute.LocalName) {
                    case "name":
                    case "locked":
                    case "hidden":
                    case "count":
                    case "user":
                    case "comment":
                        break;
                    default:
                        return false;
                }
            }

            return true;
        }

        private static bool SupportsKnownInputCellAttributes(OpenXmlElement element) {
            foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                if (!string.IsNullOrEmpty(attribute.NamespaceUri)) {
                    return false;
                }

                switch (attribute.LocalName) {
                    case "r":
                    case "deleted":
                    case "undone":
                    case "val":
                    case "numFmtId":
                        break;
                    default:
                        return false;
                }
            }

            return true;
        }

        private static void WriteUnicodeString(Stream stream, string text) {
            WriteUInt16(stream, checked((ushort)text.Length));
            WriteUnicodeStringNoCch(stream, text);
        }

        private static void WriteUnicodeStringNoCch(Stream stream, string text) {
            stream.WriteByte(0x01);
            foreach (char ch in text) {
                WriteUInt16(stream, ch);
            }
        }

        private static void WriteInt16(Stream stream, short value) {
            WriteUInt16(stream, unchecked((ushort)value));
        }

        private static void WriteUInt16(Stream stream, int value) {
            stream.WriteByte((byte)(value & 0xff));
            stream.WriteByte((byte)((value >> 8) & 0xff));
        }

        private sealed class WorksheetScenarios {
            internal WorksheetScenarios(IReadOnlyList<ScenarioData> scenarios, ushort? currentScenarioIndex, ushort? shownScenarioIndex, IReadOnlyList<CellRange> resultRanges) {
                Scenarios = scenarios;
                CurrentScenarioIndex = currentScenarioIndex;
                ShownScenarioIndex = shownScenarioIndex;
                ResultRanges = resultRanges;
            }

            internal IReadOnlyList<ScenarioData> Scenarios { get; }

            internal ushort? CurrentScenarioIndex { get; }

            internal ushort? ShownScenarioIndex { get; }

            internal IReadOnlyList<CellRange> ResultRanges { get; }
        }

        private sealed class ScenarioData {
            internal ScenarioData(string name, bool locked, bool hidden, string user, string comment, IReadOnlyList<ScenarioInputCell> inputCells) {
                Name = name;
                Locked = locked;
                Hidden = hidden;
                User = user;
                Comment = comment;
                InputCells = inputCells;
            }

            internal string Name { get; }

            internal bool Locked { get; }

            internal bool Hidden { get; }

            internal string User { get; }

            internal string Comment { get; }

            internal IReadOnlyList<ScenarioInputCell> InputCells { get; }
        }

        private readonly struct ScenarioInputCell {
            internal ScenarioInputCell(ushort row, ushort column, bool deleted, string value) {
                Row = row;
                Column = column;
                Deleted = deleted;
                Value = value;
            }

            internal ushort Row { get; }

            internal ushort Column { get; }

            internal bool Deleted { get; }

            internal string Value { get; }
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
