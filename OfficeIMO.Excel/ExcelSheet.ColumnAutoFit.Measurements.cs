using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private AutoFitMeasurementPlan BuildAutoFitMeasurementPlanForAllColumns(CancellationToken ct) {
            var worksheet = WorksheetRoot;
            SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) {
                return new AutoFitMeasurementPlan(Array.Empty<int>(), new List<AutoFitMeasurement>());
            }

            var columnsList = new List<int>();
            var targetColumns = new Dictionary<int, int>();
            var measurements = new List<AutoFitMeasurement>();
            var uniqueMeasurements = new HashSet<AutoFitMeasurementKey>();
            var sharedStringMeasurements = new HashSet<(int TargetIndex, uint StyleIndex, int SharedStringId)>();
            var simpleTextMaxLengths = new Dictionary<(int TargetIndex, uint StyleIndex), int>();
            var textContext = CreateAutoFitTextContext();

            foreach (var row in sheetData.Elements<Row>()) {
                ct.ThrowIfCancellationRequested();

                foreach (var cell in row.Elements<Cell>()) {
                    string? reference = cell.CellReference?.Value;
                    if (string.IsNullOrEmpty(reference)) {
                        continue;
                    }

                    int columnIndex = GetColumnIndex(reference!);
                    if (!targetColumns.TryGetValue(columnIndex, out int targetIndex)) {
                        targetIndex = columnsList.Count;
                        targetColumns[columnIndex] = targetIndex;
                        columnsList.Add(columnIndex);
                    }

                    AddAutoFitMeasurement(cell, targetIndex, textContext, uniqueMeasurements, sharedStringMeasurements, simpleTextMaxLengths, measurements);
                }
            }

            var columns = worksheet.GetFirstChild<Columns>();
            if (columns != null) {
                foreach (var column in columns.Elements<Column>()) {
                    uint min = column.Min?.Value ?? 0;
                    uint max = column.Max?.Value ?? 0;
                    for (uint i = min; i <= max; i++) {
                        int columnIndex = (int)i;
                        if (!targetColumns.ContainsKey(columnIndex)) {
                            targetColumns[columnIndex] = columnsList.Count;
                            columnsList.Add(columnIndex);
                        }
                    }
                }
            }

            return new AutoFitMeasurementPlan(columnsList, measurements);
        }

        private AutoFitMeasurementPlan BuildAutoFitMeasurementPlan(IReadOnlyList<int> columnsList, CancellationToken ct) {
            var worksheet = WorksheetRoot;
            SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null || columnsList.Count == 0) {
                return new AutoFitMeasurementPlan(columnsList, new List<AutoFitMeasurement>());
            }

            var targetColumns = new Dictionary<int, int>(columnsList.Count);
            for (int i = 0; i < columnsList.Count; i++) {
                targetColumns[columnsList[i]] = i;
            }

            var measurements = new List<AutoFitMeasurement>();
            var uniqueMeasurements = new HashSet<AutoFitMeasurementKey>();
            var sharedStringMeasurements = new HashSet<(int TargetIndex, uint StyleIndex, int SharedStringId)>();
            var simpleTextMaxLengths = new Dictionary<(int TargetIndex, uint StyleIndex), int>();
            var textContext = CreateAutoFitTextContext();
            foreach (var row in sheetData.Elements<Row>()) {
                ct.ThrowIfCancellationRequested();

                foreach (var cell in row.Elements<Cell>()) {
                    string? reference = cell.CellReference?.Value;
                    if (string.IsNullOrEmpty(reference)) {
                        continue;
                    }

                    int columnIndex = GetColumnIndex(reference!);
                    if (!targetColumns.TryGetValue(columnIndex, out int targetIndex)) {
                        continue;
                    }

                    AddAutoFitMeasurement(cell, targetIndex, textContext, uniqueMeasurements, sharedStringMeasurements, simpleTextMaxLengths, measurements);
                }
            }

            return new AutoFitMeasurementPlan(columnsList, measurements);
        }

        private void AddAutoFitMeasurement(
            Cell cell,
            int targetIndex,
            AutoFitTextContext textContext,
            HashSet<AutoFitMeasurementKey> uniqueMeasurements,
            HashSet<(int TargetIndex, uint StyleIndex, int SharedStringId)> sharedStringMeasurements,
            Dictionary<(int TargetIndex, uint StyleIndex), int> simpleTextMaxLengths,
            List<AutoFitMeasurement> measurements) {
            uint styleIndex = cell.StyleIndex?.Value ?? 0U;
            if (cell.DataType?.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString
                && TryGetSharedStringIndex(cell, out int sharedStringId)
                && !sharedStringMeasurements.Add((targetIndex, styleIndex, sharedStringId))) {
                return;
            }

            if (TryAddDateAutoFitSampleMeasurement(cell, targetIndex, styleIndex, textContext, uniqueMeasurements, measurements)) {
                return;
            }

            if (CanSkipRawSimpleAutoFitMeasurement(cell, targetIndex, styleIndex, textContext, simpleTextMaxLengths)) {
                return;
            }

            string text = GetCellAutoFitText(cell, textContext);
            if (string.IsNullOrWhiteSpace(text)) {
                return;
            }

            if (CanUseSimpleAutoFitLengthShortcut(cell, text)) {
                var simpleKey = (targetIndex, styleIndex);
                if (simpleTextMaxLengths.TryGetValue(simpleKey, out int maxLength) && text.Length <= maxLength) {
                    return;
                }

                simpleTextMaxLengths[simpleKey] = text.Length;
            }

            var runs = GetCellAutoFitRichTextRuns(cell, textContext);
            if (uniqueMeasurements.Add(new AutoFitMeasurementKey(targetIndex, styleIndex, text, runs))) {
                measurements.Add(new AutoFitMeasurement(targetIndex, styleIndex, text, runs));
            }
        }

        private bool TryAddDateAutoFitSampleMeasurement(
            Cell cell,
            int targetIndex,
            uint styleIndex,
            AutoFitTextContext textContext,
            HashSet<AutoFitMeasurementKey> uniqueMeasurements,
            List<AutoFitMeasurement> measurements) {
            var dataType = cell.DataType?.Value;
            if (dataType != null && dataType != DocumentFormat.OpenXml.Spreadsheet.CellValues.Number) {
                return false;
            }

            uint numberFormatId = GetCellNumberFormatId(cell, textContext);
            string? formatCode = GetNumberFormatCode(numberFormatId, textContext);
            if (!IsDateNumberFormat(numberFormatId, formatCode)
                || !TryGetAutoFitDateSample(numberFormatId, formatCode, out string sample)) {
                return false;
            }

            string raw = cell.CellValue?.InnerText ?? string.Empty;
            if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out _)) {
                return false;
            }

            if (uniqueMeasurements.Add(new AutoFitMeasurementKey(targetIndex, styleIndex, sample, null))) {
                measurements.Add(new AutoFitMeasurement(targetIndex, styleIndex, sample, null));
            }

            return true;
        }

        private bool CanSkipRawSimpleAutoFitMeasurement(
            Cell cell,
            int targetIndex,
            uint styleIndex,
            AutoFitTextContext textContext,
            Dictionary<(int TargetIndex, uint StyleIndex), int> simpleTextMaxLengths) {
            var dataType = cell.DataType?.Value;
            if (dataType != null && dataType != DocumentFormat.OpenXml.Spreadsheet.CellValues.Number) {
                return false;
            }

            string raw = cell.CellValue?.InnerText ?? string.Empty;
            if (string.IsNullOrWhiteSpace(raw)) {
                return false;
            }

            uint numberFormatId = GetCellNumberFormatId(cell, textContext);
            if (numberFormatId != 0U) {
                string? formatCode = GetNumberFormatCode(numberFormatId, textContext);
                if (!string.IsNullOrWhiteSpace(formatCode)
                    && !string.Equals(formatCode, "General", StringComparison.OrdinalIgnoreCase)) {
                    return false;
                }
            }

            for (int i = 0; i < raw.Length; i++) {
                if (!IsSimpleAutoFitCharacter(raw[i])) {
                    return false;
                }
            }

            var simpleKey = (targetIndex, styleIndex);
            return simpleTextMaxLengths.TryGetValue(simpleKey, out int maxLength) && raw.Length <= maxLength;
        }

        private static bool CanUseSimpleAutoFitLengthShortcut(Cell cell, string text) {
            var dataType = cell.DataType?.Value;
            if (dataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString
                || dataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString) {
                return false;
            }

            for (int i = 0; i < text.Length; i++) {
                char current = text[i];
                if (current == '\n' || current == '\r' || !IsSimpleAutoFitCharacter(current)) {
                    return false;
                }
            }

            return true;
        }


        private static bool TryMeasureSimpleAutoFitTextWidth(
            string text,
            uint styleIndex,
            ExcelTextMeasurer.Style styleInfo,
            ExcelTextMeasurer textMeasurer,
            Dictionary<uint, Dictionary<char, float>> charWidthCache,
            out float measured) {
            measured = 0;
            if (string.IsNullOrEmpty(text)) {
                return false;
            }

            for (int i = 0; i < text.Length; i++) {
                if (!IsSimpleAutoFitCharacter(text[i])) {
                    return false;
                }
            }

            if (!charWidthCache.TryGetValue(styleIndex, out var perCharWidths)) {
                perCharWidths = new Dictionary<char, float>();
                charWidthCache[styleIndex] = perCharWidths;
            }

            float total = 0;
            for (int i = 0; i < text.Length; i++) {
                char current = text[i];
                if (!perCharWidths.TryGetValue(current, out float width)) {
                    width = textMeasurer.MeasureWidthOrDefault(current.ToString(), styleInfo, styleInfo.MaximumDigitWidth);
                    perCharWidths[current] = width;
                }

                total += width;
            }

            // Single-glyph summation can undercount string-level layout slightly on some fonts,
            // so bias upward by roughly one digit width to stay safely on the generous side.
            measured = total + styleInfo.MaximumDigitWidth;
            return true;
        }

        private readonly struct AutoFitMeasurement {
            internal AutoFitMeasurement(int targetIndex, uint styleIndex, string text, IReadOnlyList<AutoFitTextRun>? richTextRuns) {
                _targetIndex = targetIndex;
                _styleIndex = styleIndex;
                _text = text;
                _richTextRuns = richTextRuns;
            }

            private readonly int _targetIndex;
            private readonly uint _styleIndex;
            private readonly string _text;
            private readonly IReadOnlyList<AutoFitTextRun>? _richTextRuns;

            internal int TargetIndex => _targetIndex;
            internal uint StyleIndex => _styleIndex;
            internal string Text => _text;
            internal IReadOnlyList<AutoFitTextRun>? RichTextRuns => _richTextRuns;
        }

        private readonly struct AutoFitMeasurementKey : IEquatable<AutoFitMeasurementKey> {
            internal AutoFitMeasurementKey(int targetIndex, uint styleIndex, string text, IReadOnlyList<AutoFitTextRun>? richTextRuns) {
                _targetIndex = targetIndex;
                _styleIndex = styleIndex;
                _text = text;
                _richTextSignature = CreateRichTextSignature(richTextRuns);
            }

            private readonly int _targetIndex;
            private readonly uint _styleIndex;
            private readonly string _text;
            private readonly string? _richTextSignature;

            public bool Equals(AutoFitMeasurementKey other)
                => _targetIndex == other._targetIndex
                && _styleIndex == other._styleIndex
                && string.Equals(_text, other._text, StringComparison.Ordinal)
                && string.Equals(_richTextSignature, other._richTextSignature, StringComparison.Ordinal);

            public override bool Equals(object? obj)
                => obj is AutoFitMeasurementKey other && Equals(other);

            public override int GetHashCode() {
                unchecked {
                    int hash = _targetIndex;
                    hash = (hash * 397) ^ (int)_styleIndex;
                    hash = (hash * 397) ^ StringComparer.Ordinal.GetHashCode(_text);
                    hash = (hash * 397) ^ (_richTextSignature == null ? 0 : StringComparer.Ordinal.GetHashCode(_richTextSignature));
                    return hash;
                }
            }

            private static string? CreateRichTextSignature(IReadOnlyList<AutoFitTextRun>? runs) {
                if (runs == null || runs.Count == 0) {
                    return null;
                }

                var builder = new StringBuilder();
                for (int i = 0; i < runs.Count; i++) {
                    if (i > 0) {
                        builder.Append('|');
                    }

                    builder.Append(runs[i].Signature);
                }

                return builder.ToString();
            }
        }

        private sealed class AutoFitMeasurementPlan {
            internal AutoFitMeasurementPlan(IReadOnlyList<int> columns, List<AutoFitMeasurement> measurements) {
                Columns = columns;
                Measurements = measurements;
            }

            internal IReadOnlyList<int> Columns { get; }
            internal List<AutoFitMeasurement> Measurements { get; }
        }

        private sealed class AutoFitParallelState {
            internal AutoFitParallelState(int columnCount) {
                Widths = new double[columnCount];
                StyleCache = new Dictionary<uint, ExcelTextMeasurer.Style>();
                TextWidthCache = new Dictionary<(uint styleIndex, string text), float>();
                CharWidthCache = new Dictionary<uint, Dictionary<char, float>>();
            }

            internal double[] Widths { get; }
            internal Dictionary<uint, ExcelTextMeasurer.Style> StyleCache { get; }
            internal Dictionary<(uint styleIndex, string text), float> TextWidthCache { get; }
            internal Dictionary<uint, Dictionary<char, float>> CharWidthCache { get; }
        }
    }
}
