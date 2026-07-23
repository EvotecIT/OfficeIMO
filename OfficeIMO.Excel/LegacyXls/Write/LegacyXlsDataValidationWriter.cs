using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static class LegacyXlsDataValidationWriter {
        private const uint AllowBlankFlag = 0x00000100;
        private const uint SuppressDropDownFlag = 0x00000200;
        private const uint ShowInputMessageFlag = 0x00040000;
        private const uint ShowErrorMessageFlag = 0x00080000;
        private const uint ExplicitListFlag = 0x00000080;

        internal static bool SupportsWorksheetDataValidations(
            ExcelSheet sheet,
            int sheetIndex,
            LegacyXlsFormulaNameIndex formulaNameIndex,
            out string? reason) {
            reason = null;
            DataValidations? validations = GetWorksheetDataValidationCollection(sheet);
            if (validations == null) {
                return true;
            }

            if (!SupportsDataValidationCollection(validations, out reason)) {
                return false;
            }

            foreach (DataValidation validation in validations.Elements<DataValidation>()) {
                if (!TryCreateRule(validation, sheetIndex, formulaNameIndex, out _, out reason)) {
                    return false;
                }
            }

            return true;
        }

        internal static bool TryCreateCollectionPayload(ExcelSheet sheet, out byte[]? payload) {
            payload = null;
            int count = GetWorksheetDataValidations(sheet).Count;
            if (count == 0) {
                return false;
            }

            using var stream = new MemoryStream();
            WriteUInt16(stream, 0);
            WriteUInt32(stream, 0);
            WriteUInt32(stream, 0);
            WriteUInt32(stream, 0xffffffff);
            WriteUInt32(stream, checked((uint)count));
            payload = stream.ToArray();
            return true;
        }

        internal static IReadOnlyList<byte[]> CreateValidationPayloads(
            ExcelSheet sheet,
            int sheetIndex,
            LegacyXlsFormulaNameIndex formulaNameIndex) {
            var payloads = new List<byte[]>();
            foreach (DataValidation validation in GetWorksheetDataValidations(sheet)) {
                if (TryCreateRule(validation, sheetIndex, formulaNameIndex, out DataValidationRule rule, out _)) {
                    payloads.Add(BuildRulePayload(rule));
                }
            }

            return payloads;
        }

        private static IReadOnlyList<DataValidation> GetWorksheetDataValidations(ExcelSheet sheet) {
            DataValidations? validations = GetWorksheetDataValidationCollection(sheet);
            return validations == null
                ? Array.Empty<DataValidation>()
                : validations.Elements<DataValidation>().ToArray();
        }

        private static DataValidations? GetWorksheetDataValidationCollection(ExcelSheet sheet) {
            return sheet.WorksheetPart.Worksheet?.GetFirstChild<DataValidations>();
        }

        private static bool SupportsDataValidationCollection(DataValidations validations, out string? reason) {
            reason = null;
            if (HasExtensionMetadata(validations)) {
                reason = "data validation extension metadata";
                return false;
            }

            foreach (OpenXmlAttribute attribute in validations.GetAttributes()) {
                if (!string.IsNullOrEmpty(attribute.NamespaceUri)
                    || !IsSupportedDataValidationCollectionAttribute(attribute.LocalName)) {
                    reason = "data validation collection metadata";
                    return false;
                }
            }

            return true;
        }

        private static bool IsSupportedDataValidationCollectionAttribute(string localName) {
            return string.Equals(localName, "count", StringComparison.Ordinal);
        }

        private static bool TryCreateRule(
            DataValidation validation,
            int sheetIndex,
            LegacyXlsFormulaNameIndex formulaNameIndex,
            out DataValidationRule rule,
            out string? reason) {
            rule = default;
            reason = null;
            if (HasExtensionMetadata(validation)) {
                reason = "data validation extension metadata";
                return false;
            }

            if (!SupportsDataValidationAttributes(validation, out reason)) {
                return false;
            }

            DataValidationValues type = validation.Type?.Value ?? DataValidationValues.None;
            if (!TryGetValidationTypeCode(type, out uint typeCode)) {
                reason = "data validation types outside the BIFF8 validation subset";
                return false;
            }

            if (!TryGetOperatorCode(validation.Operator?.Value, out uint operatorCode)) {
                reason = "data validation operators outside the BIFF8 validation subset";
                return false;
            }

            if (!TryGetErrorStyleCode(validation.ErrorStyle?.Value, out uint errorStyleCode)) {
                reason = "data validation error styles outside the BIFF8 validation subset";
                return false;
            }

            if (!SupportsSingleFormulaElement<Formula1>(validation, "Formula1", out reason)
                || !SupportsSingleFormulaElement<Formula2>(validation, "Formula2", out reason)) {
                return false;
            }

            string formula1 = validation.GetFirstChild<Formula1>()?.Text ?? string.Empty;
            string formula2 = validation.GetFirstChild<Formula2>()?.Text ?? string.Empty;

            if (!TryParseRanges(validation.SequenceOfReferences?.InnerText, out IReadOnlyList<CellRange> ranges, out reason)) {
                return false;
            }

            ushort anchorRow = ranges[0].FirstRow;
            ushort anchorColumn = ranges[0].FirstColumn;
            if (!TryEncodeFormula(type, formula1, required: type != DataValidationValues.None, sheetIndex, formulaNameIndex, anchorRow, anchorColumn, out byte[] formula1Tokens, out reason)) {
                return false;
            }

            bool requiresSecondFormula = type != DataValidationValues.None
                && type != DataValidationValues.List
                && type != DataValidationValues.Custom
                && (validation.Operator?.Value == DataValidationOperatorValues.Between
                    || validation.Operator?.Value == DataValidationOperatorValues.NotBetween);
            if (!TryEncodeFormula(type, formula2, required: requiresSecondFormula, sheetIndex, formulaNameIndex, anchorRow, anchorColumn, out byte[] formula2Tokens, out reason)) {
                return false;
            }

            string promptTitle = validation.PromptTitle?.Value ?? string.Empty;
            string errorTitle = validation.ErrorTitle?.Value ?? string.Empty;
            string prompt = validation.Prompt?.Value ?? string.Empty;
            string error = validation.Error?.Value ?? string.Empty;
            if (!SupportsTextPayloads(promptTitle, errorTitle, prompt, error, out reason)) {
                return false;
            }

            if (!SupportsRulePayloadLength(promptTitle, errorTitle, prompt, error, formula1Tokens, formula2Tokens, ranges.Count, out reason)) {
                return false;
            }

            rule = new DataValidationRule(
                typeCode,
                operatorCode,
                errorStyleCode,
                IsExplicitList(type, formula1),
                validation.AllowBlank?.Value == true,
                type == DataValidationValues.List && validation.ShowDropDown?.Value == true,
                validation.ShowInputMessage?.Value == true,
                validation.ShowErrorMessage?.Value == true,
                promptTitle,
                errorTitle,
                prompt,
                error,
                formula1Tokens,
                formula2Tokens,
                ranges);
            return true;
        }

        private static bool SupportsTextPayloads(string promptTitle, string errorTitle, string prompt, string error, out string? reason) {
            reason = null;
            long textPayloadLength = 0;
            foreach (string text in new[] { promptTitle, errorTitle, prompt, error }) {
                long unicodeByteCount = Encoding.Unicode.GetByteCount(text);
                if (text.Length > ushort.MaxValue || unicodeByteCount > ushort.MaxValue - 3L) {
                    reason = "data validation text payload lengths outside BIFF8 limits";
                    return false;
                }

                textPayloadLength += 3L + unicodeByteCount;
            }

            if (textPayloadLength > ushort.MaxValue - 14L) {
                reason = "data validation text payload lengths outside BIFF8 limits";
                return false;
            }

            return true;
        }

        private static bool SupportsRulePayloadLength(
            string promptTitle,
            string errorTitle,
            string prompt,
            string error,
            byte[] formula1Tokens,
            byte[] formula2Tokens,
            int rangeCount,
            out string? reason) {
            reason = null;
            long payloadLength = 4L
                + GetUnicodeStringPayloadLength(promptTitle)
                + GetUnicodeStringPayloadLength(errorTitle)
                + GetUnicodeStringPayloadLength(prompt)
                + GetUnicodeStringPayloadLength(error)
                + 4L + formula1Tokens.Length
                + 4L + formula2Tokens.Length
                + 2L + (8L * rangeCount);
            if (payloadLength > ushort.MaxValue) {
                reason = "data validation record payload lengths outside BIFF8 limits";
                return false;
            }

            return true;
        }

        private static long GetUnicodeStringPayloadLength(string text) {
            return 3L + Encoding.Unicode.GetByteCount(text);
        }

        private static bool SupportsDataValidationAttributes(DataValidation validation, out string? reason) {
            reason = null;
            foreach (OpenXmlAttribute attribute in validation.GetAttributes()) {
                if (!string.IsNullOrEmpty(attribute.NamespaceUri)) {
                    reason = "data validation extension metadata";
                    return false;
                }

                if (!IsSupportedDataValidationAttribute(attribute.LocalName)) {
                    reason = string.Equals(attribute.LocalName, "imeMode", StringComparison.Ordinal)
                        ? "data validation IME mode metadata"
                        : "data validation metadata";
                    return false;
                }
            }

            return true;
        }

        private static bool IsSupportedDataValidationAttribute(string localName) {
            return string.Equals(localName, "type", StringComparison.Ordinal)
                || string.Equals(localName, "errorStyle", StringComparison.Ordinal)
                || string.Equals(localName, "operator", StringComparison.Ordinal)
                || string.Equals(localName, "allowBlank", StringComparison.Ordinal)
                || string.Equals(localName, "showDropDown", StringComparison.Ordinal)
                || string.Equals(localName, "showInputMessage", StringComparison.Ordinal)
                || string.Equals(localName, "showErrorMessage", StringComparison.Ordinal)
                || string.Equals(localName, "errorTitle", StringComparison.Ordinal)
                || string.Equals(localName, "error", StringComparison.Ordinal)
                || string.Equals(localName, "promptTitle", StringComparison.Ordinal)
                || string.Equals(localName, "prompt", StringComparison.Ordinal)
                || string.Equals(localName, "sqref", StringComparison.Ordinal);
        }

        private static bool HasExtensionMetadata(OpenXmlElement element) {
            return element.Elements<ExtensionList>().Any(extensionList => extensionList.Elements<Extension>().Any());
        }

        private static bool SupportsSingleFormulaElement<TElement>(DataValidation validation, string elementName, out string? reason)
            where TElement : OpenXmlElement {
            reason = null;
            TElement[] formulas = validation.Elements<TElement>().Take(2).ToArray();
            if (formulas.Length > 1) {
                reason = $"data validation formulas with duplicate {elementName} elements";
                return false;
            }

            if (formulas.Length == 0) {
                return true;
            }

            if (formulas[0].HasChildren || formulas[0].GetAttributes().Any()) {
                reason = "data validation formula metadata";
                return false;
            }

            return true;
        }

        private static bool TryEncodeFormula(
            DataValidationValues type,
            string formula,
            bool required,
            int sheetIndex,
            LegacyXlsFormulaNameIndex formulaNameIndex,
            ushort anchorRow,
            ushort anchorColumn,
            out byte[] tokens,
            out string? reason) {
            tokens = Array.Empty<byte>();
            reason = null;
            if (string.IsNullOrWhiteSpace(formula)) {
                if (required) {
                    reason = "data validation formulas are required for this validation type";
                    return false;
                }

                return true;
            }

            string? formulaReason;
            bool encoded = type == DataValidationValues.List
                ? LegacyXlsFormulaEncoder.TryEncodeListSourceWithRelativeReferenceAnchor(formula, formulaNameIndex, sheetIndex, anchorRow, anchorColumn, out tokens, out formulaReason)
                : LegacyXlsFormulaEncoder.TryEncodeWithRelativeReferenceAnchor(formula, formulaNameIndex, sheetIndex, anchorRow, anchorColumn, out tokens, out formulaReason);
            if (!encoded) {
                reason = "data validation formulas outside the native XLS formula subset: " + formulaReason;
                return false;
            }

            if (tokens.Length > ushort.MaxValue) {
                reason = "data validation formula token payload lengths outside BIFF8 limits";
                return false;
            }

            return true;
        }

        private static bool TryParseRanges(string? sequenceOfReferences, out IReadOnlyList<CellRange> ranges, out string? reason) {
            ranges = Array.Empty<CellRange>();
            reason = null;
            if (string.IsNullOrWhiteSpace(sequenceOfReferences)) {
                reason = "data validation ranges";
                return false;
            }

            string[] parts = sequenceOfReferences!.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0 || parts.Length > 432) {
                reason = "data validation range counts outside BIFF8 limits";
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
                        reason = "data validation ranges";
                        return false;
                    }

                    lastRow = firstRow;
                    lastColumn = firstColumn;
                }

                if (firstRow < 1 || firstColumn < 1 || lastRow > 65536 || lastColumn > 256) {
                    reason = "data validation ranges outside BIFF8 worksheet limits";
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

        private static byte[] BuildRulePayload(DataValidationRule rule) {
            using var stream = new MemoryStream();
            uint flags = rule.TypeCode
                | (rule.ErrorStyleCode << 4)
                | (rule.OperatorCode << 20);
            if (rule.ExplicitList) flags |= ExplicitListFlag;
            if (rule.AllowBlank) flags |= AllowBlankFlag;
            if (rule.SuppressDropDown) flags |= SuppressDropDownFlag;
            if (rule.ShowInputMessage) flags |= ShowInputMessageFlag;
            if (rule.ShowErrorMessage) flags |= ShowErrorMessageFlag;

            WriteUInt32(stream, flags);
            WriteUnicodeString(stream, rule.PromptTitle);
            WriteUnicodeString(stream, rule.ErrorTitle);
            WriteUnicodeString(stream, rule.Prompt);
            WriteUnicodeString(stream, rule.Error);
            WriteFormula(stream, rule.Formula1Tokens);
            WriteFormula(stream, rule.Formula2Tokens);
            WriteUInt16(stream, checked((ushort)rule.Ranges.Count));
            foreach (CellRange range in rule.Ranges) {
                WriteUInt16(stream, range.FirstRow);
                WriteUInt16(stream, range.LastRow);
                WriteUInt16(stream, range.FirstColumn);
                WriteUInt16(stream, range.LastColumn);
            }

            return stream.ToArray();
        }

        private static bool TryGetValidationTypeCode(DataValidationValues type, out uint code) {
            if (type == DataValidationValues.None) code = 0x00;
            else if (type == DataValidationValues.Whole) code = 0x01;
            else if (type == DataValidationValues.Decimal) code = 0x02;
            else if (type == DataValidationValues.List) code = 0x03;
            else if (type == DataValidationValues.Date) code = 0x04;
            else if (type == DataValidationValues.Time) code = 0x05;
            else if (type == DataValidationValues.TextLength) code = 0x06;
            else if (type == DataValidationValues.Custom) code = 0x07;
            else {
                code = 0;
                return false;
            }

            return true;
        }

        private static bool TryGetOperatorCode(DataValidationOperatorValues? value, out uint code) {
            if (!value.HasValue || value.Value == DataValidationOperatorValues.Between) code = 0x00;
            else if (value.Value == DataValidationOperatorValues.NotBetween) code = 0x01;
            else if (value.Value == DataValidationOperatorValues.Equal) code = 0x02;
            else if (value.Value == DataValidationOperatorValues.NotEqual) code = 0x03;
            else if (value.Value == DataValidationOperatorValues.GreaterThan) code = 0x04;
            else if (value.Value == DataValidationOperatorValues.LessThan) code = 0x05;
            else if (value.Value == DataValidationOperatorValues.GreaterThanOrEqual) code = 0x06;
            else if (value.Value == DataValidationOperatorValues.LessThanOrEqual) code = 0x07;
            else {
                code = 0;
                return false;
            }

            return true;
        }

        private static bool TryGetErrorStyleCode(DataValidationErrorStyleValues? value, out uint code) {
            if (!value.HasValue || value.Value == DataValidationErrorStyleValues.Stop) code = 0x00;
            else if (value.Value == DataValidationErrorStyleValues.Warning) code = 0x01;
            else if (value.Value == DataValidationErrorStyleValues.Information) code = 0x02;
            else {
                code = 0;
                return false;
            }

            return true;
        }

        private static bool IsExplicitList(DataValidationValues type, string formula) {
            return type == DataValidationValues.List && IsQuotedStringLiteral(formula);
        }

        private static bool IsQuotedStringLiteral(string formula) {
            string trimmed = formula.Trim();
            return trimmed.Length >= 2 && trimmed[0] == '"' && trimmed[trimmed.Length - 1] == '"';
        }

        private static void WriteFormula(Stream stream, byte[] tokens) {
            WriteUInt16(stream, checked((ushort)tokens.Length));
            WriteUInt16(stream, 0);
            stream.Write(tokens, 0, tokens.Length);
        }

        private static void WriteUnicodeString(Stream stream, string text) {
            WriteUInt16(stream, checked((ushort)text.Length));
            stream.WriteByte(0x01);
            byte[] textBytes = Encoding.Unicode.GetBytes(text);
            stream.Write(textBytes, 0, textBytes.Length);
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

        private readonly struct DataValidationRule {
            internal DataValidationRule(
                uint typeCode,
                uint operatorCode,
                uint errorStyleCode,
                bool explicitList,
                bool allowBlank,
                bool suppressDropDown,
                bool showInputMessage,
                bool showErrorMessage,
                string promptTitle,
                string errorTitle,
                string prompt,
                string error,
                byte[] formula1Tokens,
                byte[] formula2Tokens,
                IReadOnlyList<CellRange> ranges) {
                TypeCode = typeCode;
                OperatorCode = operatorCode;
                ErrorStyleCode = errorStyleCode;
                ExplicitList = explicitList;
                AllowBlank = allowBlank;
                SuppressDropDown = suppressDropDown;
                ShowInputMessage = showInputMessage;
                ShowErrorMessage = showErrorMessage;
                PromptTitle = promptTitle;
                ErrorTitle = errorTitle;
                Prompt = prompt;
                Error = error;
                Formula1Tokens = formula1Tokens;
                Formula2Tokens = formula2Tokens;
                Ranges = ranges;
            }

            internal uint TypeCode { get; }
            internal uint OperatorCode { get; }
            internal uint ErrorStyleCode { get; }
            internal bool ExplicitList { get; }
            internal bool AllowBlank { get; }
            internal bool SuppressDropDown { get; }
            internal bool ShowInputMessage { get; }
            internal bool ShowErrorMessage { get; }
            internal string PromptTitle { get; }
            internal string ErrorTitle { get; }
            internal string Prompt { get; }
            internal string Error { get; }
            internal byte[] Formula1Tokens { get; }
            internal byte[] Formula2Tokens { get; }
            internal IReadOnlyList<CellRange> Ranges { get; }
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
