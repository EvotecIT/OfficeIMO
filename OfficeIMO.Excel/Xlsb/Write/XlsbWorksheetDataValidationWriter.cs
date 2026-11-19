using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Biff12;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Encodes worksheet data validations as a DVALS collection (MS-XLSB 2.4.55, 2.4.356, 2.5.36, 2.5.37, 2.5.98.8).</summary>
    internal static class XlsbWorksheetDataValidationWriter {
        internal const int BrtBeginDVals = 573;
        internal const int BrtEndDVals = 574;
        internal const int BrtDVal = 64;
        internal const int BrtDValList = 681;
        internal const int BrtACBegin = 37;
        internal const int BrtACEnd = 38;
        internal const int BrtUid = 3072;
        private const int MaximumValidations = 65534;
        private const int MaximumRanges = 8191;
        private const int MaximumFormulaBytes = 16384;
        private const uint MaximumWindowCoordinate = 65535;

        // BrtDVal bit fields from the least significant bit.
        internal const uint ExplicitListFlag = 0x00000080; // the "unused" bit 7, which Excel sets for an explicit list
        internal const uint AllowBlankFlag = 0x00000100;
        internal const uint SuppressComboFlag = 0x00000200;
        internal const uint ShowInputMessageFlag = 0x00040000;
        internal const uint ShowErrorMessageFlag = 0x00080000;

        /// <summary>valType values 0 to 7.</summary>
        internal static readonly DataValidationValues[] Types = {
            DataValidationValues.None, DataValidationValues.Whole, DataValidationValues.Decimal, DataValidationValues.List,
            DataValidationValues.Date, DataValidationValues.Time, DataValidationValues.TextLength, DataValidationValues.Custom,
        };

        /// <summary>errStyle values 0 to 2.</summary>
        internal static readonly DataValidationErrorStyleValues[] ErrorStyles = {
            DataValidationErrorStyleValues.Stop, DataValidationErrorStyleValues.Warning, DataValidationErrorStyleValues.Information,
        };

        /// <summary>mdImeMode values 0x00 to 0x0A.</summary>
        internal static readonly DataValidationImeModeValues[] ImeModes = {
            DataValidationImeModeValues.NoControl, DataValidationImeModeValues.On, DataValidationImeModeValues.Off,
            DataValidationImeModeValues.Disabled, DataValidationImeModeValues.Hiragana, DataValidationImeModeValues.FullKatakana,
            DataValidationImeModeValues.HalfKatakana, DataValidationImeModeValues.FullAlpha, DataValidationImeModeValues.HalfAlpha,
            DataValidationImeModeValues.FullHangul, DataValidationImeModeValues.HalfHangul,
        };

        /// <summary>typOperator values 0 to 7.</summary>
        internal static readonly DataValidationOperatorValues[] Operators = {
            DataValidationOperatorValues.Between, DataValidationOperatorValues.NotBetween, DataValidationOperatorValues.Equal,
            DataValidationOperatorValues.NotEqual, DataValidationOperatorValues.GreaterThan, DataValidationOperatorValues.LessThan,
            DataValidationOperatorValues.GreaterThanOrEqual, DataValidationOperatorValues.LessThanOrEqual,
        };

        private static readonly string[] CollectionAttributes = { "count", "disablePrompts", "xWindow", "yWindow" };

        private static readonly string[] ValidationAttributes = {
            "type", "errorStyle", "imeMode", "operator", "allowBlank", "showDropDown", "showInputMessage", "showErrorMessage",
            "errorTitle", "error", "promptTitle", "prompt", "sqref",
        };

        /// <summary>True for list, custom and "any value" validations, which take one formula and ignore typOperator.</summary>
        internal static bool IgnoresOperator(int valType) => valType is 0 or 3 or 7;

        internal static void Append(XlsbWorksheetFeatureRecords records, XlsbWorksheetFeatureContext context) =>
            records.DataValidations.AddRange(Create(context));

        internal static IReadOnlyList<XlsbGeneratedRecord> Create(XlsbWorksheetFeatureContext context) {
            DataValidations? collection = context.Worksheet.GetFirstChild<DataValidations>();
            if (collection == null) return Array.Empty<XlsbGeneratedRecord>();
            string sheetName = context.Sheet.Name;
            if (collection.ChildElements.Any(child => child is not DataValidation)
                || XlsbOpenXmlAttributeValidator.FindUnsupported(collection, CollectionAttributes) is OpenXmlAttribute) {
                throw new NotSupportedException($"Native XLSB generation does not support data-validation extensions on worksheet '{sheetName}'.");
            }
            DataValidation[] items = collection.Elements<DataValidation>().ToArray();
            // DVALS requires at least one BrtDVal; an empty collection has nothing to keep.
            if (items.Length == 0) return Array.Empty<XlsbGeneratedRecord>();
            if (items.Length > MaximumValidations) {
                throw new NotSupportedException($"Native XLSB generation supports at most {MaximumValidations} data validations on worksheet '{sheetName}'.");
            }
            uint xLeft = collection.XWindow?.Value ?? 0U;
            uint yTop = collection.YWindow?.Value ?? 0U;
            if (xLeft > MaximumWindowCoordinate || yTop > MaximumWindowCoordinate) {
                throw new NotSupportedException($"Native XLSB generation found an out-of-range data-validation prompt position on worksheet '{sheetName}'.");
            }

            var result = new List<XlsbGeneratedRecord>(items.Length + 2);
            using (var begin = new MemoryStream(18)) {
                XlsbPayloadWriter.WriteUInt16(begin, collection.DisablePrompts?.Value == true ? (ushort)1 : (ushort)0);
                XlsbPayloadWriter.WriteUInt32(begin, xLeft);
                XlsbPayloadWriter.WriteUInt32(begin, yTop);
                XlsbPayloadWriter.WriteUInt32(begin, 0U);
                XlsbPayloadWriter.WriteUInt32(begin, (uint)items.Length);
                result.Add(new XlsbGeneratedRecord(BrtBeginDVals, begin.ToArray()));
            }
            foreach (DataValidation item in items) result.Add(new XlsbGeneratedRecord(BrtDVal, CreateValidation(context, item)));
            result.Add(new XlsbGeneratedRecord(BrtEndDVals, Array.Empty<byte>()));
            return result;
        }

        private static byte[] CreateValidation(XlsbWorksheetFeatureContext context, DataValidation item) {
            string sheetName = context.Sheet.Name;
            if (item.ChildElements.Any(child => child is not Formula1 and not Formula2)
                || item.Elements<Formula1>().Count() > 1 || item.Elements<Formula2>().Count() > 1
                || item.ChildElements.Any(child => child.HasChildren || child.GetAttributes().Count != 0)
                || XlsbOpenXmlAttributeValidator.FindUnsupported(item, ValidationAttributes) is OpenXmlAttribute) {
                throw new NotSupportedException($"Native XLSB generation found an unsupported data-validation entry on worksheet '{sheetName}'.");
            }

            int valType = IndexOf(Types, item.Type, 0, "type", sheetName);
            int errStyle = IndexOf(ErrorStyles, item.ErrorStyle, 0, "error style", sheetName);
            int imeMode = IndexOf(ImeModes, item.ImeMode, 0, "IME mode", sheetName);
            int typOperator = IgnoresOperator(valType) ? 0 : IndexOf(Operators, item.Operator, 0, "operator", sheetName);
            string? formula1 = item.GetFirstChild<Formula1>()?.Text;
            string? formula2 = item.GetFirstChild<Formula2>()?.Text;
            bool hasFormula1 = !string.IsNullOrEmpty(formula1);
            bool hasFormula2 = !string.IsNullOrEmpty(formula2);
            bool twoFormulas = !IgnoresOperator(valType) && typOperator <= 1;
            if (valType == 0 ? hasFormula1 || hasFormula2 : !hasFormula1 || hasFormula2 != twoFormulas) {
                throw new NotSupportedException($"Native XLSB generation found a data validation on worksheet '{sheetName}' whose formulas do not match its type and operator.");
            }

            using var payload = new MemoryStream();
            bool explicitList = valType == 3 && IsStringLiteral(formula1!);
            uint flags = (uint)valType | (uint)errStyle << 4 | (uint)imeMode << 10 | (uint)typOperator << 20;
            if (explicitList) flags |= ExplicitListFlag;
            if (item.AllowBlank?.Value == true) flags |= AllowBlankFlag;
            // showDropDown="1" hides the in-cell drop-down, which is what fSuppressCombo = 1 means.
            if (item.ShowDropDown?.Value == true) flags |= SuppressComboFlag;
            if (item.ShowInputMessage?.Value == true) flags |= ShowInputMessageFlag;
            if (item.ShowErrorMessage?.Value == true) flags |= ShowErrorMessageFlag;
            XlsbPayloadWriter.WriteUInt32(payload, flags);

            string references = item.SequenceOfReferences?.InnerText ?? string.Empty;
            XlsbPayloadWriter.WriteSqRfX(payload, references, sheetName, "data-validation", MaximumRanges);
            string firstRange = references.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)[0];
            XlsbPayloadWriter.TryParseRange(firstRange, out int anchorRow, out int anchorColumn, out _, out _);

            WriteString(payload, item.ErrorTitle?.Value, 32, "error title", sheetName);
            WriteString(payload, item.Error?.Value, 225, "error message", sheetName);
            WriteString(payload, item.PromptTitle?.Value, 32, "prompt title", sheetName);
            WriteString(payload, item.Prompt?.Value, 255, "prompt", sheetName);

            var anchor = (Row: anchorRow - 1, Column: anchorColumn - 1);
            WriteFormula(payload, context, hasFormula1 ? formula1 : null, valType == 3 ? (explicitList ? FormulaRole.ExplicitList : FormulaRole.ListSource) : FormulaRole.Value, anchor);
            WriteFormula(payload, context, hasFormula2 ? formula2 : null, FormulaRole.Value, anchor);
            return payload.ToArray();
        }

        private enum FormulaRole { Value, ListSource, ExplicitList }

        private static void WriteFormula(MemoryStream payload, XlsbWorksheetFeatureContext context, string? formula, FormulaRole role, (int Row, int Column) anchor) {
            if (formula == null) {
                XlsbPayloadWriter.WriteUInt32(payload, 0U); // cce
                XlsbPayloadWriter.WriteUInt32(payload, 0U); // cb
                return;
            }
            byte resultClass = role == FormulaRole.ListSource ? ExcelNativeFormulaOperandTypes.Reference : ExcelNativeFormulaOperandTypes.Value;
            byte[] encoded = context.EncodeFormula(formula, "data-validation", resultClass);
            int cce = checked((int)BitConverter.ToUInt32(encoded, 0));
            if (cce < 1 || cce >= MaximumFormulaBytes) {
                throw new NotSupportedException($"Native XLSB generation cannot encode the data-validation formula '{formula}' on worksheet '{context.Sheet.Name}' within the DVParsedFormula size limit.");
            }
            string? reason = Validate(encoded, 4, cce, role);
            if (reason != null) {
                throw new NotSupportedException($"Native XLSB generation cannot encode the data-validation formula '{formula}' on worksheet '{context.Sheet.Name}': {reason}");
            }
            ConvertToRelative(encoded, 4, cce, anchor);
            payload.Write(encoded, 0, encoded.Length);
        }

        /// <summary>Applies the DVParsedFormula restrictions (MS-XLSB 2.5.98.8); returns why the tokens break them, or null.</summary>
        private static string? Validate(byte[] encoded, int start, int length, FormulaRole role) {
            var tokens = new List<(int Offset, byte Token, int Operand)>();
            try {
                XlsbFormulaTokenScanner.Visit(encoded, start, length, (offset, token, operand) => tokens.Add((offset, token, operand)));
            } catch (NotSupportedException) {
                return "the formula uses tokens a data validation cannot hold";
            }
            if (role == FormulaRole.ExplicitList) {
                // Excel stores an explicit list as one PtgStr holding the items exactly as typed, separators included.
                return tokens.Count == 1 && tokens[0].Token == 0x17 ? null : "an explicit list must be a single string constant";
            }
            bool has3dArea = false;
            foreach (var entry in tokens) {
                byte token = entry.Token;
                int kind = token & 0x1F;
                bool classed = token >= 0x20;
                if (token is 0x01 or 0x18 or 0x0F or 0x10
                    || classed && kind is 0x00 or 0x06 or 0x08 or 0x19 or 0x1A or 0x1C) {
                    return "data validation formulas cannot use array constants, intersections, unions, structured or external references, or single-cell references to other sheets";
                }
                if (classed && kind is 0x1B or 0x1D) has3dArea = true;
            }
            if (role == FormulaRole.ListSource) {
                if (has3dArea && tokens.Count != 1) return "a list that refers to another sheet must be that range alone";
                return IsValueType(encoded, tokens) ? "a list source must be a range, a name or a function that returns a reference" : null;
            }
            if (has3dArea) return "only a list source may refer to another sheet";
            if (tokens.Count == 1 && tokens[0].Token >= 0x20 && (tokens[0].Token & 0x1F) is 0x05 or 0x0B or 0x0D) {
                return "a range alone is not a value";
            }
            return IsValueType(encoded, tokens) ? null : "the formula must produce a value";
        }

        /// <summary>Whether the root of the parse tree is a VALUE_TYPE (MS-XLSB 2.5.98.88).</summary>
        private static bool IsValueType(byte[] encoded, List<(int Offset, byte Token, int Operand)> tokens) {
            for (int index = tokens.Count - 1; index >= 0; index--) {
                byte token = tokens[index].Token;
                // Display tokens: PtgParen and PtgAttrSpace/PtgAttrSpaceSemi.
                if (token == 0x15 || token == 0x19 && encoded[tokens[index].Operand] is 0x40 or 0x41) continue;
                // PtgIsect, PtgUnion and PtgRange combine references; other untyped tokens produce values.
                if (token is 0x0F or 0x10 or 0x11) return false;
                if (token < 0x20) return true;
                int kind = token & 0x1F;
                if (kind is 0x0A or 0x0B or 0x1C or 0x1D) return true;
                return (token & 0x60) != ExcelNativeFormulaOperandTypes.Reference;
            }
            return true;
        }

        /// <summary>
        /// Rewrites PtgRef/PtgArea with relative parts as PtgRefN/PtgAreaN offsets from the first cell of the first range,
        /// the form Excel uses for data validation formulas. Absolute references keep their coordinates.
        /// </summary>
        private static void ConvertToRelative(byte[] encoded, int start, int length, (int Row, int Column) anchor) {
            XlsbFormulaTokenScanner.Visit(encoded, start, length, (offset, token, operand) => {
                int kind = token & 0x1F;
                if (token < 0x20 || kind is not (0x04 or 0x05)) return;
                bool area = kind == 0x05;
                int columnsAt = operand + (area ? 8 : 4);
                bool relative = false;
                for (int part = 0; part < (area ? 2 : 1); part++) {
                    if ((ReadUInt16(encoded, columnsAt + part * 2) & 0xC000) != 0) relative = true;
                }
                if (!relative) return;
                for (int part = 0; part < (area ? 2 : 1); part++) {
                    int rowAt = operand + part * 4;
                    int columnAt = columnsAt + part * 2;
                    ushort columnBits = ReadUInt16(encoded, columnAt);
                    if ((columnBits & 0x8000) != 0) {
                        int row = checked((int)XlsbFormulaTokenScanner.ReadUInt32(encoded, rowAt));
                        XlsbFormulaTokenScanner.WriteUInt32(encoded, rowAt, unchecked((uint)(row - anchor.Row)));
                    }
                    if ((columnBits & 0x4000) != 0) {
                        int column = columnBits & 0x3FFF;
                        columnBits = (ushort)((columnBits & 0xC000) | ((column - anchor.Column) & 0x3FFF));
                        encoded[columnAt] = (byte)columnBits;
                        encoded[columnAt + 1] = (byte)(columnBits >> 8);
                    }
                }
                encoded[offset] = (byte)(token + 0x08); // PtgRef 0x04 -> PtgRefN 0x0C, PtgArea 0x05 -> PtgAreaN 0x0D
            });
        }

        private static ushort ReadUInt16(byte[] bytes, int offset) => (ushort)(bytes[offset] | bytes[offset + 1] << 8);

        private static bool IsStringLiteral(string formula) {
            string text = formula.StartsWith("=", StringComparison.Ordinal) ? formula.Substring(1) : formula;
            return text.TrimStart().StartsWith("\"", StringComparison.Ordinal);
        }

        private static void WriteString(MemoryStream payload, string? value, int maximumLength, string purpose, string sheetName) {
            if (value != null && value.Length > maximumLength) {
                throw new NotSupportedException($"Native XLSB generation supports data-validation {purpose}s of at most {maximumLength} characters on worksheet '{sheetName}'.");
            }
            XlsbPayloadWriter.WriteNullableWideString(payload, value);
        }

        private static int IndexOf<T>(T[] values, EnumValue<T>? value, int defaultIndex, string purpose, string sheetName) where T : struct, IEnumValue, IEnumValueFactory<T> {
            if (value == null || !value.HasValue) return defaultIndex;
            int index = Array.IndexOf(values, value.Value);
            if (index < 0) throw new NotSupportedException($"Native XLSB generation cannot encode the data-validation {purpose} '{value.InnerText}' on worksheet '{sheetName}'.");
            return index;
        }
    }
}
