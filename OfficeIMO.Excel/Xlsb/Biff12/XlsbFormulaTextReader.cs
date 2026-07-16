using OfficeIMO.Excel.LegacyXls.Biff;
using System.Globalization;

namespace OfficeIMO.Excel.Xlsb.Biff12 {
    /// <summary>
    /// Decodes the common scalar BIFF12 formula-token subset used by normal worksheet formulas.
    /// Unsupported tokens return false so callers can retain the token bytes and cached value.
    /// </summary>
    internal static class XlsbFormulaTextReader {
        internal static bool TryRead(byte[] tokens, out string? formulaText) {
            if (tokens == null) throw new ArgumentNullException(nameof(tokens));

            formulaText = null;
            var stack = new Stack<FormulaExpression>();
            int offset = 0;
            while (offset < tokens.Length) {
                byte token = tokens[offset++];
                switch (token) {
                    case 0x03:
                        if (!ApplyBinary(stack, "+", 1)) return false;
                        break;
                    case 0x04:
                        if (!ApplyBinary(stack, "-", 1)) return false;
                        break;
                    case 0x05:
                        if (!ApplyBinary(stack, "*", 2)) return false;
                        break;
                    case 0x06:
                        if (!ApplyBinary(stack, "/", 2)) return false;
                        break;
                    case 0x07:
                        if (!ApplyBinary(stack, "^", 3)) return false;
                        break;
                    case 0x08:
                        if (!ApplyBinary(stack, "&", 0)) return false;
                        break;
                    case 0x09:
                        if (!ApplyBinary(stack, "<", 0)) return false;
                        break;
                    case 0x0A:
                        if (!ApplyBinary(stack, "<=", 0)) return false;
                        break;
                    case 0x0B:
                        if (!ApplyBinary(stack, "=", 0)) return false;
                        break;
                    case 0x0C:
                        if (!ApplyBinary(stack, ">=", 0)) return false;
                        break;
                    case 0x0D:
                        if (!ApplyBinary(stack, ">", 0)) return false;
                        break;
                    case 0x0E:
                        if (!ApplyBinary(stack, "<>", 0)) return false;
                        break;
                    case 0x0F:
                        if (!ApplyBinary(stack, " ", 0)) return false;
                        break;
                    case 0x10:
                        if (!ApplyBinary(stack, ",", 0)) return false;
                        break;
                    case 0x11:
                        if (!ApplyBinary(stack, ":", 0)) return false;
                        break;
                    case 0x12:
                    case 0x13:
                        if (!ApplyUnary(stack, token == 0x12 ? "+" : "-")) return false;
                        break;
                    case 0x14:
                        if (stack.Count < 1) return false;
                        FormulaExpression percent = stack.Pop();
                        stack.Push(new FormulaExpression(Parenthesize(percent, 4) + "%", 4));
                        break;
                    case 0x15:
                        if (stack.Count < 1) return false;
                        stack.Push(new FormulaExpression("(" + stack.Pop().Text + ")", 4));
                        break;
                    case 0x16:
                        stack.Push(new FormulaExpression(string.Empty, 4));
                        break;
                    case 0x17:
                        if (!TryReadFormulaString(tokens, ref offset, out string? text)) return false;
                        stack.Push(new FormulaExpression(QuoteFormulaString(text!), 4));
                        break;
                    case 0x19:
                        if (!TryApplyAttribute(tokens, ref offset, stack)) return false;
                        break;
                    case 0x1C:
                        if (offset >= tokens.Length) return false;
                        stack.Push(new FormulaExpression(BiffErrorValue.ToText(tokens[offset++]), 4));
                        break;
                    case 0x1D:
                        if (offset >= tokens.Length) return false;
                        stack.Push(new FormulaExpression(tokens[offset++] == 0 ? "FALSE" : "TRUE", 4));
                        break;
                    case 0x1E:
                        if (!TryReadUInt16(tokens, ref offset, out ushort integer)) return false;
                        stack.Push(new FormulaExpression(integer.ToString(CultureInfo.InvariantCulture), 4));
                        break;
                    case 0x1F:
                        if (!TryReadDouble(tokens, ref offset, out double number)) return false;
                        stack.Push(new FormulaExpression(number.ToString("G15", CultureInfo.InvariantCulture), 4));
                        break;
                    case 0x21:
                    case 0x41:
                    case 0x61:
                        if (!TryReadUInt16(tokens, ref offset, out ushort fixedFunction)
                            || !ApplyFixedFunction(stack, fixedFunction)) return false;
                        break;
                    case 0x22:
                    case 0x42:
                    case 0x62:
                        if (offset >= tokens.Length) return false;
                        byte parameterCount = tokens[offset++];
                        if (!TryReadUInt16(tokens, ref offset, out ushort functionBits)
                            || !ApplyVariableFunction(stack, parameterCount, functionBits)) return false;
                        break;
                    case 0x24:
                    case 0x44:
                    case 0x64:
                        if (!TryReadReference(tokens, ref offset, out string? reference)) return false;
                        stack.Push(new FormulaExpression(reference!, 4));
                        break;
                    case 0x25:
                    case 0x45:
                    case 0x65:
                        if (!TryReadArea(tokens, ref offset, out string? area)) return false;
                        stack.Push(new FormulaExpression(area!, 4));
                        break;
                    default:
                        return false;
                }
            }

            if (stack.Count != 1) return false;
            formulaText = stack.Pop().Text;
            return !string.IsNullOrWhiteSpace(formulaText);
        }

        private static bool TryReadReference(byte[] tokens, ref int offset, out string? reference) {
            reference = null;
            if (!TryReadUInt32(tokens, ref offset, out uint row)
                || !TryReadUInt16(tokens, ref offset, out ushort columnBits)) return false;
            if (row >= 1_048_576) return false;

            int column = columnBits & 0x3FFF;
            if (column >= 16_384) return false;
            bool columnRelative = (columnBits & 0x4000) != 0;
            bool rowRelative = (columnBits & 0x8000) != 0;
            reference = FormatReference(checked((int)row), column, rowRelative, columnRelative);
            return true;
        }

        private static bool TryReadFormulaString(byte[] tokens, ref int offset, out string? value) {
            value = null;
            if (!TryReadUInt16(tokens, ref offset, out ushort characterCount) || characterCount > 255) return false;
            int byteCount = checked(characterCount * 2);
            if (offset > tokens.Length - byteCount) return false;
            value = Encoding.Unicode.GetString(tokens, offset, byteCount);
            offset += byteCount;
            return true;
        }

        private static string QuoteFormulaString(string value) {
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }

        private static bool TryApplyAttribute(byte[] tokens, ref int offset, Stack<FormulaExpression> stack) {
            if (offset >= tokens.Length) return false;
            byte attribute = tokens[offset++];
            switch (attribute) {
                case 0x01: // PtgAttrSemi
                case 0x02: // PtgAttrIf
                case 0x08: // PtgAttrGoTo
                case 0x40: // PtgAttrSpace
                case 0x41: // PtgAttrSpaceSemi
                case 0x80: // PtgAttrIfError
                    return TrySkip(tokens, ref offset, 2);
                case 0x04: // PtgAttrChoose
                    if (!TryReadUInt16(tokens, ref offset, out ushort offsetCount)) return false;
                    return TrySkip(tokens, ref offset, checked((offsetCount + 1) * 2));
                case 0x10: // PtgAttrSum
                    return TrySkip(tokens, ref offset, 2) && ApplyFunction(stack, "SUM", 1);
                default:
                    return false;
            }
        }

        private static bool TryReadArea(byte[] tokens, ref int offset, out string? area) {
            area = null;
            if (!TryReadUInt32(tokens, ref offset, out uint firstRow)
                || !TryReadUInt32(tokens, ref offset, out uint lastRow)
                || !TryReadUInt16(tokens, ref offset, out ushort firstColumnBits)
                || !TryReadUInt16(tokens, ref offset, out ushort lastColumnBits)) return false;
            if (firstRow >= 1_048_576 || lastRow >= 1_048_576) return false;

            int firstColumn = firstColumnBits & 0x3FFF;
            int lastColumn = lastColumnBits & 0x3FFF;
            if (firstColumn >= 16_384 || lastColumn >= 16_384) return false;
            string first = FormatReference(
                checked((int)firstRow),
                firstColumn,
                (firstColumnBits & 0x8000) != 0,
                (firstColumnBits & 0x4000) != 0);
            string last = FormatReference(
                checked((int)lastRow),
                lastColumn,
                (lastColumnBits & 0x8000) != 0,
                (lastColumnBits & 0x4000) != 0);
            area = first + ":" + last;
            return true;
        }

        private static string FormatReference(int zeroBasedRow, int zeroBasedColumn, bool rowRelative, bool columnRelative) {
            int columnNumber = zeroBasedColumn + 1;
            var columnName = new StringBuilder();
            while (columnNumber > 0) {
                columnNumber--;
                columnName.Insert(0, (char)('A' + (columnNumber % 26)));
                columnNumber /= 26;
            }

            return (columnRelative ? string.Empty : "$")
                + columnName
                + (rowRelative ? string.Empty : "$")
                + (zeroBasedRow + 1).ToString(CultureInfo.InvariantCulture);
        }

        private static bool ApplyFixedFunction(Stack<FormulaExpression> stack, ushort functionId) {
            if (!BiffFormulaFunctionMetadata.TryGetFixedFunctionMetadata(functionId, out string? name, out int count)
                || stack.Count < count) return false;
            return ApplyFunction(stack, name!, count);
        }

        private static bool ApplyVariableFunction(Stack<FormulaExpression> stack, byte count, ushort functionBits) {
            if ((functionBits & 0x8000) != 0) return false;
            ushort functionId = (ushort)(functionBits & 0x7FFF);
            if (!BiffFormulaFunctionMetadata.TryGetFunctionName(functionId, out string? name)
                || !BiffFormulaFunctionMetadata.IsSupportedVariableFunctionArgumentCount(functionId, count)
                || stack.Count < count) return false;
            return ApplyFunction(stack, name!, count);
        }

        private static bool ApplyFunction(Stack<FormulaExpression> stack, string name, int count) {
            string[] arguments = new string[count];
            for (int index = count - 1; index >= 0; index--) {
                arguments[index] = stack.Pop().Text;
            }

            stack.Push(new FormulaExpression(name + "(" + string.Join(",", arguments) + ")", 4));
            return true;
        }

        private static bool ApplyBinary(Stack<FormulaExpression> stack, string operation, int precedence) {
            if (stack.Count < 2) return false;
            FormulaExpression right = stack.Pop();
            FormulaExpression left = stack.Pop();
            stack.Push(new FormulaExpression(
                Parenthesize(left, precedence) + operation + Parenthesize(right, precedence + 1),
                precedence));
            return true;
        }

        private static bool ApplyUnary(Stack<FormulaExpression> stack, string operation) {
            if (stack.Count < 1) return false;
            FormulaExpression value = stack.Pop();
            stack.Push(new FormulaExpression(operation + Parenthesize(value, 4), 4));
            return true;
        }

        private static string Parenthesize(FormulaExpression expression, int requiredPrecedence) {
            return expression.Precedence < requiredPrecedence
                ? "(" + expression.Text + ")"
                : expression.Text;
        }

        private static bool TryReadUInt16(byte[] data, ref int offset, out ushort value) {
            if (offset > data.Length - 2) {
                value = 0;
                return false;
            }

            value = (ushort)(data[offset] | (data[offset + 1] << 8));
            offset += 2;
            return true;
        }

        private static bool TryReadUInt32(byte[] data, ref int offset, out uint value) {
            if (offset > data.Length - 4) {
                value = 0;
                return false;
            }

            value = (uint)(data[offset]
                | (data[offset + 1] << 8)
                | (data[offset + 2] << 16)
                | (data[offset + 3] << 24));
            offset += 4;
            return true;
        }

        private static bool TryReadDouble(byte[] data, ref int offset, out double value) {
            if (offset > data.Length - 8) {
                value = 0;
                return false;
            }

            byte[] bytes = new byte[8];
            Buffer.BlockCopy(data, offset, bytes, 0, bytes.Length);
            value = BitConverter.ToDouble(bytes, 0);
            offset += bytes.Length;
            return true;
        }

        private static bool TrySkip(byte[] data, ref int offset, int count) {
            if (count < 0 || offset > data.Length - count) return false;
            offset += count;
            return true;
        }

        private readonly struct FormulaExpression {
            internal FormulaExpression(string text, int precedence) {
                Text = text;
                Precedence = precedence;
            }

            internal string Text { get; }

            internal int Precedence { get; }
        }
    }
}
