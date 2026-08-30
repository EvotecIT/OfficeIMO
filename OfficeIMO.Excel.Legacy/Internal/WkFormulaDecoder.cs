using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace OfficeIMO.Excel.Legacy;

internal static class WkFormulaDecoder {
    private const int ExcelFormulaCharacterLimit = 8192;
    private const int MaximumExpressionDepth = 128;

    private static readonly Dictionary<byte, FunctionInfo> Functions = new() {
        [0x1F] = new("NA", 0), [0x21] = new("ABS", 1), [0x22] = new("INT", 1), [0x23] = new("SQRT", 1),
        [0x24] = new("LOG10", 1), [0x25] = new("LN", 1), [0x26] = new("PI", 0), [0x27] = new("SIN", 1),
        [0x28] = new("COS", 1), [0x29] = new("TAN", 1), [0x2A] = new("ATAN2", 2), [0x2B] = new("ATAN", 1),
        [0x2C] = new("ASIN", 1), [0x2D] = new("ACOS", 1), [0x2E] = new("EXP", 1), [0x2F] = new("MOD", 2),
        [0x33] = new("FALSE", 0), [0x34] = new("TRUE", 0), [0x36] = new("DATE", 3), [0x3B] = new("IF", 3),
        [0x3C] = new("DAY", 1), [0x3D] = new("MONTH", 1), [0x3E] = new("YEAR", 1), [0x3F] = new("ROUND", 2),
        [0x40] = new("TIME", 3), [0x41] = new("HOUR", 1), [0x42] = new("MINUTE", 1), [0x43] = new("SECOND", 1),
        [0x46] = new("LEN", 1), [0x47] = new("VALUE", 1), [0x49] = new("MID", 3), [0x4A] = new("CHAR", 1),
        [0x50] = new("SUM", -1), [0x51] = new("AVERAGE", -1), [0x52] = new("COUNT", -1), [0x53] = new("MIN", -1), [0x54] = new("MAX", -1),
        [0x56] = new("NPV", 2), [0x57] = new("VAR", -1), [0x58] = new("STDEV", -1), [0x5A] = new("HLOOKUP", 3),
        [0x60] = new("INDEX", 3), [0x61] = new("COLUMNS", 1), [0x62] = new("ROWS", 1), [0x64] = new("UPPER", 1),
        [0x65] = new("LOWER", 1), [0x66] = new("LEFT", 2), [0x67] = new("RIGHT", 2), [0x69] = new("PROPER", 1), [0x6B] = new("TRIM", 1)
    };

    internal static bool TryDecode(byte[] data, int offset, int length, int currentRowZeroBased, int currentColumnZeroBased,
        OfficeLegacyImportLimits limits, int maxTextCharacters, out string? formula, out string error) {
        formula = null;
        error = string.Empty;
        int maximumCharacters = Math.Min(ExcelFormulaCharacterLimit, Math.Max(0, maxTextCharacters));
        int maximumNodes = Math.Min(ExcelFormulaCharacterLimit, limits.MaxItems);
        var stack = new Stack<ExpressionNode>();
        int end = offset + length;
        int cursor = offset;
        int nodeCount = 0;
        int tokenCount = 0;
        try {
            while (cursor < end) {
                if (++tokenCount > limits.MaxRecords) throw new InvalidDataException("Formula exceeds the configured token limit.");
                byte token = data[cursor++];
                switch (token) {
                    case 0x00: Push(stack, Literal(ReadDouble(data, ref cursor, end).ToString("R", CultureInfo.InvariantCulture), maximumCharacters), ref nodeCount, maximumNodes); break;
                    case 0x01: Push(stack, Literal(ReadReference(data, ref cursor, end, currentRowZeroBased, currentColumnZeroBased), maximumCharacters), ref nodeCount, maximumNodes); break;
                    case 0x02: {
                        ExpressionNode first = Literal(ReadReference(data, ref cursor, end, currentRowZeroBased, currentColumnZeroBased), maximumCharacters);
                        ExpressionNode last = Literal(ReadReference(data, ref cursor, end, currentRowZeroBased, currentColumnZeroBased), maximumCharacters);
                        Push(stack, Combine(ExpressionKind.Range, ":", new[] { first, last }, maximumCharacters), ref nodeCount, maximumNodes, 3);
                        break;
                    }
                    case 0x03:
                        if (stack.Count != 1 || cursor != end) throw new InvalidDataException("Formula terminator did not leave one complete expression.");
                        ExpressionNode result = stack.Pop();
                        var builder = new StringBuilder(result.RenderedLength);
                        result.Render(builder);
                        formula = builder.ToString();
                        return true;
                    case 0x04: Unary(stack, string.Empty, ref nodeCount, maximumNodes, maximumCharacters); break;
                    case 0x05:
                        Require(cursor, 2, end);
                        short integer = (short)(data[cursor] | (data[cursor + 1] << 8));
                        cursor += 2;
                        Push(stack, Literal(integer.ToString(CultureInfo.InvariantCulture), maximumCharacters), ref nodeCount, maximumNodes);
                        break;
                    case 0x06: {
                        int zero = Array.IndexOf(data, (byte)0, cursor, end - cursor);
                        if (zero < 0) throw new InvalidDataException("Formula string token has no terminator.");
                        string text = Encoding.ASCII.GetString(data, cursor, zero - cursor).Replace("\"", "\"\"");
                        Push(stack, Literal("\"" + text + "\"", maximumCharacters), ref nodeCount, maximumNodes);
                        cursor = zero + 1;
                        break;
                    }
                    case 0x08: Unary(stack, "-", ref nodeCount, maximumNodes, maximumCharacters); break;
                    case 0x16: Function(stack, "NOT", 1, ref nodeCount, maximumNodes, maximumCharacters); break;
                    case 0x17: Unary(stack, "+", ref nodeCount, maximumNodes, maximumCharacters); break;
                    case 0x09: Binary(stack, "+", ref nodeCount, maximumNodes, maximumCharacters); break;
                    case 0x0A: Binary(stack, "-", ref nodeCount, maximumNodes, maximumCharacters); break;
                    case 0x0B: Binary(stack, "*", ref nodeCount, maximumNodes, maximumCharacters); break;
                    case 0x0C: Binary(stack, "/", ref nodeCount, maximumNodes, maximumCharacters); break;
                    case 0x0D: Binary(stack, "^", ref nodeCount, maximumNodes, maximumCharacters); break;
                    case 0x0E: Binary(stack, "=", ref nodeCount, maximumNodes, maximumCharacters); break;
                    case 0x0F: Binary(stack, "<>", ref nodeCount, maximumNodes, maximumCharacters); break;
                    case 0x10: Binary(stack, "<=", ref nodeCount, maximumNodes, maximumCharacters); break;
                    case 0x11: Binary(stack, ">=", ref nodeCount, maximumNodes, maximumCharacters); break;
                    case 0x12: Binary(stack, "<", ref nodeCount, maximumNodes, maximumCharacters); break;
                    case 0x13: Binary(stack, ">", ref nodeCount, maximumNodes, maximumCharacters); break;
                    case 0x14: Function(stack, "AND", 2, ref nodeCount, maximumNodes, maximumCharacters); break;
                    case 0x15: Function(stack, "OR", 2, ref nodeCount, maximumNodes, maximumCharacters); break;
                    case 0x18: Binary(stack, "&", ref nodeCount, maximumNodes, maximumCharacters); break;
                    default:
                        if (!Functions.TryGetValue(token, out FunctionInfo? function)) throw new InvalidDataException($"Unsupported formula token 0x{token:X2}.");
                        int arity = function.Arity;
                        if (arity < 0) { Require(cursor, 1, end); arity = data[cursor++]; }
                        Function(stack, function.Name, arity, ref nodeCount, maximumNodes, maximumCharacters);
                        break;
                }
                if (stack.Count > limits.MaxItems) throw new InvalidDataException("Formula exceeds the configured expression-stack limit.");
            }
            throw new InvalidDataException("Formula token stream has no terminator.");
        } catch (InvalidDataException exception) {
            error = exception.Message;
            return false;
        }
    }

    private static ExpressionNode Literal(string value, int maximumCharacters) {
        if (value.Length > maximumCharacters) throw new InvalidDataException("Formula exceeds the supported rendered-character limit.");
        return new ExpressionNode(ExpressionKind.Literal, value, Array.Empty<ExpressionNode>(), value.Length, 1);
    }

    private static ExpressionNode Combine(ExpressionKind kind, string value, ExpressionNode[] children, int maximumCharacters) {
        int length = kind switch {
            ExpressionKind.Binary => 2 + value.Length,
            ExpressionKind.Range => value.Length,
            ExpressionKind.Function => value.Length + 2 + Math.Max(0, children.Length - 1),
            ExpressionKind.Unary => value.Length + 2,
            _ => value.Length
        };
        int depth = 1;
        foreach (ExpressionNode child in children) {
            if (child.RenderedLength > maximumCharacters - length) throw new InvalidDataException("Formula exceeds the supported rendered-character limit.");
            length += child.RenderedLength;
            depth = Math.Max(depth, child.Depth + 1);
        }
        if (depth > MaximumExpressionDepth) throw new InvalidDataException("Formula exceeds the supported expression-depth limit.");
        return new ExpressionNode(kind, value, children, length, depth);
    }

    private static void Push(Stack<ExpressionNode> stack, ExpressionNode expression, ref int nodeCount, int maximumNodes, int additionalNodes = 1) {
        if (additionalNodes > maximumNodes - nodeCount) throw new InvalidDataException("Formula exceeds the configured expression-node limit.");
        nodeCount += additionalNodes;
        stack.Push(expression);
    }

    private static void Unary(Stack<ExpressionNode> stack, string prefix, ref int nodeCount, int maximumNodes, int maximumCharacters) {
        ExpressionNode child = Pop(stack);
        Push(stack, Combine(ExpressionKind.Unary, prefix, new[] { child }, maximumCharacters), ref nodeCount, maximumNodes);
    }

    private static void Binary(Stack<ExpressionNode> stack, string op, ref int nodeCount, int maximumNodes, int maximumCharacters) {
        ExpressionNode right = Pop(stack);
        ExpressionNode left = Pop(stack);
        Push(stack, Combine(ExpressionKind.Binary, op, new[] { left, right }, maximumCharacters), ref nodeCount, maximumNodes);
    }

    private static void Function(Stack<ExpressionNode> stack, string name, int arity, ref int nodeCount, int maximumNodes, int maximumCharacters) {
        if (arity < 0 || stack.Count < arity) throw new InvalidDataException($"Formula function {name} has an invalid argument count.");
        var args = new ExpressionNode[arity];
        for (int index = arity - 1; index >= 0; index--) args[index] = stack.Pop();
        Push(stack, Combine(ExpressionKind.Function, name, args, maximumCharacters), ref nodeCount, maximumNodes);
    }

    private static ExpressionNode Pop(Stack<ExpressionNode> stack) {
        if (stack.Count == 0) throw new InvalidDataException("Formula token stack underflow.");
        return stack.Pop();
    }

    private static string ReadReference(byte[] data, ref int cursor, int end, int currentRow, int currentColumn) {
        Require(cursor, 4, end);
        ushort columnToken = (ushort)(data[cursor] | (data[cursor + 1] << 8));
        ushort rowToken = (ushort)(data[cursor + 2] | (data[cursor + 3] << 8));
        cursor += 4;
        bool relativeColumn = (columnToken & 0x8000) != 0;
        bool relativeRow = (rowToken & 0x8000) != 0;
        int column = relativeColumn ? currentColumn + unchecked((sbyte)(columnToken & 0xFF)) : columnToken & 0xFF;
        int row;
        if (relativeRow) {
            int delta = rowToken & 0x3FFF;
            if ((delta & 0x2000) != 0) delta -= 0x4000;
            row = currentRow + delta;
        } else row = rowToken & 0x1FFF;
        if (row < 0 || row >= 1048576 || column < 0 || column >= 16384) throw new InvalidDataException("Formula cell reference is outside the workbook model.");
        return (relativeColumn ? string.Empty : "$") + ColumnName(column + 1) + (relativeRow ? string.Empty : "$") + (row + 1).ToString(CultureInfo.InvariantCulture);
    }

    private static double ReadDouble(byte[] data, ref int cursor, int end) {
        Require(cursor, 8, end);
        double value;
        if (BitConverter.IsLittleEndian) value = BitConverter.ToDouble(data, cursor);
        else { var copy = new byte[8]; Buffer.BlockCopy(data, cursor, copy, 0, 8); Array.Reverse(copy); value = BitConverter.ToDouble(copy, 0); }
        cursor += 8;
        if (double.IsNaN(value) || double.IsInfinity(value)) throw new InvalidDataException("Formula numeric token is not finite.");
        return value;
    }

    private static void Require(int cursor, int count, int end) {
        if (cursor < 0 || count < 0 || cursor > end - count) throw new InvalidDataException("Truncated formula token payload.");
    }

    private static string ColumnName(int column) {
        var result = new StringBuilder();
        while (column > 0) { column--; result.Insert(0, (char)('A' + column % 26)); column /= 26; }
        return result.ToString();
    }

    private enum ExpressionKind { Literal, Unary, Binary, Function, Range }

    private sealed class ExpressionNode {
        internal ExpressionNode(ExpressionKind kind, string value, ExpressionNode[] children, int renderedLength, int depth) {
            Kind = kind; Value = value; Children = children; RenderedLength = renderedLength; Depth = depth;
        }
        internal ExpressionKind Kind { get; }
        internal string Value { get; }
        internal ExpressionNode[] Children { get; }
        internal int RenderedLength { get; }
        internal int Depth { get; }

        internal void Render(StringBuilder builder) {
            switch (Kind) {
                case ExpressionKind.Literal: builder.Append(Value); break;
                case ExpressionKind.Range: Children[0].Render(builder); builder.Append(Value); Children[1].Render(builder); break;
                case ExpressionKind.Unary: builder.Append(Value).Append('('); Children[0].Render(builder); builder.Append(')'); break;
                case ExpressionKind.Binary: builder.Append('('); Children[0].Render(builder); builder.Append(Value); Children[1].Render(builder); builder.Append(')'); break;
                case ExpressionKind.Function:
                    builder.Append(Value).Append('(');
                    for (int index = 0; index < Children.Length; index++) { if (index > 0) builder.Append(','); Children[index].Render(builder); }
                    builder.Append(')');
                    break;
            }
        }
    }

    private sealed class FunctionInfo {
        internal FunctionInfo(string name, int arity) { Name = name; Arity = arity; }
        internal string Name { get; }
        internal int Arity { get; }
    }
}
