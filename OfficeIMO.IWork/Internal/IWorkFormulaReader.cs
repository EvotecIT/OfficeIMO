using System.Globalization;

namespace OfficeIMO.IWork.Internal;

internal sealed class IWorkFormulaResult {
    internal IWorkFormulaResult(string text, bool isComplete) {
        Text = text;
        IsComplete = isComplete;
    }

    internal string Text { get; }
    internal bool IsComplete { get; }
}

internal static class IWorkFormulaReader {
    private const int PrimaryPrecedence = 10;
    private const int UnaryPrecedence = 6;
    private static readonly IReadOnlyDictionary<int, string> FunctionNames = new Dictionary<int, string> {
        [1] = "ABS",
        [7] = "AND",
        [15] = "AVERAGE",
        [22] = "COLUMN",
        [30] = "COUNT",
        [31] = "COUNTA",
        [32] = "COUNTBLANK",
        [33] = "COUNTIF",
        [39] = "DATE",
        [41] = "DAY",
        [52] = "FALSE",
        [53] = "FIND",
        [60] = "HOUR",
        [61] = "HYPERLINK",
        [62] = "IF",
        [63] = "INDEX",
        [77] = "LEN",
        [76] = "LEFT",
        [84] = "MAX",
        [86] = "MEDIAN",
        [87] = "MID",
        [88] = "MIN",
        [89] = "MINUTE",
        [97] = "NOW",
        [101] = "OR",
        [102] = "PI",
        [112] = "ROUND",
        [119] = "SECOND",
        [124] = "RIGHT",
        [168] = "SUM",
        [169] = "SUMIF",
        [212] = "DURATION"
    };

    internal static IWorkFormulaResult Render(IWorkWireMessage formula, int zeroBasedRow, int zeroBasedColumn,
        int maximumNodes, int maximumCharacters) {
        IWorkWireMessage? nodeArray = IWorkObjectIndex.TryGetMessage(formula, 1, out bool malformed);
        if (malformed || nodeArray == null) return new IWorkFormulaResult(string.Empty, false);
        IReadOnlyList<IWorkWireMessage> nodes = IWorkObjectIndex.TryGetMessages(nodeArray, 1, out malformed);
        if (malformed || nodes.Count == 0) return new IWorkFormulaResult(string.Empty, false);
        if (nodes.Count > maximumNodes) {
            throw new InvalidDataException($"An iWork formula exceeds the configured syntax-node limit of {maximumNodes}.");
        }

        var stack = new List<Operand>();
        bool complete = true;
        foreach (IWorkWireMessage node in nodes) {
            ulong rawType = node.GetUnsigned(1) ?? 0;
            int type = rawType <= int.MaxValue ? (int)rawType : -1;
            if (type < 0) complete = false;
            if (TryBinary(type, out string? symbol, out int precedence)) {
                Operand[] operands = Pop(stack, 2, ref complete);
                stack.Add(new Operand(Bound(Wrap(operands[0], precedence) + symbol
                    + Wrap(operands[1], precedence + 1), maximumCharacters, ref complete), precedence));
                continue;
            }

            switch (type) {
                case 13:
                case 14: {
                    Operand operand = Pop(stack, 1, ref complete)[0];
                    stack.Add(new Operand(Bound((type == 13 ? "-" : "+")
                        + Wrap(operand, UnaryPrecedence), maximumCharacters, ref complete), UnaryPrecedence));
                    break;
                }
                case 15: {
                    Operand operand = Pop(stack, 1, ref complete)[0];
                    stack.Add(new Operand(Bound(Wrap(operand, UnaryPrecedence) + "%",
                        maximumCharacters, ref complete), UnaryPrecedence));
                    break;
                }
                case 16: {
                    ulong rawFunctionIndex = node.GetUnsigned(2) ?? 0;
                    int functionIndex = rawFunctionIndex <= int.MaxValue ? (int)rawFunctionIndex : -1;
                    int argumentCount = BoundedCount(node.GetUnsigned(3), maximumNodes);
                    Operand[] arguments = Pop(stack, argumentCount, ref complete);
                    if (!FunctionNames.TryGetValue(functionIndex, out string? name)) {
                        name = "FUNCTION_" + functionIndex.ToString(CultureInfo.InvariantCulture);
                        complete = false;
                    }
                    stack.Add(new Operand(Call(name, arguments, maximumCharacters, ref complete), PrimaryPrecedence));
                    break;
                }
                case 17:
                    stack.Add(new Operand(FormatNumber(node, ref complete), PrimaryPrecedence));
                    break;
                case 18: {
                    ulong? value = node.GetUnsigned(5);
                    if (!value.HasValue) complete = false;
                    stack.Add(new Operand(value.GetValueOrDefault() != 0 ? "TRUE" : "FALSE", PrimaryPrecedence));
                    break;
                }
                case 19: {
                    string? value = node.GetString(6);
                    if (value == null) complete = false;
                    stack.Add(new Operand(Quote(value ?? string.Empty,
                        maximumCharacters, ref complete), PrimaryPrecedence));
                    break;
                }
                case 20:
                case 21:
                    stack.Add(new Operand(FormatNumber(node, ref complete), PrimaryPrecedence));
                    break;
                case 22:
                case 23:
                    stack.Add(new Operand(string.Empty, PrimaryPrecedence));
                    break;
                case 24:
                case 25: {
                    int count = BoundedCount(node.GetUnsigned(type == 24 ? 11 : 13), maximumNodes);
                    Operand[] items = Pop(stack, count, ref complete);
                    stack.Add(new Operand(Delimited(type == 24 ? "{" : "(", type == 24 ? "}" : ")",
                        items, maximumCharacters, ref complete), PrimaryPrecedence));
                    break;
                }
                case 27:
                case 28:
                case 36:
                case 63:
                case 64:
                case 65:
                    stack.Add(new Operand(RenderReference(node, zeroBasedRow, zeroBasedColumn, ref complete),
                        PrimaryPrecedence));
                    break;
                case 29:
                case 45: {
                    Operand[] range = Pop(stack, 2, ref complete);
                    stack.Add(new Operand(Bound(range[0].Text + ":" + range[1].Text,
                        maximumCharacters, ref complete), PrimaryPrecedence));
                    break;
                }
                case 30:
                case 46:
                    stack.Add(new Operand("#REF!", PrimaryPrecedence));
                    break;
                case 31: {
                    int argumentCount = BoundedCount(node.GetUnsigned(18), maximumNodes);
                    Operand[] arguments = Pop(stack, argumentCount, ref complete);
                    string name = node.GetString(17) ?? "UNKNOWN";
                    if (name.Length > maximumCharacters) name = "UNKNOWN";
                    stack.Add(new Operand(Call(name, arguments, maximumCharacters, ref complete), PrimaryPrecedence));
                    complete = false;
                    break;
                }
                case 32:
                case 33: {
                    Operand operand = Pop(stack, 1, ref complete)[0];
                    string whitespace = node.GetString(25) ?? string.Empty;
                    stack.Add(new Operand(Bound(type == 32 ? operand.Text + whitespace : whitespace + operand.Text,
                        maximumCharacters, ref complete), operand.Precedence));
                    break;
                }
                case 34:
                case 35:
                    break;
                case 67:
                    stack.Add(new Operand(RenderColonTract(node, zeroBasedRow, zeroBasedColumn, ref complete),
                        PrimaryPrecedence));
                    break;
                case 69: {
                    Operand[] operands = Pop(stack, 2, ref complete);
                    stack.Add(new Operand(Bound(operands[0].Text + " " + operands[1].Text,
                        maximumCharacters, ref complete), PrimaryPrecedence));
                    break;
                }
                default:
                    stack.Add(new Operand("NODE_" + rawType.ToString(CultureInfo.InvariantCulture), PrimaryPrecedence));
                    complete = false;
                    break;
            }
        }

        if (stack.Count != 1) return new IWorkFormulaResult(string.Empty, false);
        string text = stack[0].Text;
        return new IWorkFormulaResult(text.Length == 0 ? string.Empty : "=" + text, complete && text.Length > 0);
    }

    internal static bool TryReadAbsoluteRange(IWorkWireMessage formula,
        out int firstRow, out int firstColumn, out int lastRow, out int lastColumn) {
        firstRow = firstColumn = lastRow = lastColumn = 0;
        IWorkWireMessage? nodeArray = IWorkObjectIndex.TryGetMessage(formula, 1);
        bool malformed = false;
        IReadOnlyList<IWorkWireMessage> nodes = nodeArray == null
            ? Array.Empty<IWorkWireMessage>()
            : IWorkObjectIndex.TryGetMessages(nodeArray, 1, out malformed);
        if (nodeArray == null || malformed || nodes.Count == 0) return false;
        if (nodes[0].GetUnsigned(1) == 67) {
            IWorkWireMessage? tract = IWorkObjectIndex.TryGetMessage(nodes[0], 40);
            if (tract == null
                || !TryAbsoluteRange(tract, 4, out firstRow, out lastRow)
                || !TryAbsoluteRange(tract, 3, out firstColumn, out lastColumn)) return false;
            return firstRow >= 0 && firstColumn >= 0 && lastRow >= firstRow && lastColumn >= firstColumn;
        }
        if (nodes.Count < 3 || nodes[0].GetUnsigned(1) != 36
            || nodes[1].GetUnsigned(1) != 36 || nodes[2].GetUnsigned(1) != 29
            || !TryAbsoluteCell(nodes[0], out firstRow, out firstColumn)
            || !TryAbsoluteCell(nodes[1], out lastRow, out lastColumn)) return false;
        if (lastRow < firstRow) (firstRow, lastRow) = (lastRow, firstRow);
        if (lastColumn < firstColumn) (firstColumn, lastColumn) = (lastColumn, firstColumn);
        return true;
    }

    private static bool TryBinary(int type, out string? symbol, out int precedence) {
        (symbol, precedence) = type switch {
            1 => ("+", 3), 2 => ("-", 3), 3 => ("*", 4), 4 => ("/", 4), 5 => ("^", 5),
            6 => ("&", 2), 7 => (">", 1), 8 => (">=", 1), 9 => ("<", 1), 10 => ("<=", 1),
            11 => ("=", 1), 12 => ("<>", 1), _ => (null, 0)
        };
        return symbol != null;
    }

    private static string RenderReference(IWorkWireMessage node, int row, int column, ref bool complete) {
        IWorkWireMessage? columnMessage = IWorkObjectIndex.TryGetMessage(node, 26, out bool malformedColumn);
        IWorkWireMessage? rowMessage = IWorkObjectIndex.TryGetMessage(node, 27, out bool malformedRow);
        if (malformedColumn || malformedRow
            || node.HasBytes(26) && columnMessage == null
            || node.HasBytes(27) && rowMessage == null) {
            complete = false;
            return "#REF!";
        }
        if (columnMessage == null && rowMessage == null) {
            complete = false;
            return "#REF!";
        }
        int? resolvedColumn = ResolveCoordinate(columnMessage, column, out bool absoluteColumn);
        int? resolvedRow = ResolveCoordinate(rowMessage, row, out bool absoluteRow);
        if (resolvedColumn < 0 || resolvedRow < 0
            || resolvedColumn > 16_383 || resolvedRow > 1_048_575) {
            complete = false;
            return "#REF!";
        }
        string address = CellAddress(resolvedColumn, resolvedRow, absoluteColumn, absoluteRow);
        if (node.HasBytes(28)) {
            complete = false;
            return "OTHER_TABLE::" + address;
        }
        return address;
    }

    private static int? ResolveCoordinate(IWorkWireMessage? message, int origin, out bool absolute) {
        absolute = false;
        if (message == null) return null;
        ulong? raw = message.GetUnsigned(1);
        if (!raw.HasValue) return null;
        long decoded = (long)(raw.Value >> 1) ^ -((long)raw.Value & 1L);
        if (decoded < int.MinValue || decoded > int.MaxValue) return null;
        int value = (int)decoded;
        absolute = (message.GetUnsigned(2) ?? 0) != 0;
        long resolved = absolute ? value : (long)origin + value;
        return resolved is >= int.MinValue and <= int.MaxValue ? (int)resolved : null;
    }

    private static bool TryAbsoluteCell(IWorkWireMessage node, out int row, out int column) {
        row = column = 0;
        int? resolvedColumn = ResolveCoordinate(IWorkObjectIndex.TryGetMessage(node, 26), 0,
            out bool absoluteColumn);
        int? resolvedRow = ResolveCoordinate(IWorkObjectIndex.TryGetMessage(node, 27), 0,
            out bool absoluteRow);
        if (!absoluteColumn || !absoluteRow || !resolvedColumn.HasValue || !resolvedRow.HasValue
            || resolvedColumn.Value < 0 || resolvedRow.Value < 0) return false;
        column = resolvedColumn.Value;
        row = resolvedRow.Value;
        return true;
    }

    private static string RenderColonTract(IWorkWireMessage node, int row, int column, ref bool complete) {
        IWorkWireMessage? tract = IWorkObjectIndex.TryGetMessage(node, 40, out bool malformedTract);
        if (malformedTract || tract == null) {
            complete = false;
            return "#REF!";
        }
        if (!TryRange(tract, 3, 1, column, out int firstColumn, out int lastColumn, out bool absoluteColumn)
            || !TryRange(tract, 4, 2, row, out int firstRow, out int lastRow, out bool absoluteRow)) {
            complete = false;
            return "#REF!";
        }
        string first = CellAddress(firstColumn, firstRow, absoluteColumn, absoluteRow);
        string last = CellAddress(lastColumn, lastRow, absoluteColumn, absoluteRow);
        if (first == "#REF!" || last == "#REF!") complete = false;
        return first == last ? first : first + ":" + last;
    }

    private static bool TryRange(IWorkWireMessage tract, int absoluteField, int relativeField, int origin,
        out int first, out int last, out bool absolute) {
        if (TryAbsoluteRange(tract, absoluteField, out first, out last)) {
            absolute = true;
            return true;
        }
        absolute = false;
        IReadOnlyList<IWorkWireMessage> ranges = IWorkObjectIndex.TryGetMessages(tract, relativeField, out bool malformed);
        if (malformed || ranges.Count == 0) {
            first = last = 0;
            return false;
        }
        ulong rawBegin = ranges[0].GetUnsigned(1) ?? 0;
        ulong rawEnd = ranges[0].GetUnsigned(2) ?? rawBegin;
        if (rawBegin > uint.MaxValue || rawEnd > uint.MaxValue) {
            first = last = 0;
            return false;
        }
        int begin = unchecked((int)(uint)rawBegin);
        int end = unchecked((int)(uint)rawEnd);
        long resolvedFirst = (long)origin + begin;
        long resolvedLast = (long)origin + end;
        if (resolvedFirst < 0 || resolvedFirst > int.MaxValue
            || resolvedLast < resolvedFirst || resolvedLast > int.MaxValue) {
            first = last = 0;
            return false;
        }
        first = (int)resolvedFirst;
        last = (int)resolvedLast;
        return first >= 0 && last >= first;
    }

    private static bool TryAbsoluteRange(IWorkWireMessage tract, int field, out int first, out int last) {
        IReadOnlyList<IWorkWireMessage> ranges = IWorkObjectIndex.TryGetMessages(tract, field, out bool malformed);
        if (malformed || ranges.Count == 0
            || ranges[0].GetUnsigned(1) is not ulong rawFirst || rawFirst > int.MaxValue) {
            first = last = 0;
            return false;
        }
        ulong rawLast = ranges[0].GetUnsigned(2) ?? rawFirst;
        if (rawLast > int.MaxValue) {
            first = last = 0;
            return false;
        }
        first = (int)rawFirst;
        last = (int)rawLast;
        return true;
    }

    private static string CellAddress(int? column, int? row, bool absoluteColumn, bool absoluteRow) {
        if (column is < 0 or > 16_383 || row is < 0 or > 1_048_575) return "#REF!";
        string columnText = column.HasValue ? ColumnName(column.Value) : string.Empty;
        string rowText = row.HasValue ? checked(row.Value + 1).ToString(CultureInfo.InvariantCulture) : string.Empty;
        return (column.HasValue && absoluteColumn ? "$" : string.Empty) + columnText
            + (row.HasValue && absoluteRow ? "$" : string.Empty) + rowText;
    }

    private static string ColumnName(int column) {
        string result = string.Empty;
        int value = checked(column + 1);
        while (value > 0) {
            int remainder = (value - 1) % 26;
            result = (char)('A' + remainder) + result;
            value = (value - remainder - 1) / 26;
        }
        return result;
    }

    private static int BoundedCount(ulong? raw, int maximum) {
        ulong value = raw ?? 0;
        if (value > (ulong)maximum || value > int.MaxValue) {
            throw new InvalidDataException($"An iWork formula declares an argument count above the configured limit of {maximum}.");
        }
        return (int)value;
    }

    private static string Bound(string value, int maximum, ref bool complete) {
        if (value.Length <= maximum) return value;
        complete = false;
        return "#FORMULA!";
    }

    private static string Quote(string value, int maximum, ref bool complete) {
        if (value.Length > maximum) {
            complete = false;
            return "#FORMULA!";
        }
        return Bound("\"" + value.Replace("\"", "\"\"") + "\"", maximum, ref complete);
    }

    private static string Call(string name, IReadOnlyList<Operand> arguments,
        int maximum, ref bool complete) {
        if (name.Length >= maximum) {
            complete = false;
            return "#FORMULA!";
        }
        return Delimited(name + "(", ")", arguments, maximum, ref complete);
    }

    private static string Delimited(string prefix, string suffix, IReadOnlyList<Operand> values,
        int maximum, ref bool complete) {
        long length = (long)prefix.Length + suffix.Length + Math.Max(0, values.Count - 1);
        foreach (Operand value in values) {
            length += value.Text.Length;
            if (length > maximum) {
                complete = false;
                return "#FORMULA!";
            }
        }
        return prefix + string.Join(",", values.Select(value => value.Text)) + suffix;
    }

    private static string FormatNumber(IWorkWireMessage node, ref bool complete) {
        double? value = node.GetDouble(4) ?? node.GetDouble(7) ?? node.GetDouble(8);
        if (!value.HasValue || double.IsNaN(value.Value) || double.IsInfinity(value.Value)) {
            complete = false;
            return "#NUM!";
        }
        return value.Value.ToString("R", CultureInfo.InvariantCulture);
    }

    private static Operand[] Pop(List<Operand> stack, int count, ref bool complete) {
        if (count < 0 || count > 4096) throw new InvalidDataException("An iWork formula declares an invalid argument count.");
        var result = new Operand[count];
        int missing = Math.Max(0, count - stack.Count);
        if (missing > 0) complete = false;
        int available = count - missing;
        int start = stack.Count - available;
        for (int index = 0; index < missing; index++) result[index] = new Operand(string.Empty, PrimaryPrecedence);
        for (int index = 0; index < available; index++) result[missing + index] = stack[start + index];
        if (available > 0) stack.RemoveRange(start, available);
        return result;
    }

    private static string Wrap(Operand operand, int minimumPrecedence) =>
        operand.Precedence < minimumPrecedence ? "(" + operand.Text + ")" : operand.Text;

    private readonly struct Operand {
        internal Operand(string text, int precedence) {
            Text = text;
            Precedence = precedence;
        }

        internal string Text { get; }
        internal int Precedence { get; }
    }
}
