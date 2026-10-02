using System.Globalization;

namespace OfficeIMO.IWork.Internal;

internal sealed class IWorkFormulaResult {
    internal IWorkFormulaResult(string text, bool isComplete, bool requiresTableBinding = false, bool hasBoundedBodyRanges = false) {
        Text = text;
        IsComplete = isComplete;
        RequiresTableBinding = requiresTableBinding;
        HasBoundedBodyRanges = hasBoundedBodyRanges;
    }

    internal string Text { get; }
    internal bool IsComplete { get; }
    internal bool RequiresTableBinding { get; }
    internal bool HasBoundedBodyRanges { get; }
}

internal static partial class IWorkFormulaReader {
    private const int PrimaryPrecedence = 10;
    private const int SumFunctionIndex = 168;
    private const int UnaryPrecedence = 6;
    // Native iWork function IDs, qualified against the pinned independent function-identity manifest.
    // These are not Excel/BIFF IDs; unqualified IDs must remain incomplete to protect editable output.
    private static readonly IReadOnlyDictionary<int, FunctionDefinition> Functions =
        new Dictionary<int, FunctionDefinition> {
        [1] = new("ABS", 1, 1),
        [7] = new("AND", 1, 255),
        [15] = new("AVERAGE", 1, 255),
        [16] = new("AVERAGEA", 1, 255),
        [17] = new("CEILING", 2, 2),
        [22] = new("COLUMN", 0, 1),
        [23] = new("COLUMNS", 1, 1),
        [25] = new("CONCATENATE", 1, 255),
        [30] = new("COUNT", 1, 255),
        [31] = new("COUNTA", 1, 255),
        [32] = new("COUNTBLANK", 1, 1),
        [33] = new("COUNTIF", 2, 2),
        [39] = new("DATE", 3, 3),
        [41] = new("DAY", 1, 1),
        [44] = new("DEGREES", 1, 1),
        [48] = new("EVEN", 1, 1),
        [49] = new("EXACT", 2, 2),
        [50] = new("EXP", 1, 1),
        [24] = new("COMBIN", 2, 2),
        [51] = new("FACT", 1, 1),
        [52] = new("FALSE", 0, 0),
        [53] = new("FIND", 2, 3),
        [55] = new("FLOOR", 2, 2),
        [60] = new("HOUR", 1, 1),
        [61] = new("HYPERLINK", 1, 2),
        [62] = new("IF", 2, 3),
        [63] = new("INDEX", 2, 4),
        [65] = new("INT", 1, 1),
        [69] = new("ISBLANK", 1, 1),
        [70] = new("ISERROR", 1, 1),
        [76] = new("LEFT", 1, 2),
        [77] = new("LEN", 1, 1),
        [78] = new("LN", 1, 1),
        [79] = new("LOG", 1, 2),
        [80] = new("LOG10", 1, 1),
        [82] = new("LOWER", 1, 1),
        [84] = new("MAX", 1, 255),
        [85] = new("MAXA", 1, 255),
        [86] = new("MEDIAN", 1, 255),
        [87] = new("MID", 3, 3),
        [88] = new("MIN", 1, 255),
        [89] = new("MINA", 1, 255),
        [90] = new("MINUTE", 1, 1),
        [92] = new("MOD", 2, 2),
        [96] = new("NOT", 1, 1),
        [97] = new("NOW", 0, 0),
        [100] = new("ODD", 1, 1),
        [101] = new("OFFSET", 3, 5),
        [102] = new("OR", 1, 255),
        [104] = new("PI", 0, 0),
        [107] = new("POWER", 2, 2),
        [112] = new("PROB", 3, 4),
        [113] = new("PRODUCT", 1, 255),
        [117] = new("RADIANS", 1, 1),
        [119] = new("RANDBETWEEN", 2, 2),
        [124] = new("RIGHT", 1, 2),
        [126] = new("ROUND", 2, 2),
        [127] = new("ROUNDDOWN", 2, 2),
        [128] = new("ROUNDUP", 2, 2),
        [129] = new("ROW", 0, 1),
        [130] = new("ROWS", 1, 1),
        [132] = new("SECOND", 1, 1),
        [133] = new("SIGN", 1, 1),
        [139] = new("SQRT", 1, 1),
        [145] = new("SUMIF", 2, 3),
        [147] = new("SUMSQ", 1, 255),
        [155] = new("TRIM", 1, 1),
        [156] = new("TRUE", 0, 0),
        [157] = new("TRUNC", 1, 2),
        [158] = new("UPPER", 1, 1),
        [SumFunctionIndex] = new("SUM", 1, 255),
        [235] = new("IFERROR", 2, 2),
        [304] = new("ISNUMBER", 1, 1),
        [305] = new("ISTEXT", 1, 1),
        [328] = new("TEXTJOIN", 3, 254)
    };

    internal static IWorkFormulaResult Render(IWorkWireMessage formula, int zeroBasedRow, int zeroBasedColumn,
        int maximumNodes, int maximumCharacters, IReadOnlyDictionary<Guid, IWorkFormulaTableBinding>? tableQualifiers = null, IWorkFormulaTableBinding? owningTable = null) {
        if (!TryReadNodes(formula, maximumNodes, out IReadOnlyList<IWorkWireMessage> nodes))
            return new IWorkFormulaResult(string.Empty, false);

        var stack = new List<Operand>();
        bool complete = true, boundedBodyRanges = false, needsLocalBodyBinding = false;
        foreach (IWorkWireMessage node in nodes) {
            if (node.FieldCount(1) != 1
                || node.HasUnexpectedWireKind(1, IWorkWireKind.Varint)) complete = false;
            ulong rawType = node.GetUnsigned(1) ?? 0;
            int type = rawType <= int.MaxValue ? (int)rawType : -1;
            if (type < 0) complete = false;
            if (TryBinary(type, out string? symbol, out int precedence)) {
                Operand[] operands = Pop(stack, 2, ref complete);
                int leftPrecedence = symbol == "^" ? precedence + 2 : precedence;
                stack.Add(new Operand(Bound(Wrap(operands[0], leftPrecedence) + symbol
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
                    if (node.FieldCount(2) > 1 || node.FieldCount(3) > 1
                        || node.HasUnexpectedWireKind(2, IWorkWireKind.Varint)
                        || node.HasUnexpectedWireKind(3, IWorkWireKind.Varint)) complete = false;
                    ulong rawFunctionIndex = node.GetUnsigned(2) ?? 0;
                    int functionIndex = rawFunctionIndex <= int.MaxValue ? (int)rawFunctionIndex : -1;
                    int argumentCount = BoundedCount(node.GetUnsigned(3), maximumNodes);
                    Operand[] arguments = Pop(stack, argumentCount, ref complete);
                    string name;
                    if (Functions.TryGetValue(functionIndex, out FunctionDefinition function)) {
                        name = function.Name;
                        if (argumentCount < function.MinimumArguments
                            || argumentCount > function.MaximumArguments) complete = false;
                    } else {
                        name = functionIndex == 212
                            ? "DURATION"
                            : "FUNCTION_" + functionIndex.ToString(CultureInfo.InvariantCulture);
                        complete = false;
                    }
                    stack.Add(new Operand(Call(name, arguments, maximumCharacters, ref complete), PrimaryPrecedence));
                    break;
                }
                case 17:
                    stack.Add(new Operand(FormatNumber(node, ref complete), PrimaryPrecedence));
                    break;
                case 18: {
                    if (node.FieldCount(5) > 1
                        || node.HasUnexpectedWireKind(5, IWorkWireKind.Varint)) complete = false;
                    ulong? value = node.GetUnsigned(5);
                    if (!value.HasValue) complete = false;
                    if (value > 1) complete = false;
                    stack.Add(new Operand(value.GetValueOrDefault() != 0 ? "TRUE" : "FALSE", PrimaryPrecedence));
                    break;
                }
                case 19: {
                    string? value = node.GetString(6, out bool valueComplete);
                    if (!valueComplete || value == null) complete = false;
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
                    int countField = type == 24 ? 11 : 13;
                    if (node.FieldCount(countField) > 1
                        || node.HasUnexpectedWireKind(countField, IWorkWireKind.Varint)) complete = false;
                    int count = BoundedCount(node.GetUnsigned(countField), maximumNodes);
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
                case 65: {
                    string reference = RenderReference(node, zeroBasedRow, zeroBasedColumn, maximumCharacters, tableQualifiers,
                        owningTable, type == 36, ref complete, ref boundedBodyRanges, ref needsLocalBodyBinding);
                    Guid? identifier = type == 36 && node.HasField(26) && node.HasField(27)
                        ? ReadReferencedTableIdentifier(node) : null;
                    string? qualifier = identifier.HasValue && tableQualifiers != null
                        && tableQualifiers.TryGetValue(identifier.Value, out IWorkFormulaTableBinding? target)
                        && reference.StartsWith(target.Qualifier, StringComparison.Ordinal) ? target.Qualifier : null;
                    stack.Add(new Operand(reference, PrimaryPrecedence, identifier, qualifier,
                        qualifier == null ? null : reference.Substring(qualifier.Length)));
                    break;
                }
                case 29:
                case 45: {
                    Operand[] range = Pop(stack, 2, ref complete);
                    string last = range[1].Text;
                    if (range[0].ReferenceQualifier != null || range[1].ReferenceQualifier != null) {
                        // Independently qualified endpoint pairs use the same target identity.
                        // A local or differently qualified endpoint must not become a guessed cross-table range.
                        if (range[0].ReferenceIdentifier == range[1].ReferenceIdentifier
                            && range[0].ReferenceQualifier != null
                            && range[0].ReferenceQualifier == range[1].ReferenceQualifier)
                            last = range[1].ReferenceAddress!;
                        else complete = false;
                    }
                    stack.Add(new Operand(Bound(range[0].Text + ":" + last,
                        maximumCharacters, ref complete), PrimaryPrecedence));
                    break;
                }
                case 30:
                case 46:
                    stack.Add(new Operand("#REF!", PrimaryPrecedence));
                    break;
                case 31: {
                    if (node.FieldCount(18) > 1
                        || node.HasUnexpectedWireKind(18, IWorkWireKind.Varint)) complete = false;
                    int argumentCount = BoundedCount(node.GetUnsigned(18), maximumNodes);
                    Operand[] arguments = Pop(stack, argumentCount, ref complete);
                    string name = node.GetString(17, out bool nameComplete) ?? "UNKNOWN";
                    if (!nameComplete) complete = false;
                    if (name.Length > maximumCharacters) name = "UNKNOWN";
                    stack.Add(new Operand(Call(name, arguments, maximumCharacters, ref complete), PrimaryPrecedence));
                    complete = false;
                    break;
                }
                case 32:
                case 33: {
                    Operand operand = Pop(stack, 1, ref complete)[0];
                    string whitespace = node.GetString(25, out bool whitespaceComplete) ?? string.Empty;
                    if (!whitespaceComplete) complete = false;
                    if (!whitespace.All(IsFormulaWhitespace)) {
                        whitespace = string.Empty;
                        complete = false;
                    }
                    stack.Add(new Operand(Bound(type == 32 ? operand.Text + whitespace : whitespace + operand.Text,
                        maximumCharacters, ref complete), operand.Precedence, operand.ReferenceIdentifier, operand.ReferenceQualifier,
                        operand.ReferenceAddress == null ? null : type == 32
                            ? operand.ReferenceAddress + whitespace : whitespace + operand.ReferenceAddress));
                    break;
                }
                case 34:
                case 35:
                    break;
                case 67:
                    stack.Add(new Operand(RenderColonTract(node, zeroBasedRow, zeroBasedColumn, maximumCharacters, tableQualifiers, owningTable, ref complete, ref boundedBodyRanges, ref needsLocalBodyBinding),
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

        bool requiresTableBinding = needsLocalBodyBinding || nodes.Any(node => node.HasField(28));
        if (stack.Count != 1) return new IWorkFormulaResult(string.Empty, false, requiresTableBinding, boundedBodyRanges);
        string text = stack[0].Text;
        return new IWorkFormulaResult(text.Length == 0 ? string.Empty : "=" + text, complete && text.Length > 0, requiresTableBinding, boundedBodyRanges);
    }

    internal static long MeasureRenderingOperations(IWorkWireMessage formula, int maximumNodes) {
        if (formula.FieldCount(1) != 1
            || formula.HasUnexpectedWireKind(1, IWorkWireKind.Bytes)
            || formula.GetBytes(1) is not byte[] nodeArrayBytes) return 1;
        int totalFieldCount;
        try {
            formula.CountNestedFields(nodeArrayBytes, 1, out totalFieldCount);
        } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
            return 1;
        }
        if (totalFieldCount > maximumNodes) {
            throw new InvalidDataException(
                $"An iWork formula exceeds the configured syntax-node limit of {maximumNodes}.");
        }
        long nodeCount = Math.Max(1, totalFieldCount);
        return checked(nodeCount * nodeCount);
    }

    private readonly struct FunctionDefinition {
        internal FunctionDefinition(string name, int minimumArguments, int maximumArguments) {
            Name = name;
            MinimumArguments = minimumArguments;
            MaximumArguments = maximumArguments;
        }

        internal string Name { get; }
        internal int MinimumArguments { get; }
        internal int MaximumArguments { get; }
    }

    internal static bool TryReadAbsoluteRange(IWorkWireMessage formula, int maximumNodes,
        out int firstRow, out int firstColumn, out int lastRow, out int lastColumn,
        IWorkWireMessage? owningTable = null) {
        firstRow = firstColumn = lastRow = lastColumn = 0;
        if (!TryReadNodes(formula, maximumNodes, out IReadOnlyList<IWorkWireMessage> nodes)) return false;
        if (nodes.Any(node => node.FieldCount(1) != 1
                || node.HasUnexpectedWireKind(1, IWorkWireKind.Varint)
                || node.HasField(28) && !ReferencesOwningTable(node, owningTable))) return false;
        if (nodes[0].GetUnsigned(1) == 67) {
            if (nodes.Count != 1) return false;
            IWorkWireMessage? tract = IWorkObjectIndex.TryGetMessage(nodes[0], 40);
            if (tract == null
                || !TryAbsoluteRange(tract, 4, out firstRow, out lastRow)
                || !TryAbsoluteRange(tract, 3, out firstColumn, out lastColumn)) return false;
            return firstRow >= 0 && firstColumn >= 0 && lastRow >= firstRow && lastColumn >= firstColumn;
        }
        bool hasSupportedRangeWrapper = nodes.Count == 4
            && nodes[3].GetUnsigned(1) == 16
            && nodes[3].FieldCount(2) <= 1
            && nodes[3].FieldCount(3) <= 1
            && !nodes[3].HasUnexpectedWireKind(2, IWorkWireKind.Varint)
            && !nodes[3].HasUnexpectedWireKind(3, IWorkWireKind.Varint)
            && nodes[3].GetUnsigned(2) == SumFunctionIndex
            && nodes[3].GetUnsigned(3) == 1;
        if (nodes.Count != 3 && !hasSupportedRangeWrapper) return false;
        if (nodes[0].GetUnsigned(1) != 36
            || nodes[1].GetUnsigned(1) != 36 || nodes[2].GetUnsigned(1) != 29
            || !TryAbsoluteCell(nodes[0], out firstRow, out firstColumn)
            || !TryAbsoluteCell(nodes[1], out lastRow, out lastColumn)) return false;
        if (lastRow < firstRow) (firstRow, lastRow) = (lastRow, firstRow);
        if (lastColumn < firstColumn) (firstColumn, lastColumn) = (lastColumn, firstColumn);
        return true;
    }

    private static bool TryReadNodes(IWorkWireMessage formula, int maximumNodes,
        out IReadOnlyList<IWorkWireMessage> nodes) {
        nodes = Array.Empty<IWorkWireMessage>();
        if (formula.FieldCount(1) != 1) return false;
        byte[]? nodeArrayBytes = formula.GetBytes(1);
        if (nodeArrayBytes == null || formula.HasUnexpectedWireKind(1, IWorkWireKind.Bytes)) return false;
        int nodeCount;
        int totalFieldCount;
        try {
            nodeCount = formula.CountNestedFields(nodeArrayBytes, 1,
                out totalFieldCount);
        } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
            return false;
        }
        if (totalFieldCount > maximumNodes) {
            throw new InvalidDataException(
                $"An iWork formula exceeds the configured syntax-node limit of {maximumNodes}.");
        }
        IWorkWireMessage nodeArray;
        try {
            nodeArray = formula.ParseNestedMessage(nodeArrayBytes);
        } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
            return false;
        }
        IReadOnlyList<IWorkWireMessage> parsed = IWorkObjectIndex.TryGetMessages(
            nodeArray, 1, out bool malformed);
        if (malformed || parsed.Count == 0 || parsed.Count != nodeCount) return false;
        nodes = parsed;
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

    private static string RenderReference(IWorkWireMessage node, int row, int column, int maximumCharacters,
        IReadOnlyDictionary<Guid, IWorkFormulaTableBinding>? tableQualifiers, IWorkFormulaTableBinding? owningTable, bool supportsCoordinateBinding,
        ref bool complete, ref bool boundedBodyRanges, ref bool needsLocalBodyBinding) {
        IWorkWireMessage? columnMessage = IWorkObjectIndex.TryGetMessage(node, 26, out bool malformedColumn);
        IWorkWireMessage? rowMessage = IWorkObjectIndex.TryGetMessage(node, 27, out bool malformedRow);
        if (malformedColumn || malformedRow
            || node.HasUnexpectedWireKind(26, IWorkWireKind.Bytes)
            || node.HasUnexpectedWireKind(27, IWorkWireKind.Bytes)
            || node.HasBytes(26) && columnMessage == null
            || node.HasBytes(27) && rowMessage == null) {
            complete = false;
            return "#REF!";
        }
        if (columnMessage == null && rowMessage == null) {
            complete = false;
            return "#REF!";
        }
        int? resolvedColumn = ResolveCoordinate(columnMessage, column,
            out bool absoluteColumn, out bool columnComplete);
        int? resolvedRow = ResolveCoordinate(rowMessage, row,
            out bool absoluteRow, out bool rowComplete);
        if (!columnComplete || !rowComplete
            || resolvedColumn < 0 || resolvedRow < 0
            || resolvedColumn > 16_383 || resolvedRow > 1_048_575) {
            complete = false;
            return "#REF!";
        }
        string address = CellAddress(resolvedColumn, resolvedRow, absoluteColumn, absoluteRow);
        if (resolvedColumn == null || resolvedRow == null) address += ":" + address;
        if (!supportsCoordinateBinding) {
            if (resolvedColumn == null || resolvedRow == null) complete = false;
            return PreserveUnresolvedTableReference(node, address, ref complete);
        }
        if (resolvedColumn == null || resolvedRow == null) needsLocalBodyBinding = true;
        return BindTableReference(node, resolvedColumn, resolvedColumn, resolvedRow, resolvedRow,
            absoluteColumn, absoluteColumn, absoluteRow, absoluteRow, maximumCharacters,
            tableQualifiers, owningTable, ref complete, ref boundedBodyRanges);
    }

    // Dropping a declared table identity would turn an external reference into a
    // complete local formula or merge. Keep unsupported or unresolved identities incomplete.
    private static string PreserveUnresolvedTableReference(IWorkWireMessage node, string address,
        ref bool complete) {
        if (!node.HasField(28)) return address;
        complete = false;
        return node.HasBytes(28) ? "OTHER_TABLE::" + address : "#REF!";
    }

    private static int? ResolveCoordinate(IWorkWireMessage? message, int origin, out bool absolute,
        out bool complete) {
        absolute = false;
        complete = true;
        if (message == null) return null;
        if (message.FieldCount(1) != 1 || message.FieldCount(2) > 1
            || message.HasUnexpectedWireKind(1, IWorkWireKind.Varint)
            || message.HasUnexpectedWireKind(2, IWorkWireKind.Varint)) {
            complete = false;
            return null;
        }
        ulong? raw = message.GetUnsigned(1);
        if (!raw.HasValue) {
            complete = false;
            return null;
        }
        long decoded = (long)(raw.Value >> 1) ^ -((long)raw.Value & 1L);
        if (decoded < int.MinValue || decoded > int.MaxValue) {
            complete = false;
            return null;
        }
        int value = (int)decoded;
        ulong? rawAbsolute = message.GetUnsigned(2);
        if (rawAbsolute > 1) {
            complete = false;
            return null;
        }
        absolute = rawAbsolute == 1;
        long resolved = absolute ? value : (long)origin + value;
        if (resolved is >= int.MinValue and <= int.MaxValue) return (int)resolved;
        complete = false;
        return null;
    }

    private static bool TryAbsoluteCell(IWorkWireMessage node, out int row, out int column) {
        row = column = 0;
        int? resolvedColumn = ResolveCoordinate(IWorkObjectIndex.TryGetMessage(node, 26), 0,
            out bool absoluteColumn, out bool columnComplete);
        int? resolvedRow = ResolveCoordinate(IWorkObjectIndex.TryGetMessage(node, 27), 0,
            out bool absoluteRow, out bool rowComplete);
        if (!columnComplete || !rowComplete || !absoluteColumn || !absoluteRow
            || !resolvedColumn.HasValue || !resolvedRow.HasValue
            || resolvedColumn.Value < 0 || resolvedRow.Value < 0) return false;
        column = resolvedColumn.Value;
        row = resolvedRow.Value;
        return true;
    }

    private static bool TryAbsoluteRange(IWorkWireMessage tract, int field, out int first, out int last) {
        IReadOnlyList<IWorkWireMessage> ranges = IWorkObjectIndex.TryGetMessages(tract, field, out bool malformed);
        if (malformed || ranges.Count != 1
            || ranges[0].FieldCount(1) != 1 || ranges[0].FieldCount(2) > 1
            || ranges[0].HasUnexpectedWireKind(1, IWorkWireKind.Varint)
            || ranges[0].HasUnexpectedWireKind(2, IWorkWireKind.Varint)
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

    private static bool IsFormulaWhitespace(char value) =>
        value is ' ' or '\t' or '\r' or '\n';

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
        int numericFieldCount = node.FieldCount(4) + node.FieldCount(7) + node.FieldCount(8);
        if (numericFieldCount != 1
            || node.HasUnexpectedWireKind(4, IWorkWireKind.Fixed64)
            || node.HasUnexpectedWireKind(7, IWorkWireKind.Fixed64)
            || node.HasUnexpectedWireKind(8, IWorkWireKind.Fixed64)) complete = false;
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
        internal Operand(string text, int precedence, Guid? referenceIdentifier = null, string? referenceQualifier = null,
            string? referenceAddress = null) {
            Text = text;
            Precedence = precedence;
            ReferenceIdentifier = referenceIdentifier;
            ReferenceQualifier = referenceQualifier;
            ReferenceAddress = referenceAddress;
        }

        internal string Text { get; }
        internal int Precedence { get; }
        internal Guid? ReferenceIdentifier { get; }
        internal string? ReferenceQualifier { get; }
        internal string? ReferenceAddress { get; }
    }
}
