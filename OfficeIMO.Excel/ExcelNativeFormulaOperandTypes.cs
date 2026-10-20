using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Model;
using OfficeIMO.Excel.Xlsb.Biff12;

namespace OfficeIMO.Excel {
    /// <summary>Assigns native reference/value/array classes from the formula expression and each function's argument contract.</summary>
    internal static partial class ExcelNativeFormulaOperandTypes {
        internal const byte Reference = 0x20, Value = 0x40, Array = 0x60;

        internal static bool TryApply(byte[] bytes, bool binary12, byte resultType, out string? reason) {
            reason = null;
            var tokens = new List<TypeToken>();
            if (binary12) {
                XlsbFormulaTokenScanner.Visit(bytes, 0, bytes.Length, (offset, token, operand) => {
                    ushort? function = null; int? count = null;
                    int kind = token & 0x1f;
                    if (token >= 0x20 && kind is 1 or 2) {
                        int at = operand + (kind == 2 ? 1 : 0);
                        function = (ushort)((bytes[at] | bytes[at + 1] << 8) & 0x7fff);
                        if (kind == 2) count = bytes[operand] & 0x7f;
                        else if (BiffFormulaFunctionMetadata.TryGetFixedFunctionMetadata(function.Value, out _, out int arity)) count = arity;
                    }
                    tokens.Add(new TypeToken(offset, token, function, count));
                });
            } else {
                var scanned = new List<LegacyXlsFormulaTokenRecord>();
                BiffFormulaTokenScanner.ScanTokens(bytes, "CellFormula", null, null, 0, 0x0006, scanned);
                if (scanned.Any(item => item.OperandByteCount == null)
                    || scanned.Count == 0 || scanned.Last().TokenOffset + 1 + scanned.Last().OperandByteCount != bytes.Length) {
                    reason = "The native formula operand layout cannot be typed safely."; return false;
                }
                tokens.AddRange(scanned.Select(item => new TypeToken(item.TokenOffset, item.Token, item.FunctionId, item.FunctionParameterCount)));
            }
            var stack = new Stack<Node>();
            foreach (TypeToken token in tokens) {
                if (token.Code == 0x19) continue; // IF/CHOOSE jumps and volatility do not consume expression operands.
                int count = token.Function.HasValue ? token.Count ?? -1
                    : token.Code is >= 0x03 and <= 0x11 ? 2 : token.Code is >= 0x12 and <= 0x15 ? 1 : 0;
                if (count < 0 || count > stack.Count) { reason = "The native function argument layout cannot be typed safely."; return false; }
                var children = new Node[count];
                for (int index = count - 1; index >= 0; index--) children[index] = stack.Pop();
                stack.Push(new Node(token, children));
            }
            if (stack.Count != 1) { reason = "The native formula expression cannot be typed safely."; return false; }
            // Walk iteratively: formula size, rather than CLR recursion depth, bounds this operation.
            var pending = new Stack<(Node Node, byte Type, bool ArrayEvaluation)>(); pending.Push((stack.Pop(), resultType, false));
            while (pending.Count != 0) {
                var work = pending.Pop(); Node node = work.Node; byte required = work.Type;
                TypeToken token = node.Token; int kind = token.Code & 0x1f;
                if (token.Function.HasValue) {
                    string name = BiffFormulaFunctionMetadata.TryGetKnownFunctionName(token.Function.Value, out string? known) ? known! : string.Empty;
                    byte output = required == Reference ? ReturnsReference(name) ? Reference : ReturnsArray(name) ? Array : Value : required;
                    // Array evaluation survives scalar functions and value operators even when
                    // an aggregate requests its argument as a reference. Reference-returning
                    // functions keep scalar address/size arguments outside that evaluation.
                    bool scalarFunction = !ReturnsReference(name) && !ReturnsArray(name);
                    bool conditionalFunction = name is "IF" or "CHOOSE";
                    bool arrayEvaluation = (scalarFunction || conditionalFunction) && (work.ArrayEvaluation || output == Array);
                    if ((scalarFunction || conditionalFunction) && work.ArrayEvaluation) output = Array;
                    bytes[token.Offset] = (byte)(kind | output);
                    for (int index = 0; index < node.Children.Length; index++) {
                        byte argument = ArgumentType(name, index);
                        if (argument == 0) argument = output; // IF/CHOOSE result branches follow the requested result class.
                        else if (argument == Value && output == Array && (scalarFunction || conditionalFunction)) argument = Array;
                        pending.Push((node.Children[index], argument, arrayEvaluation));
                    }
                } else if (node.Children.Length != 0) {
                    byte operand = token.Code == 0x15 ? required : token.Code is 0x0f or 0x10 or 0x11 ? Reference : required == Array ? Array : Value;
                    foreach (Node child in node.Children) pending.Push((child, operand, work.ArrayEvaluation));
                } else if (token.Code >= 0x20 && kind is 0x03 or 0x04 or 0x05 or 0x0a or 0x0b or 0x0c or 0x0d or 0x19 or 0x1a or 0x1b or 0x1c or 0x1d) {
                    bytes[token.Offset] = (byte)(kind | (required == Value && work.ArrayEvaluation ? Array : required));
                }
            }
            return true;
        }

        private sealed record TypeToken(int Offset, byte Code, ushort? Function, int? Count);
        private sealed record Node(TypeToken Token, Node[] Children);
    }
}
