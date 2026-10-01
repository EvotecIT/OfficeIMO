using System.Security.Cryptography;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        // Integer bounds within Int32 are the supported portable subset. Rejection
        // sampling avoids modulo bias, including intervals spanning the entire Int32 range.
        private bool TryEvaluateRandomBetweenValue(string args, out FormulaArgumentValue result) {
            result = default;
            IReadOnlyList<string> tokens = SplitFormulaArguments(args, preserveEmpty: true);
            if (tokens.Count != 2 || tokens.Any(string.IsNullOrWhiteSpace)
                || !TryResolveFormulaArgument(tokens[0], out FormulaArgumentValue lower)
                || !TryResolveFormulaArgument(tokens[1], out FormulaArgumentValue upper)
                || lower.IsUnresolvedFormula || upper.IsUnresolvedFormula) return false;
            if (lower.IsError) { result = lower; return true; }
            if (upper.IsError) { result = upper; return true; }
            if (!IsFiniteFormulaNumber(lower) || !IsFiniteFormulaNumber(upper)) return false;
            double first = lower.Number!.Value, last = upper.Number!.Value;
            if (first != Math.Truncate(first) || last != Math.Truncate(last)
                || first < int.MinValue || first > int.MaxValue || last < int.MinValue || last > int.MaxValue) return false;
            if (first > last) { result = FormulaArgumentValue.Error("#NUM!"); return true; }
            double number = first;
            if (first != last) {
                ulong count = (ulong)((long)last - (long)first + 1);
                const ulong universe = 1UL << 32;
                ulong limit = universe - universe % count;
                byte[] bytes = new byte[4];
                uint sample;
                using (RandomNumberGenerator generator = RandomNumberGenerator.Create()) {
                    do {
                        generator.GetBytes(bytes);
                        sample = BitConverter.ToUInt32(bytes, 0);
                    } while ((ulong)sample >= limit);
                }
                number = first + (double)((ulong)sample % count);
            }
            result = new FormulaArgumentValue(number, InvariantNumberText.Get(number));
            return true;
        }
    }
}
