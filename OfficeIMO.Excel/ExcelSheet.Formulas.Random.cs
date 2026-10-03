using System.Security.Cryptography;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        // Keep all bounds and results inside the consecutive exact integer range of double.
        // Rejection sampling avoids modulo bias without floating-point range scaling.
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
            const double maximumExactInteger = 9007199254740991d; // 2^53 - 1
            if (first != Math.Truncate(first) || last != Math.Truncate(last)
                || first < -maximumExactInteger || first > maximumExactInteger
                || last < -maximumExactInteger || last > maximumExactInteger) return false;
            if (first > last) { result = FormulaArgumentValue.Error("#NUM!"); return true; }
            double number = first;
            if (first != last) {
                ulong count = (ulong)((long)last - (long)first + 1);
                // 2^64 is not representable as ulong. Reject its remainder at the
                // bottom of the sample space, leaving a multiple of count outcomes.
                ulong threshold = unchecked(0UL - count) % count;
                byte[] bytes = new byte[8];
                ulong sample;
                using (RandomNumberGenerator generator = RandomNumberGenerator.Create()) {
                    do {
                        generator.GetBytes(bytes);
                        sample = BitConverter.ToUInt64(bytes, 0);
                    } while (sample < threshold);
                }
                // Add as integers first: a cross-zero interval can have an offset
                // above 2^53 even though its final value is exactly representable.
                number = (long)first + (long)(sample % count);
            }
            result = new FormulaArgumentValue(number, InvariantNumberText.Get(number));
            return true;
        }
    }
}
