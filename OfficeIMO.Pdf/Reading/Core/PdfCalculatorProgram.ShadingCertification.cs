namespace OfficeIMO.Pdf;

internal sealed partial class PdfCalculatorProgram {
    /// <summary>
    /// Proves that every output is monotonic over a one-dimensional domain.
    /// Only exact quadratic-or-lower arithmetic is accepted; opaque or
    /// discontinuous operators deliberately fail closed for strict shading.
    /// </summary>
    internal bool TryCertifyOneInputMonotonicOutput(double domainMinimum, double domainMaximum, int outputCount) {
        if (HasConditional || outputCount < 1 || !IsFinite(domainMinimum) || !IsFinite(domainMaximum)) return false;
        var stack = new List<QuadraticValue> { QuadraticValue.Variable };
        if (!TryExecuteQuadratic(_instructions, stack) || stack.Count != outputCount) return false;
        double minimum = Math.Min(domainMinimum, domainMaximum);
        double maximum = Math.Max(domainMinimum, domainMaximum);
        for (int index = 0; index < stack.Count; index++) {
            QuadraticValue value = stack[index];
            double derivativeAtMinimum = value.C1 + (2D * value.C2 * minimum);
            double derivativeAtMaximum = value.C1 + (2D * value.C2 * maximum);
            if (!IsFinite(derivativeAtMinimum) || !IsFinite(derivativeAtMaximum) ||
                (Math.Min(derivativeAtMinimum, derivativeAtMaximum) < 0D &&
                 Math.Max(derivativeAtMinimum, derivativeAtMaximum) > 0D)) return false;
        }
        return true;
    }

    private static bool TryExecuteQuadratic(Instruction[] instructions, List<QuadraticValue> stack) {
        for (int index = 0; index < instructions.Length; index++) {
            Instruction instruction = instructions[index];
            switch (instruction.Kind) {
                case InstructionKind.Integer:
                    if (!TryPushQuadratic(stack, QuadraticValue.Constant(instruction.IntegerValue, instruction.IntegerValue))) return false;
                    break;
                case InstructionKind.Real:
                    if (!TryPushQuadratic(stack, QuadraticValue.Constant(instruction.RealValue, null))) return false;
                    break;
                case InstructionKind.Operator:
                    if (!TryApplyQuadraticOperator(instruction.Operator, stack)) return false;
                    break;
                default:
                    return false;
            }
        }
        return true;
    }

    private static bool TryApplyQuadraticOperator(CalculatorOperator operation, List<QuadraticValue> stack) {
        switch (operation) {
            case CalculatorOperator.Dup:
                return stack.Count > 0 && TryPushQuadratic(stack, stack[stack.Count - 1]);
            case CalculatorOperator.Exch:
                if (stack.Count < 2) return false;
                (stack[stack.Count - 2], stack[stack.Count - 1]) = (stack[stack.Count - 1], stack[stack.Count - 2]);
                return true;
            case CalculatorOperator.Pop:
                return TryPopQuadratic(stack, out _);
            case CalculatorOperator.Copy:
                if (!TryPopQuadraticInteger(stack, out int copyCount) || copyCount < 0 || copyCount > stack.Count || stack.Count > MaxStackValues - copyCount) return false;
                stack.AddRange(stack.Skip(stack.Count - copyCount).Take(copyCount).ToArray());
                return true;
            case CalculatorOperator.Index:
                if (!TryPopQuadraticInteger(stack, out int depth) || depth < 0 || depth >= stack.Count) return false;
                return TryPushQuadratic(stack, stack[stack.Count - depth - 1]);
            case CalculatorOperator.Roll:
                if (!TryPopQuadraticInteger(stack, out int shift) || !TryPopQuadraticInteger(stack, out int count) || count < 0 || count > stack.Count) return false;
                if (count < 2) return true;
                int normalized = shift % count;
                if (normalized < 0) normalized += count;
                if (normalized == 0) return true;
                int start = stack.Count - count;
                QuadraticValue[] values = stack.Skip(start).Take(count).ToArray();
                for (int valueIndex = 0; valueIndex < count; valueIndex++) stack[start + ((valueIndex + normalized) % count)] = values[valueIndex];
                return true;
            case CalculatorOperator.Neg:
                return TryPopQuadratic(stack, out QuadraticValue negated) &&
                    TryPushQuadratic(stack, new QuadraticValue(-negated.C0, -negated.C1, -negated.C2, null));
            case CalculatorOperator.Cvr:
                return stack.Count > 0;
            case CalculatorOperator.Add:
            case CalculatorOperator.Sub:
            case CalculatorOperator.Mul:
            case CalculatorOperator.Div:
                return TryApplyQuadraticArithmetic(operation, stack);
            default:
                return false;
        }
    }

    private static bool TryApplyQuadraticArithmetic(CalculatorOperator operation, List<QuadraticValue> stack) {
        if (!TryPopQuadratic(stack, out QuadraticValue right) || !TryPopQuadratic(stack, out QuadraticValue left)) return false;
        QuadraticValue result;
        switch (operation) {
            case CalculatorOperator.Add:
                result = new QuadraticValue(left.C0 + right.C0, left.C1 + right.C1, left.C2 + right.C2, null);
                break;
            case CalculatorOperator.Sub:
                result = new QuadraticValue(left.C0 - right.C0, left.C1 - right.C1, left.C2 - right.C2, null);
                break;
            case CalculatorOperator.Mul:
                if ((left.C2 != 0D && (right.C1 != 0D || right.C2 != 0D)) ||
                    (right.C2 != 0D && (left.C1 != 0D || left.C2 != 0D)) ||
                    (left.C1 != 0D && right.C2 != 0D) || (right.C1 != 0D && left.C2 != 0D)) return false;
                result = new QuadraticValue(
                    left.C0 * right.C0,
                    (left.C0 * right.C1) + (left.C1 * right.C0),
                    (left.C0 * right.C2) + (left.C1 * right.C1) + (left.C2 * right.C0),
                    null);
                break;
            default:
                if (!right.IsConstant || right.C0 == 0D) return false;
                result = new QuadraticValue(left.C0 / right.C0, left.C1 / right.C0, left.C2 / right.C0, null);
                break;
        }
        return result.IsFinite && TryPushQuadratic(stack, result);
    }

    private static bool TryPushQuadratic(List<QuadraticValue> stack, QuadraticValue value) {
        if (!value.IsFinite || stack.Count >= MaxStackValues) return false;
        stack.Add(value);
        return true;
    }

    private static bool TryPopQuadratic(List<QuadraticValue> stack, out QuadraticValue value) {
        if (stack.Count == 0) { value = default; return false; }
        int index = stack.Count - 1;
        value = stack[index];
        stack.RemoveAt(index);
        return true;
    }

    private static bool TryPopQuadraticInteger(List<QuadraticValue> stack, out int value) {
        value = 0;
        if (!TryPopQuadratic(stack, out QuadraticValue item) || item.IntegerConstant is not int integer) return false;
        value = integer;
        return true;
    }

    private readonly struct QuadraticValue {
        internal QuadraticValue(double c0, double c1, double c2, int? integerConstant) {
            C0 = c0;
            C1 = c1;
            C2 = c2;
            IntegerConstant = integerConstant;
        }
        internal double C0 { get; }
        internal double C1 { get; }
        internal double C2 { get; }
        internal int? IntegerConstant { get; }
        internal bool IsConstant => C1 == 0D && C2 == 0D;
        internal bool IsFinite => PdfCalculatorProgram.IsFinite(C0) && PdfCalculatorProgram.IsFinite(C1) && PdfCalculatorProgram.IsFinite(C2);
        internal static QuadraticValue Variable => new QuadraticValue(0D, 1D, 0D, null);
        internal static QuadraticValue Constant(double value, int? integerConstant) => new QuadraticValue(value, 0D, 0D, integerConstant);
    }
}
