namespace OfficeIMO.Pdf;

internal sealed partial class PdfCalculatorProgram {
    private const int MaxConditionalAnalysisPaths = 256;

    internal bool TryGetOneInputConditionalBoundaries(
        double domainMinimum,
        double domainMaximum,
        out double[] boundaries) {
        boundaries = Array.Empty<double>();
        if (!HasConditional || !IsFinite(domainMinimum) || !IsFinite(domainMaximum)) return !HasConditional;

        var discovered = new HashSet<double>();
        var initial = new SymbolicStack();
        initial.Values.Add(SymbolicValue.Affine(1D, 0D));
        if (!TryAnalyzeInstructions(_instructions, new List<SymbolicStack> { initial }, discovered, out _)) return false;

        boundaries = discovered
            .Where(value => value >= domainMinimum && value <= domainMaximum)
            .Distinct()
            .OrderBy(static value => value)
            .ToArray();
        return true;
    }

    private static bool TryAnalyzeInstructions(
        Instruction[] instructions,
        List<SymbolicStack> inputStates,
        HashSet<double> boundaries,
        out List<SymbolicStack> outputStates) {
        outputStates = inputStates;
        for (int instructionIndex = 0; instructionIndex < instructions.Length; instructionIndex++) {
            Instruction instruction = instructions[instructionIndex];
            var nextStates = new List<SymbolicStack>();
            for (int stateIndex = 0; stateIndex < outputStates.Count; stateIndex++) {
                SymbolicStack state = outputStates[stateIndex];
                if (instruction.Kind != InstructionKind.Conditional) {
                    if (!TryApplySymbolicInstruction(instruction, state)) return false;
                    nextStates.Add(state);
                    continue;
                }

                if (!state.TryPop(out SymbolicValue condition) || condition.Kind != SymbolicValueKind.Boolean || !condition.IsKnown) {
                    return false;
                }
                foreach (double boundary in condition.Boundaries) {
                    if (!IsFinite(boundary)) return false;
                    boundaries.Add(boundary);
                }

                if (condition.HasConstantValue) {
                    Instruction[]? selected = condition.ConstantValue ? instruction.TrueBranch : instruction.FalseBranch;
                    if (selected == null) {
                        nextStates.Add(state);
                    } else if (!TryAnalyzeInstructions(selected, new List<SymbolicStack> { state }, boundaries, out List<SymbolicStack> selectedStates)) {
                        return false;
                    } else {
                        nextStates.AddRange(selectedStates);
                    }
                    continue;
                }

                var trueState = state.Clone();
                if (!TryAnalyzeInstructions(instruction.TrueBranch!, new List<SymbolicStack> { trueState }, boundaries, out List<SymbolicStack> trueStates)) {
                    return false;
                }
                nextStates.AddRange(trueStates);

                if (instruction.FalseBranch == null) {
                    nextStates.Add(state);
                } else if (!TryAnalyzeInstructions(instruction.FalseBranch, new List<SymbolicStack> { state }, boundaries, out List<SymbolicStack> falseStates)) {
                    return false;
                } else {
                    nextStates.AddRange(falseStates);
                }
            }
            if (nextStates.Count > MaxConditionalAnalysisPaths) return false;
            outputStates = nextStates;
        }
        return true;
    }

    private static bool TryApplySymbolicInstruction(Instruction instruction, SymbolicStack stack) {
        switch (instruction.Kind) {
            case InstructionKind.Integer:
                return stack.TryPush(SymbolicValue.Integer(instruction.IntegerValue));
            case InstructionKind.Real:
                return stack.TryPush(SymbolicValue.Affine(0D, instruction.RealValue));
            case InstructionKind.Boolean:
                return stack.TryPush(SymbolicValue.Boolean(instruction.BooleanValue));
            case InstructionKind.Operator:
                return TryApplySymbolicOperator(instruction.Operator, stack);
            default:
                return false;
        }
    }

    private static bool TryApplySymbolicOperator(CalculatorOperator operation, SymbolicStack stack) {
        switch (operation) {
            case CalculatorOperator.Dup:
                return stack.TryPeek(0, out SymbolicValue duplicate) && stack.TryPush(duplicate);
            case CalculatorOperator.Exch:
                return stack.TryExchange();
            case CalculatorOperator.Pop:
                return stack.TryPop(out _);
            case CalculatorOperator.Copy:
                return stack.TryPopInteger(out int copyCount) && stack.TryCopy(copyCount);
            case CalculatorOperator.Index:
                return stack.TryPopInteger(out int depth) && depth >= 0 && stack.TryPeek(depth, out SymbolicValue indexed) && stack.TryPush(indexed);
            case CalculatorOperator.Roll:
                return stack.TryPopInteger(out int shift) && stack.TryPopInteger(out int rollCount) && stack.TryRoll(rollCount, shift);
            case CalculatorOperator.Neg:
                return stack.TryPop(out SymbolicValue negated) && stack.TryPush(negated.TryNegate());
            case CalculatorOperator.Add:
            case CalculatorOperator.Sub:
            case CalculatorOperator.Mul:
            case CalculatorOperator.Div:
                return TryApplyAffineArithmetic(operation, stack);
            case CalculatorOperator.Eq:
            case CalculatorOperator.Ne:
            case CalculatorOperator.Ge:
            case CalculatorOperator.Gt:
            case CalculatorOperator.Le:
            case CalculatorOperator.Lt:
                return TryApplyComparison(operation, stack);
            case CalculatorOperator.Not:
                return stack.TryPop(out SymbolicValue value) && value.TryNot(out SymbolicValue inverted) && stack.TryPush(inverted);
            case CalculatorOperator.And:
            case CalculatorOperator.Or:
            case CalculatorOperator.Xor:
                return TryApplyBoolean(operation, stack);
            case CalculatorOperator.Abs:
            case CalculatorOperator.Atan:
            case CalculatorOperator.Bitshift:
            case CalculatorOperator.Ceiling:
            case CalculatorOperator.Cos:
            case CalculatorOperator.Cvi:
            case CalculatorOperator.Cvr:
            case CalculatorOperator.Exp:
            case CalculatorOperator.Floor:
            case CalculatorOperator.Idiv:
            case CalculatorOperator.Ln:
            case CalculatorOperator.Log:
            case CalculatorOperator.Mod:
            case CalculatorOperator.Round:
            case CalculatorOperator.Sin:
            case CalculatorOperator.Sqrt:
            case CalculatorOperator.Truncate:
                return TryApplyOpaqueNumeric(operation, stack);
            default:
                return false;
        }
    }

    private static bool TryApplyAffineArithmetic(CalculatorOperator operation, SymbolicStack stack) {
        if (!stack.TryPop(out SymbolicValue right) || !stack.TryPop(out SymbolicValue left) ||
            left.Kind != SymbolicValueKind.Numeric || right.Kind != SymbolicValueKind.Numeric) return false;
        SymbolicValue result = SymbolicValue.UnknownNumeric;
        if (left.IsAffine && right.IsAffine) {
            switch (operation) {
                case CalculatorOperator.Add:
                    result = SymbolicValue.Affine(left.A + right.A, left.B + right.B);
                    break;
                case CalculatorOperator.Sub:
                    result = SymbolicValue.Affine(left.A - right.A, left.B - right.B);
                    break;
                case CalculatorOperator.Mul when left.A == 0D:
                    result = SymbolicValue.Affine(right.A * left.B, right.B * left.B);
                    break;
                case CalculatorOperator.Mul when right.A == 0D:
                    result = SymbolicValue.Affine(left.A * right.B, left.B * right.B);
                    break;
                case CalculatorOperator.Div when right.A == 0D && right.B != 0D:
                    result = SymbolicValue.Affine(left.A / right.B, left.B / right.B);
                    break;
            }
        }
        return result.IsFinite && stack.TryPush(result);
    }

    private static bool TryApplyComparison(CalculatorOperator operation, SymbolicStack stack) {
        if (!stack.TryPop(out SymbolicValue right) || !stack.TryPop(out SymbolicValue left)) return false;
        if (left.Kind == SymbolicValueKind.Boolean && right.Kind == SymbolicValueKind.Boolean &&
            operation is CalculatorOperator.Eq or CalculatorOperator.Ne &&
            left.HasConstantValue && right.HasConstantValue) {
            bool equal = left.ConstantValue == right.ConstantValue;
            return stack.TryPush(SymbolicValue.Boolean(operation == CalculatorOperator.Eq ? equal : !equal));
        }
        if (left.Kind != right.Kind && operation is CalculatorOperator.Eq or CalculatorOperator.Ne) {
            return stack.TryPush(SymbolicValue.Boolean(operation == CalculatorOperator.Ne));
        }
        if (left.Kind != SymbolicValueKind.Numeric || right.Kind != SymbolicValueKind.Numeric || !left.IsAffine || !right.IsAffine) {
            return stack.TryPush(SymbolicValue.UnknownBoolean);
        }

        double coefficient = left.A - right.A;
        double constant = left.B - right.B;
        if (coefficient == 0D) {
            bool value = operation switch {
                CalculatorOperator.Eq => constant == 0D,
                CalculatorOperator.Ne => constant != 0D,
                CalculatorOperator.Ge => constant >= 0D,
                CalculatorOperator.Gt => constant > 0D,
                CalculatorOperator.Le => constant <= 0D,
                _ => constant < 0D
            };
            return stack.TryPush(SymbolicValue.Boolean(value));
        }

        double boundary = -constant / coefficient;
        return IsFinite(boundary) && stack.TryPush(SymbolicValue.BooleanBoundary(boundary));
    }

    private static bool TryApplyBoolean(CalculatorOperator operation, SymbolicStack stack) {
        if (!stack.TryPop(out SymbolicValue right) || !stack.TryPop(out SymbolicValue left) ||
            left.Kind != SymbolicValueKind.Boolean || right.Kind != SymbolicValueKind.Boolean ||
            !left.IsKnown || !right.IsKnown) return false;
        if (left.HasConstantValue && right.HasConstantValue) {
            bool result = operation switch {
                CalculatorOperator.And => left.ConstantValue && right.ConstantValue,
                CalculatorOperator.Or => left.ConstantValue || right.ConstantValue,
                _ => left.ConstantValue ^ right.ConstantValue
            };
            return stack.TryPush(SymbolicValue.Boolean(result));
        }
        return stack.TryPush(SymbolicValue.BooleanBoundaries(left.Boundaries.Concat(right.Boundaries)));
    }

    private static bool TryApplyOpaqueNumeric(CalculatorOperator operation, SymbolicStack stack) {
        int operands = operation is CalculatorOperator.Atan or CalculatorOperator.Bitshift or CalculatorOperator.Exp or
            CalculatorOperator.Idiv or CalculatorOperator.Mod ? 2 : 1;
        for (int index = 0; index < operands; index++) {
            if (!stack.TryPop(out SymbolicValue value) || value.Kind != SymbolicValueKind.Numeric) return false;
        }
        return stack.TryPush(SymbolicValue.UnknownNumeric);
    }

    private enum SymbolicValueKind { Numeric, Boolean }

    private readonly struct SymbolicValue {
        private SymbolicValue(
            SymbolicValueKind kind,
            bool isAffine,
            double a,
            double b,
            int? integerConstant,
            bool isKnown,
            bool? constantValue,
            double[]? boundaries) {
            Kind = kind;
            IsAffine = isAffine;
            A = a;
            B = b;
            IntegerConstant = integerConstant;
            IsKnown = isKnown;
            ConstantValueOrNull = constantValue;
            Boundaries = boundaries ?? Array.Empty<double>();
        }

        internal SymbolicValueKind Kind { get; }
        internal bool IsAffine { get; }
        internal double A { get; }
        internal double B { get; }
        internal int? IntegerConstant { get; }
        internal bool IsKnown { get; }
        internal bool? ConstantValueOrNull { get; }
        internal bool HasConstantValue => ConstantValueOrNull.HasValue;
        internal bool ConstantValue => ConstantValueOrNull.GetValueOrDefault();
        internal IReadOnlyList<double> Boundaries { get; }
        internal bool IsFinite => !IsAffine || (PdfCalculatorProgram.IsFinite(A) && PdfCalculatorProgram.IsFinite(B));

        internal static SymbolicValue UnknownNumeric => new SymbolicValue(SymbolicValueKind.Numeric, false, 0D, 0D, null, false, null, null);
        internal static SymbolicValue UnknownBoolean => new SymbolicValue(SymbolicValueKind.Boolean, false, 0D, 0D, null, false, null, null);
        internal static SymbolicValue Integer(int value) => new SymbolicValue(SymbolicValueKind.Numeric, true, 0D, value, value, true, null, null);
        internal static SymbolicValue Affine(double a, double b) => new SymbolicValue(SymbolicValueKind.Numeric, true, a, b, null, true, null, null);
        internal static SymbolicValue Boolean(bool value) => new SymbolicValue(SymbolicValueKind.Boolean, false, 0D, 0D, null, true, value, null);
        internal static SymbolicValue BooleanBoundary(double boundary) => BooleanBoundaries(new[] { boundary });
        internal static SymbolicValue BooleanBoundaries(IEnumerable<double> values) =>
            new SymbolicValue(SymbolicValueKind.Boolean, false, 0D, 0D, null, true, null, values.Distinct().ToArray());

        internal SymbolicValue TryNegate() =>
            Kind == SymbolicValueKind.Numeric && IsAffine
                ? Affine(-A, -B)
                : UnknownNumeric;

        internal bool TryNot(out SymbolicValue value) {
            value = UnknownBoolean;
            if (Kind != SymbolicValueKind.Boolean || !IsKnown) return false;
            value = HasConstantValue ? Boolean(!ConstantValue) : BooleanBoundaries(Boundaries);
            return true;
        }
    }

    private sealed class SymbolicStack {
        internal List<SymbolicValue> Values { get; } = new List<SymbolicValue>();

        internal SymbolicStack Clone() {
            var clone = new SymbolicStack();
            clone.Values.AddRange(Values);
            return clone;
        }

        internal bool TryPush(SymbolicValue value) {
            if (Values.Count >= MaxStackValues) return false;
            Values.Add(value);
            return true;
        }

        internal bool TryPop(out SymbolicValue value) {
            if (Values.Count == 0) {
                value = default;
                return false;
            }
            int index = Values.Count - 1;
            value = Values[index];
            Values.RemoveAt(index);
            return true;
        }

        internal bool TryPopInteger(out int value) {
            value = 0;
            if (!TryPop(out SymbolicValue symbolic) || !symbolic.IntegerConstant.HasValue) return false;
            value = symbolic.IntegerConstant.Value;
            return true;
        }

        internal bool TryPeek(int depth, out SymbolicValue value) {
            int index = Values.Count - depth - 1;
            if (index < 0) {
                value = default;
                return false;
            }
            value = Values[index];
            return true;
        }

        internal bool TryExchange() {
            if (Values.Count < 2) return false;
            int last = Values.Count - 1;
            (Values[last - 1], Values[last]) = (Values[last], Values[last - 1]);
            return true;
        }

        internal bool TryCopy(int count) {
            if (count < 0 || count > Values.Count || Values.Count > MaxStackValues - count) return false;
            Values.AddRange(Values.Skip(Values.Count - count).Take(count).ToArray());
            return true;
        }

        internal bool TryRoll(int count, int shift) {
            if (count < 0 || count > Values.Count) return false;
            if (count < 2) return true;
            int normalized = shift % count;
            if (normalized < 0) normalized += count;
            if (normalized == 0) return true;
            int start = Values.Count - count;
            SymbolicValue[] segment = Values.Skip(start).Take(count).ToArray();
            for (int index = 0; index < count; index++) Values[start + ((index + normalized) % count)] = segment[index];
            return true;
        }
    }
}
