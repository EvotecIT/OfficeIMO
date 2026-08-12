using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>Parsed, bounded evaluator for the calculator language defined by PDF Type 4 functions.</summary>
internal sealed partial class PdfCalculatorProgram {
    internal const int MaxProgramBytes = 64 * 1024;
    internal const long MaxValidationWork = 4L * 1024L * 1024L;
    private const int MaxInstructions = 4096;
    private const int MaxProcedureDepth = 16;
    private const int MaxStackValues = 256;

    private readonly Instruction[] _instructions;
    private readonly double[] _numericConstants;
    [ThreadStatic]
    private static CalculatorStack? _threadStack;

    private PdfCalculatorProgram(
        Instruction[] instructions,
        double[] numericConstants,
        int instructionCount,
        int maximumEvaluationWork,
        int sourceLength) {
        _instructions = instructions;
        _numericConstants = numericConstants;
        InstructionCount = instructionCount;
        MaximumEvaluationWork = maximumEvaluationWork;
        RetainedBytes = checked(sourceLength + instructionCount * 64L);
    }

    internal int InstructionCount { get; }

    internal int MaximumEvaluationWork { get; }

    internal long RetainedBytes { get; }

    internal IReadOnlyList<double> NumericConstants => _numericConstants;

    internal static bool TryParse(byte[] source, out PdfCalculatorProgram program) {
        program = null!;
        if (source == null || source.Length > MaxProgramBytes) return false;
        var parser = new Parser(source);
        return parser.TryParse(out program);
    }

    internal double[]? Evaluate(double[] inputs, int outputCount) {
        if (outputCount < 1 || outputCount > MaxStackValues) return null;
        var result = new double[outputCount];
        return TryEvaluate(inputs, result, 0, outputCount) ? result : null;
    }

    internal bool TryEvaluate(double[] inputs, double[] output, int outputOffset, int outputCount) {
        if (inputs == null || inputs.Length > MaxStackValues || output == null ||
            outputCount < 1 || outputCount > MaxStackValues ||
            outputOffset < 0 || outputOffset > output.Length - outputCount) return false;
        CalculatorStack stack = AcquireStack();
        try {
            for (int index = 0; index < inputs.Length; index++) {
                if (!IsFinite(inputs[index]) || !stack.TryPush(Value.Real(inputs[index]))) return false;
            }

            int remainingSteps = MaxInstructions;
            if (!TryExecute(_instructions, stack, depth: 0, ref remainingSteps) || stack.Count != outputCount) return false;

            for (int index = 0; index < outputCount; index++) {
                Value value = stack[index];
                if (!value.TryGetNumber(out double number) || !IsFinite(number)) return false;
                output[outputOffset + index] = number;
            }
            return true;
        } finally {
            stack.Release();
        }
    }

    internal bool CanEvaluateDomain(
        double[] domain,
        int inputCount,
        int outputCount,
        ref long remainingValidationWork) {
        if (domain == null || domain.Length != inputCount * 2 || inputCount < 1) return false;
        var midpoint = new double[inputCount];
        var minimum = new double[inputCount];
        var maximum = new double[inputCount];
        for (int index = 0; index < inputCount; index++) {
            minimum[index] = domain[index * 2];
            maximum[index] = domain[index * 2 + 1];
            midpoint[index] = minimum[index] * 0.5D + maximum[index] * 0.5D;
        }
        if (!TryEvaluateForValidation(midpoint, outputCount, ref remainingValidationWork) ||
            !TryEvaluateForValidation(minimum, outputCount, ref remainingValidationWork) ||
            !TryEvaluateForValidation(maximum, outputCount, ref remainingValidationWork)) return false;

        if (inputCount == 1) {
            double domainStart = minimum[0];
            double domainEnd = maximum[0];
            double delta = Math.Max(Math.Abs(domainEnd - domainStart) * 1E-9D, 1E-12D);
            foreach (double authoredValue in _numericConstants) {
                if (authoredValue < domainStart || authoredValue > domainEnd) continue;
                double[] candidates = {
                    authoredValue,
                    Math.Max(domainStart, authoredValue - delta),
                    Math.Min(domainEnd, authoredValue + delta)
                };
                for (int index = 0; index < candidates.Length; index++) {
                    if (!TryEvaluateForValidation(candidates[index], outputCount, ref remainingValidationWork)) return false;
                }
            }
            return true;
        }

        for (int input = 0; input < inputCount; input++) {
            double[] candidate = (double[])midpoint.Clone();
            candidate[input] = minimum[input];
            if (!TryEvaluateForValidation(candidate, outputCount, ref remainingValidationWork)) return false;
            candidate[input] = maximum[input];
            if (!TryEvaluateForValidation(candidate, outputCount, ref remainingValidationWork)) return false;
        }
        return true;
    }

    private bool TryEvaluateForValidation(double input, int outputCount, ref long remainingValidationWork) =>
        TryEvaluateForValidation(new[] { input }, outputCount, ref remainingValidationWork);

    private bool TryEvaluateForValidation(double[] inputs, int outputCount, ref long remainingValidationWork) {
        if (MaximumEvaluationWork > remainingValidationWork) return false;
        remainingValidationWork -= MaximumEvaluationWork;
        return Evaluate(inputs, outputCount) != null;
    }

    private static bool TryExecute(
        Instruction[] instructions,
        CalculatorStack stack,
        int depth,
        ref int remainingSteps) {
        if (depth > MaxProcedureDepth) return false;
        for (int index = 0; index < instructions.Length; index++) {
            if (--remainingSteps < 0) return false;
            Instruction instruction = instructions[index];
            switch (instruction.Kind) {
                case InstructionKind.Integer:
                    if (!stack.TryPush(Value.Integer(instruction.IntegerValue))) return false;
                    break;
                case InstructionKind.Real:
                    if (!stack.TryPush(Value.Real(instruction.RealValue))) return false;
                    break;
                case InstructionKind.Boolean:
                    if (!stack.TryPush(Value.Boolean(instruction.BooleanValue))) return false;
                    break;
                case InstructionKind.Operator:
                    if (!TryExecuteOperator(instruction.Operator, stack)) return false;
                    break;
                case InstructionKind.Conditional:
                    if (!stack.TryPopBoolean(out bool condition)) return false;
                    Instruction[]? branch = condition ? instruction.TrueBranch : instruction.FalseBranch;
                    if (branch != null && !TryExecute(branch, stack, depth + 1, ref remainingSteps)) return false;
                    break;
                default:
                    return false;
            }
        }
        return true;
    }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private static CalculatorStack AcquireStack() {
        CalculatorStack? stack = _threadStack;
        if (stack == null) {
            stack = new CalculatorStack();
            _threadStack = stack;
        } else if (stack.IsInUse) {
            stack = new CalculatorStack();
        }
        stack.Acquire();
        return stack;
    }

    private sealed class Parser {
        private readonly byte[] _source;
        private readonly List<double> _numericConstants = new List<double>();
        private int _position;
        private int _instructionCount;

        internal Parser(byte[] source) {
            _source = source;
        }

        internal bool TryParse(out PdfCalculatorProgram program) {
            program = null!;
            if (!TryReadToken(out Token open) || open.Kind != TokenKind.OpenBrace ||
                !TryParseBlock(depth: 0, out Instruction[] instructions, out int maximumSteps) ||
                !TryReadToken(out Token end) || end.Kind != TokenKind.End) return false;

            program = new PdfCalculatorProgram(
                instructions,
                _numericConstants.Distinct().OrderBy(static value => value).ToArray(),
                _instructionCount,
                maximumSteps,
                _source.Length);
            return true;
        }

        private bool TryParseBlock(int depth, out Instruction[] instructions, out int maximumSteps) {
            instructions = Array.Empty<Instruction>();
            maximumSteps = 0;
            if (depth > MaxProcedureDepth) return false;
            var result = new List<Instruction>();
            while (TryReadToken(out Token token)) {
                if (token.Kind == TokenKind.CloseBrace) {
                    instructions = result.ToArray();
                    return true;
                }
                if (token.Kind is TokenKind.End) return false;

                Instruction instruction;
                int instructionSteps = 1;
                if (token.Kind == TokenKind.OpenBrace) {
                    if (!TryParseBlock(depth + 1, out Instruction[] trueBranch, out int trueSteps) ||
                        !TryReadToken(out Token conditionalToken)) return false;
                    Instruction[]? falseBranch = null;
                    int falseSteps = 0;
                    if (conditionalToken.Kind == TokenKind.OpenBrace) {
                        if (!TryParseBlock(depth + 1, out falseBranch, out falseSteps) ||
                            !TryReadToken(out conditionalToken)) return false;
                    }
                    if (conditionalToken.Kind != TokenKind.Word ||
                        (falseBranch == null && !string.Equals(conditionalToken.Text, "if", StringComparison.Ordinal)) ||
                        (falseBranch != null && !string.Equals(conditionalToken.Text, "ifelse", StringComparison.Ordinal))) return false;
                    instruction = Instruction.Conditional(trueBranch, falseBranch);
                    instructionSteps = checked(1 + Math.Max(trueSteps, falseSteps));
                } else if (token.Kind != TokenKind.Word || !TryCreateInstruction(token.Text!, out instruction)) {
                    return false;
                } else {
                    instructionSteps = GetEvaluationWork(instruction);
                }

                if (++_instructionCount > MaxInstructions) return false;
                maximumSteps = checked(maximumSteps + instructionSteps);
                result.Add(instruction);
            }
            return false;
        }

        private static int GetEvaluationWork(Instruction instruction) {
            if (instruction.Kind != InstructionKind.Operator) return 1;
            return instruction.Operator switch {
                CalculatorOperator.Copy => MaxStackValues,
                CalculatorOperator.Roll => MaxStackValues * 3,
                _ => 1
            };
        }

        private bool TryCreateInstruction(string token, out Instruction instruction) {
            instruction = default;
            if (int.TryParse(token, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out int integer)) {
                instruction = Instruction.Integer(integer);
                _numericConstants.Add(integer);
                return true;
            }
            if (double.TryParse(
                    token,
                    NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint,
                    CultureInfo.InvariantCulture,
                    out double real) && IsFinite(real)) {
                instruction = Instruction.Real(real);
                _numericConstants.Add(real);
                return true;
            }
            if (string.Equals(token, "true", StringComparison.Ordinal)) {
                instruction = Instruction.Boolean(true);
                return true;
            }
            if (string.Equals(token, "false", StringComparison.Ordinal)) {
                instruction = Instruction.Boolean(false);
                return true;
            }
            if (string.Equals(token, "if", StringComparison.Ordinal) ||
                string.Equals(token, "ifelse", StringComparison.Ordinal) ||
                !TryReadOperator(token, out CalculatorOperator calculatorOperator)) return false;
            instruction = Instruction.Operation(calculatorOperator);
            return true;
        }

        private bool TryReadToken(out Token token) {
            SkipWhitespaceAndComments();
            if (_position >= _source.Length) {
                token = new Token(TokenKind.End, null);
                return true;
            }

            byte current = _source[_position++];
            if (current == (byte)'{') {
                token = new Token(TokenKind.OpenBrace, null);
                return true;
            }
            if (current == (byte)'}') {
                token = new Token(TokenKind.CloseBrace, null);
                return true;
            }

            int start = _position - 1;
            while (_position < _source.Length) {
                byte value = _source[_position];
                if (IsWhitespace(value) || value is (byte)'{' or (byte)'}' or (byte)'%') break;
                _position++;
            }
            if (_position == start) {
                token = default;
                return false;
            }
            for (int index = start; index < _position; index++) {
                if (_source[index] > 0x7F) {
                    token = default;
                    return false;
                }
            }
            token = new Token(TokenKind.Word, System.Text.Encoding.ASCII.GetString(_source, start, _position - start));
            return true;
        }

        private void SkipWhitespaceAndComments() {
            while (_position < _source.Length) {
                byte value = _source[_position];
                if (IsWhitespace(value)) {
                    _position++;
                    continue;
                }
                if (value != (byte)'%') return;
                _position++;
                while (_position < _source.Length && _source[_position] is not (byte)'\r' and not (byte)'\n') _position++;
            }
        }

        private static bool IsWhitespace(byte value) => value is 0 or 9 or 10 or 12 or 13 or 32;
    }

    private sealed class CalculatorStack {
        private readonly Value[] _values = new Value[MaxStackValues];

        internal int Count { get; private set; }

        internal bool IsInUse { get; private set; }

        internal Value this[int index] => _values[index];

        internal void Acquire() {
            Count = 0;
            IsInUse = true;
        }

        internal void Release() {
            Count = 0;
            IsInUse = false;
        }

        internal bool TryPush(Value value) {
            if (Count >= _values.Length) return false;
            _values[Count++] = value;
            return true;
        }

        internal bool TryPop(out Value value) {
            if (Count < 1) {
                value = default;
                return false;
            }
            value = _values[--Count];
            return true;
        }

        internal bool TryPopInteger(out int value) {
            value = 0;
            if (!TryPop(out Value item) || item.Kind != ValueKind.Integer) return false;
            value = item.IntegerValue;
            return true;
        }

        internal bool TryPopBoolean(out bool value) {
            value = false;
            if (!TryPop(out Value item) || item.Kind != ValueKind.Boolean) return false;
            value = item.BooleanValue;
            return true;
        }

        internal bool TryPeek(int depth, out Value value) {
            int index = Count - depth - 1;
            if (index < 0) {
                value = default;
                return false;
            }
            value = _values[index];
            return true;
        }

        internal bool TryCopy(int count) {
            if (count < 0 || count > Count || Count > _values.Length - count) return false;
            Array.Copy(_values, Count - count, _values, Count, count);
            Count += count;
            return true;
        }

        internal bool TryExchange() {
            if (Count < 2) return false;
            Value value = _values[Count - 1];
            _values[Count - 1] = _values[Count - 2];
            _values[Count - 2] = value;
            return true;
        }

        internal bool TryPopValue() {
            if (Count < 1) return false;
            Count--;
            return true;
        }

        internal bool TryRoll(int count, int shift) {
            if (count < 0 || count > Count) return false;
            if (count < 2) return true;
            int normalized = shift % count;
            if (normalized < 0) normalized += count;
            if (normalized == 0) return true;
            int start = Count - count;
            Array.Reverse(_values, start, count);
            Array.Reverse(_values, start, normalized);
            Array.Reverse(_values, start + normalized, count - normalized);
            return true;
        }
    }

    private readonly struct Value {
        private Value(ValueKind kind, int integerValue, double realValue, bool booleanValue) {
            Kind = kind;
            IntegerValue = integerValue;
            RealValue = realValue;
            BooleanValue = booleanValue;
        }

        internal ValueKind Kind { get; }
        internal int IntegerValue { get; }
        internal double RealValue { get; }
        internal bool BooleanValue { get; }

        internal static Value Integer(int value) => new Value(ValueKind.Integer, value, value, false);
        internal static Value Real(double value) => new Value(ValueKind.Real, 0, value, false);
        internal static Value Boolean(bool value) => new Value(ValueKind.Boolean, 0, 0D, value);

        internal bool TryGetNumber(out double value) {
            value = Kind == ValueKind.Integer ? IntegerValue : RealValue;
            return Kind is ValueKind.Integer or ValueKind.Real;
        }
    }

    private readonly struct Instruction {
        private Instruction(
            InstructionKind kind,
            int integerValue,
            double realValue,
            bool booleanValue,
            CalculatorOperator calculatorOperator,
            Instruction[]? trueBranch,
            Instruction[]? falseBranch) {
            Kind = kind;
            IntegerValue = integerValue;
            RealValue = realValue;
            BooleanValue = booleanValue;
            Operator = calculatorOperator;
            TrueBranch = trueBranch;
            FalseBranch = falseBranch;
        }

        internal InstructionKind Kind { get; }
        internal int IntegerValue { get; }
        internal double RealValue { get; }
        internal bool BooleanValue { get; }
        internal CalculatorOperator Operator { get; }
        internal Instruction[]? TrueBranch { get; }
        internal Instruction[]? FalseBranch { get; }

        internal static Instruction Integer(int value) => new Instruction(InstructionKind.Integer, value, 0D, false, default, null, null);
        internal static Instruction Real(double value) => new Instruction(InstructionKind.Real, 0, value, false, default, null, null);
        internal static Instruction Boolean(bool value) => new Instruction(InstructionKind.Boolean, 0, 0D, value, default, null, null);
        internal static Instruction Operation(CalculatorOperator value) => new Instruction(InstructionKind.Operator, 0, 0D, false, value, null, null);
        internal static Instruction Conditional(Instruction[] trueBranch, Instruction[]? falseBranch) =>
            new Instruction(InstructionKind.Conditional, 0, 0D, false, default, trueBranch, falseBranch);
    }

    private readonly struct Token {
        internal Token(TokenKind kind, string? text) {
            Kind = kind;
            Text = text;
        }
        internal TokenKind Kind { get; }
        internal string? Text { get; }
    }

    private enum TokenKind { End, OpenBrace, CloseBrace, Word }
    private enum InstructionKind { Integer, Real, Boolean, Operator, Conditional }
    private enum ValueKind { Integer, Real, Boolean }

    private enum CalculatorOperator {
        Abs, Add, And, Atan, Bitshift, Ceiling, Copy, Cos, Cvi, Cvr, Div, Dup, Eq, Exch, Exp,
        Floor, Ge, Gt, Idiv, Index, Le, Ln, Log, Lt, Mod, Mul, Ne, Neg, Not, Or, Pop, Roll,
        Round, Sin, Sqrt, Sub, Truncate, Xor
    }
}
