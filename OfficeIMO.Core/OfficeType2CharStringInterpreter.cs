using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Drawing;

internal interface IOfficeCffPathSink {
    void MoveTo(double x, double y);
    void LineTo(double x, double y);
    void CurveTo(double control1X, double control1Y, double control2X, double control2Y, double x, double y);
    void CloseContour();
}

/// <summary>Shared operation budget for one externally requested CFF text render.</summary>
internal sealed class OfficeCffOperationBudget {
    private const int DefaultMaximumOperations = 1_000_000;
    private int _remaining;
    private uint _randomState = 0x9E3779B9U;

    internal OfficeCffOperationBudget() : this(DefaultMaximumOperations) {
    }

    internal OfficeCffOperationBudget(int maximumOperations) {
        if (maximumOperations <= 0) throw new ArgumentOutOfRangeException(nameof(maximumOperations));
        _remaining = maximumOperations;
    }

    internal int RemainingOperations => _remaining;

    internal void Consume() {
        if (_remaining-- <= 0) throw new InvalidDataException("The CFF CharString operation budget was exceeded.");
    }

    internal double NextRandom() {
        // Type 2 requires a pseudo-random sequence in (0, 1]. Keep it deterministic for
        // reproducible document output and shared across every glyph in one bounded render.
        uint value = _randomState;
        value ^= value << 13;
        value ^= value >> 17;
        value ^= value << 5;
        _randomState = value;
        return (value + 1D) / (uint.MaxValue + 1D);
    }
}

/// <summary>Bounded Type 2 CharString interpreter for CFF1 and CFF2 glyph outlines.</summary>
internal sealed class OfficeType2CharStringInterpreter {
    private enum ExecutionResult {
        Exhausted,
        Return,
        EndChar
    }

    private const int MaximumStack = 513;
    private const int MaximumSubroutineDepth = 32;
    private readonly OfficeCffFontData _font;
    private readonly OfficeCffFontData.CffIndex _localSubroutines;
    private readonly OfficeCffFontData.CffIndex _globalSubroutines;
    private readonly IOfficeCffPathSink _sink;
    private readonly CancellationToken _cancellationToken;
    private readonly OfficeCffOperationBudget _operationBudget;
    private readonly int _seacDepth;
    private readonly List<double> _stack = new List<double>();
    private readonly double[] _transient = new double[32];
    private double _x;
    private double _y;
    private int _stemCount;
    private int _variationDataIndex;
    private bool _widthConsumed;
    private bool _contourOpen;

    internal OfficeType2CharStringInterpreter(
        OfficeCffFontData font,
        int glyphId,
        IOfficeCffPathSink sink,
        CancellationToken cancellationToken)
        : this(font, glyphId, sink, cancellationToken, new OfficeCffOperationBudget(), 0) {
    }

    internal OfficeType2CharStringInterpreter(
        OfficeCffFontData font,
        int glyphId,
        IOfficeCffPathSink sink,
        CancellationToken cancellationToken,
        OfficeCffOperationBudget operationBudget)
        : this(font, glyphId, sink, cancellationToken, operationBudget, 0) {
    }

    private OfficeType2CharStringInterpreter(
        OfficeCffFontData font,
        int glyphId,
        IOfficeCffPathSink sink,
        CancellationToken cancellationToken,
        OfficeCffOperationBudget operationBudget,
        int seacDepth) {
        _font = font ?? throw new ArgumentNullException(nameof(font));
        _localSubroutines = font.GetLocalSubroutines(glyphId);
        _globalSubroutines = font.GlobalSubroutines;
        _sink = sink ?? throw new ArgumentNullException(nameof(sink));
        _cancellationToken = cancellationToken;
        _operationBudget = operationBudget ?? throw new ArgumentNullException(nameof(operationBudget));
        _seacDepth = seacDepth;
        _widthConsumed = font.IsCff2;
    }

    internal void Render(OfficeCffFontData.CffSlice charString) {
        Execute(charString, 0);
        CloseContour();
        if (_stack.Count != 0) throw new InvalidDataException("The CFF CharString ends with unconsumed operands.");
    }

    private ExecutionResult Execute(OfficeCffFontData.CffSlice program, int depth) {
        if (depth > MaximumSubroutineDepth) throw new InvalidDataException("The CFF CharString subroutine depth is excessive.");
        int offset = program.Offset;
        int end = checked(program.Offset + program.Length);
        byte[] data = program.Data;
        while (offset < end) {
            _cancellationToken.ThrowIfCancellationRequested();
            _operationBudget.Consume();
            int value = data[offset];
            if (value >= 32 || value == 28 || value == 255) {
                Push(OfficeCffFontData.ReadNumber(data, ref offset, end, charString: true));
                continue;
            }
            offset++;
            switch (value) {
                case 1: ConsumeStems(); break;
                case 3: ConsumeStems(); break;
                case 4: MoveVertical(); break;
                case 5: RelativeLines(); break;
                case 6: AlternatingLines(horizontalFirst: true); break;
                case 7: AlternatingLines(horizontalFirst: false); break;
                case 8: RelativeCurves(); break;
                case 10:
                    if (CallSubroutine(_localSubroutines, depth) == ExecutionResult.EndChar) {
                        if (offset != end) throw new InvalidDataException("A CFF CharString contains data after endchar.");
                        return ExecutionResult.EndChar;
                    }
                    break;
                case 11:
                    if (depth == 0) throw new InvalidDataException("A top-level CFF CharString cannot use return.");
                    if (offset != end) throw new InvalidDataException("A CFF CharString subroutine contains data after return.");
                    return ExecutionResult.Return;
                case 12:
                    if (offset >= end) throw new InvalidDataException("A CFF escaped operator is truncated.");
                    ExecuteEscaped(data[offset++]);
                    break;
                case 14:
                    if (offset != end) throw new InvalidDataException("A CFF CharString contains data after endchar.");
                    ConsumeEndChar();
                    return ExecutionResult.EndChar;
                case 15: SelectVariationData(); break;
                case 16: Blend(); break;
                case 18: ConsumeStems(); break;
                case 19:
                case 20:
                    ConsumeStems();
                    int maskBytes = (_stemCount + 7) / 8;
                    if (offset > end - maskBytes) throw new InvalidDataException("A CFF hint mask is truncated.");
                    offset += maskBytes;
                    break;
                case 21: MoveRelative(); break;
                case 22: MoveHorizontal(); break;
                case 23: ConsumeStems(); break;
                case 24: CurveThenLine(); break;
                case 25: LineThenCurve(); break;
                case 26: VerticalCurves(); break;
                case 27: HorizontalCurves(); break;
                case 29:
                    if (CallSubroutine(_globalSubroutines, depth) == ExecutionResult.EndChar) {
                        if (offset != end) throw new InvalidDataException("A CFF CharString contains data after endchar.");
                        return ExecutionResult.EndChar;
                    }
                    break;
                case 30: AlternatingCurves(horizontalFirst: false); break;
                case 31: AlternatingCurves(horizontalFirst: true); break;
                default: throw new NotSupportedException("CFF CharString operator " + value + " is not supported.");
            }
        }
        return ExecutionResult.Exhausted;
    }

    private void ConsumeStems() {
        ConsumeWidthForStem();
        if ((_stack.Count & 1) != 0) throw new InvalidDataException("A CFF stem operator has an odd operand count.");
        _stemCount = checked(_stemCount + _stack.Count / 2);
        _stack.Clear();
    }

    private void ConsumeWidthForStem() {
        if (_widthConsumed) return;
        if ((_stack.Count & 1) != 0) _stack.RemoveAt(0);
        _widthConsumed = true;
    }

    private void ConsumeWidthForMove(int coordinateCount) {
        if (!_widthConsumed && _stack.Count > coordinateCount) _stack.RemoveAt(0);
        _widthConsumed = true;
    }

    private void MoveRelative() {
        ConsumeWidthForMove(2);
        RequireCount(2);
        CloseContour();
        double nextX = _x + _stack[0];
        double nextY = _y + _stack[1];
        EnsureFinitePoint(nextX, nextY);
        _x = nextX;
        _y = nextY;
        _sink.MoveTo(_x, _y);
        _contourOpen = true;
        _stack.Clear();
    }

    private void MoveHorizontal() {
        ConsumeWidthForMove(1);
        RequireCount(1);
        CloseContour();
        double nextX = _x + _stack[0];
        EnsureFinitePoint(nextX, _y);
        _x = nextX;
        _sink.MoveTo(_x, _y);
        _contourOpen = true;
        _stack.Clear();
    }

    private void MoveVertical() {
        ConsumeWidthForMove(1);
        RequireCount(1);
        CloseContour();
        double nextY = _y + _stack[0];
        EnsureFinitePoint(_x, nextY);
        _y = nextY;
        _sink.MoveTo(_x, _y);
        _contourOpen = true;
        _stack.Clear();
    }

    private void RelativeLines() {
        RequireMultiple(2, minimum: 2);
        for (int index = 0; index < _stack.Count; index += 2) LineBy(_stack[index], _stack[index + 1]);
        _stack.Clear();
    }

    private void AlternatingLines(bool horizontalFirst) {
        if (_stack.Count == 0) throw new InvalidDataException("A CFF line operator has no operands.");
        bool horizontal = horizontalFirst;
        for (int index = 0; index < _stack.Count; index++) {
            LineBy(horizontal ? _stack[index] : 0D, horizontal ? 0D : _stack[index]);
            horizontal = !horizontal;
        }
        _stack.Clear();
    }

    private void RelativeCurves() {
        RequireMultiple(6, minimum: 6);
        for (int index = 0; index < _stack.Count; index += 6) CurveBy(index);
        _stack.Clear();
    }

    private void CurveThenLine() {
        if (_stack.Count < 8 || (_stack.Count - 2) % 6 != 0) throw new InvalidDataException("A CFF rcurveline operand sequence is invalid.");
        int index = 0;
        while (index < _stack.Count - 2) {
            CurveBy(index);
            index += 6;
        }
        LineBy(_stack[index], _stack[index + 1]);
        _stack.Clear();
    }

    private void LineThenCurve() {
        if (_stack.Count < 8 || (_stack.Count - 6) % 2 != 0) throw new InvalidDataException("A CFF rlinecurve operand sequence is invalid.");
        int index = 0;
        while (index < _stack.Count - 6) {
            LineBy(_stack[index], _stack[index + 1]);
            index += 2;
        }
        CurveBy(index);
        _stack.Clear();
    }

    private void VerticalCurves() {
        if (_stack.Count < 4) throw new InvalidDataException("A CFF vvcurveto operator has too few operands.");
        int index = 0;
        double firstDx = (_stack.Count & 1) != 0 ? _stack[index++] : 0D;
        while (index < _stack.Count) {
            if (index > _stack.Count - 4) throw new InvalidDataException("A CFF vvcurveto operand sequence is invalid.");
            CurveBy(firstDx, _stack[index], _stack[index + 1], _stack[index + 2], 0D, _stack[index + 3]);
            firstDx = 0D;
            index += 4;
        }
        _stack.Clear();
    }

    private void HorizontalCurves() {
        if (_stack.Count < 4) throw new InvalidDataException("A CFF hhcurveto operator has too few operands.");
        int index = 0;
        double firstDy = (_stack.Count & 1) != 0 ? _stack[index++] : 0D;
        while (index < _stack.Count) {
            if (index > _stack.Count - 4) throw new InvalidDataException("A CFF hhcurveto operand sequence is invalid.");
            CurveBy(_stack[index], firstDy, _stack[index + 1], _stack[index + 2], _stack[index + 3], 0D);
            firstDy = 0D;
            index += 4;
        }
        _stack.Clear();
    }

    private void AlternatingCurves(bool horizontalFirst) {
        if (_stack.Count < 4) throw new InvalidDataException("A CFF alternating curve operator has too few operands.");
        int index = 0;
        bool horizontal = horizontalFirst;
        while (index < _stack.Count) {
            int remaining = _stack.Count - index;
            if (remaining < 4) throw new InvalidDataException("A CFF alternating curve operand sequence is invalid.");
            if (horizontal) {
                double finalDx = remaining == 5 ? _stack[index + 4] : 0D;
                CurveBy(_stack[index], 0D, _stack[index + 1], _stack[index + 2], finalDx, _stack[index + 3]);
            } else {
                double finalDy = remaining == 5 ? _stack[index + 4] : 0D;
                CurveBy(0D, _stack[index], _stack[index + 1], _stack[index + 2], _stack[index + 3], finalDy);
            }
            index += remaining == 5 ? 5 : 4;
            horizontal = !horizontal;
        }
        _stack.Clear();
    }

    private void ExecuteEscaped(int operation) {
        switch (operation) {
            case 3: Binary((left, right) => left != 0D && right != 0D ? 1D : 0D); break;
            case 4: Binary((left, right) => left != 0D || right != 0D ? 1D : 0D); break;
            case 5: Push(Pop() == 0D ? 1D : 0D); break;
            case 9: Push(Math.Abs(Pop())); break;
            case 10: Binary((left, right) => left + right); break;
            case 11: Binary((left, right) => left - right); break;
            case 12: Binary((left, right) => right == 0D ? throw new InvalidDataException("A CFF division uses a zero divisor.") : left / right); break;
            case 14: Push(-Pop()); break;
            case 15: Binary((left, right) => left == right ? 1D : 0D); break;
            case 18: _ = Pop(); break;
            case 20: PutTransient(); break;
            case 21: GetTransient(); break;
            case 22: Conditional(); break;
            case 23: Push(_operationBudget.NextRandom()); break;
            case 24: Binary((left, right) => left * right); break;
            case 26:
                double value = Pop();
                if (value < 0D) throw new InvalidDataException("A CFF sqrt operand is negative.");
                Push(Math.Sqrt(value));
                break;
            case 27: Push(Peek()); break;
            case 28: Exchange(); break;
            case 29: Index(); break;
            case 30: Roll(); break;
            case 34: HorizontalFlex(); break;
            case 35: Flex(); break;
            case 36: HorizontalFlex1(); break;
            case 37: Flex1(); break;
            default: throw new NotSupportedException("CFF escaped CharString operator " + operation + " is not supported.");
        }
    }

    private void HorizontalFlex() {
        RequireCount(7);
        CurveBy(_stack[0], 0D, _stack[1], _stack[2], _stack[3], 0D);
        CurveBy(_stack[4], 0D, _stack[5], -_stack[2], _stack[6], 0D);
        _stack.Clear();
    }

    private void Flex() {
        RequireCount(13);
        CurveBy(0);
        CurveBy(6);
        _stack.Clear();
    }

    private void HorizontalFlex1() {
        RequireCount(9);
        double dy6 = -(_stack[1] + _stack[3] + _stack[7]);
        CurveBy(_stack[0], _stack[1], _stack[2], _stack[3], _stack[4], 0D);
        CurveBy(_stack[5], 0D, _stack[6], _stack[7], _stack[8], dy6);
        _stack.Clear();
    }

    private void Flex1() {
        RequireCount(11);
        double dx = _stack[0] + _stack[2] + _stack[4] + _stack[6] + _stack[8];
        double dy = _stack[1] + _stack[3] + _stack[5] + _stack[7] + _stack[9];
        bool horizontallyDominant = Math.Abs(dx) > Math.Abs(dy);
        double dx6 = horizontallyDominant ? -dx : _stack[10];
        double dy6 = horizontallyDominant ? _stack[10] : -dy;
        CurveBy(0);
        CurveBy(_stack[6], _stack[7], _stack[8], _stack[9], dx6, dy6);
        _stack.Clear();
    }

    private ExecutionResult CallSubroutine(OfficeCffFontData.CffIndex index, int depth) {
        int operand = ToInteger(Pop(), "CFF subroutine index");
        int biased = checked(operand + SubroutineBias(index.Count));
        if (biased < 0 || biased >= index.Count) throw new InvalidDataException("A CFF subroutine index is outside the INDEX.");
        ExecutionResult result = Execute(index[biased], depth + 1);
        if (result == ExecutionResult.Exhausted) {
            throw new InvalidDataException("A CFF CharString subroutine does not terminate with return.");
        }
        return result;
    }

    private void SelectVariationData() {
        if (!_font.IsCff2) throw new InvalidDataException("The CFF1 CharString uses a CFF2 vsindex operator.");
        RequireCount(1);
        _variationDataIndex = ToInteger(_stack[0], "CFF2 vsindex");
        _ = _font.VariationStore?.GetScalars(_variationDataIndex)
            ?? throw new InvalidDataException("The CFF2 CharString uses vsindex without a VariationStore.");
        _stack.Clear();
    }

    private void Blend() {
        if (!_font.IsCff2 || _font.VariationStore == null) throw new InvalidDataException("The CFF CharString uses blend without a CFF2 VariationStore.");
        int valueCount = ToInteger(Pop(), "CFF2 blend count");
        IReadOnlyList<double> scalars = _font.VariationStore.GetScalars(_variationDataIndex);
        int required = checked(valueCount * (scalars.Count + 1));
        if (valueCount < 0 || _stack.Count < required) throw new InvalidDataException("The CFF2 blend operand stack is truncated.");
        int start = _stack.Count - required;
        ApplyBlendDeltas(_stack, start, valueCount, scalars);
        _stack.RemoveRange(start + valueCount, required - valueCount);
    }

    internal static void ApplyBlendDeltas(
        IList<double> stack,
        int start,
        int valueCount,
        IReadOnlyList<double> scalars) {
        for (int valueIndex = 0; valueIndex < valueCount; valueIndex++) {
            double blended = stack[start + valueIndex];
            for (int region = 0; region < scalars.Count; region++) {
                int deltaIndex = start + valueCount + region * valueCount + valueIndex;
                blended += stack[deltaIndex] * scalars[region];
            }
            stack[start + valueIndex] = blended;
        }
    }

    private void ConsumeEndChar() {
        if (!_widthConsumed && (_stack.Count == 1 || _stack.Count == 5)) _stack.RemoveAt(0);
        _widthConsumed = true;
        if (_stack.Count == 4) {
            if (_font.IsCff2) throw new InvalidDataException("A CFF2 CharString uses the CFF1 seac-compatible endchar form.");
            double accentX = _stack[0];
            double accentY = _stack[1];
            int baseCode = ToInteger(_stack[2], "CFF seac base character");
            int accentCode = ToInteger(_stack[3], "CFF seac accent character");
            _stack.Clear();
            RenderSeac(accentX, accentY, baseCode, accentCode);
            return;
        }
        if (_stack.Count != 0) throw new NotSupportedException("The CFF endchar operand form is not supported.");
        CloseContour();
    }

    private void RenderSeac(double accentX, double accentY, int baseCode, int accentCode) {
        if (_seacDepth >= 4) throw new InvalidDataException("CFF seac composition depth is excessive.");
        CloseContour();
        int baseGlyph = _font.ResolveStandardEncodingGlyph(baseCode);
        int accentGlyph = _font.ResolveStandardEncodingGlyph(accentCode);
        var baseInterpreter = new OfficeType2CharStringInterpreter(
            _font,
            baseGlyph,
            _sink,
            _cancellationToken,
            _operationBudget,
            _seacDepth + 1);
        baseInterpreter.Render(_font.GetCharString(baseGlyph));
        var accentInterpreter = new OfficeType2CharStringInterpreter(
            _font,
            accentGlyph,
            new TranslatedPathSink(_sink, accentX, accentY),
            _cancellationToken,
            _operationBudget,
            _seacDepth + 1);
        accentInterpreter.Render(_font.GetCharString(accentGlyph));
    }

    private void LineBy(double dx, double dy) {
        EnsureContour();
        double nextX = _x + dx;
        double nextY = _y + dy;
        EnsureFinitePoint(nextX, nextY);
        _x = nextX;
        _y = nextY;
        _sink.LineTo(_x, _y);
    }

    private void CurveBy(int index) => CurveBy(
        _stack[index], _stack[index + 1],
        _stack[index + 2], _stack[index + 3],
        _stack[index + 4], _stack[index + 5]);

    private void CurveBy(double dx1, double dy1, double dx2, double dy2, double dx3, double dy3) {
        EnsureContour();
        double control1X = _x + dx1;
        double control1Y = _y + dy1;
        double control2X = control1X + dx2;
        double control2Y = control1Y + dy2;
        double nextX = control2X + dx3;
        double nextY = control2Y + dy3;
        EnsureFinitePoint(control1X, control1Y);
        EnsureFinitePoint(control2X, control2Y);
        EnsureFinitePoint(nextX, nextY);
        _x = nextX;
        _y = nextY;
        _sink.CurveTo(control1X, control1Y, control2X, control2Y, nextX, nextY);
    }

    private void EnsureContour() {
        if (_contourOpen) return;
        EnsureFinitePoint(_x, _y);
        _sink.MoveTo(_x, _y);
        _contourOpen = true;
    }

    private void CloseContour() {
        if (!_contourOpen) return;
        _sink.CloseContour();
        _contourOpen = false;
    }

    private void PutTransient() {
        int index = ToInteger(Pop(), "CFF transient-array index");
        double value = Pop();
        if (index < 0 || index >= _transient.Length) throw new InvalidDataException("A CFF transient-array index is invalid.");
        _transient[index] = value;
    }

    private void GetTransient() {
        int index = ToInteger(Pop(), "CFF transient-array index");
        if (index < 0 || index >= _transient.Length) throw new InvalidDataException("A CFF transient-array index is invalid.");
        Push(_transient[index]);
    }

    private sealed class TranslatedPathSink : IOfficeCffPathSink {
        private readonly IOfficeCffPathSink _inner;
        private readonly double _x;
        private readonly double _y;

        internal TranslatedPathSink(IOfficeCffPathSink inner, double x, double y) {
            _inner = inner;
            _x = x;
            _y = y;
        }

        public void MoveTo(double x, double y) {
            double translatedX = x + _x;
            double translatedY = y + _y;
            EnsureFinitePoint(translatedX, translatedY);
            _inner.MoveTo(translatedX, translatedY);
        }

        public void LineTo(double x, double y) {
            double translatedX = x + _x;
            double translatedY = y + _y;
            EnsureFinitePoint(translatedX, translatedY);
            _inner.LineTo(translatedX, translatedY);
        }

        public void CurveTo(
            double control1X,
            double control1Y,
            double control2X,
            double control2Y,
            double x,
            double y) {
            double translatedControl1X = control1X + _x;
            double translatedControl1Y = control1Y + _y;
            double translatedControl2X = control2X + _x;
            double translatedControl2Y = control2Y + _y;
            double translatedX = x + _x;
            double translatedY = y + _y;
            EnsureFinitePoint(translatedControl1X, translatedControl1Y);
            EnsureFinitePoint(translatedControl2X, translatedControl2Y);
            EnsureFinitePoint(translatedX, translatedY);
            _inner.CurveTo(
                translatedControl1X,
                translatedControl1Y,
                translatedControl2X,
                translatedControl2Y,
                translatedX,
                translatedY);
        }
        public void CloseContour() => _inner.CloseContour();
    }

    private static void EnsureFinitePoint(double x, double y) {
        if (double.IsNaN(x) || double.IsInfinity(x) || double.IsNaN(y) || double.IsInfinity(y)) {
            throw new InvalidDataException("A CFF path coordinate is not finite.");
        }
    }

    private void Conditional() {
        double second = Pop();
        double first = Pop();
        double secondValue = Pop();
        double firstValue = Pop();
        Push(first <= second ? firstValue : secondValue);
    }

    private void Exchange() {
        double right = Pop();
        double left = Pop();
        Push(right);
        Push(left);
    }

    private void Index() {
        int index = ToInteger(Pop(), "CFF stack index");
        if (_stack.Count == 0) throw new InvalidDataException("The CFF index operator uses an empty stack.");
        if (index < 0) index = 0;
        if (index >= _stack.Count) index = _stack.Count - 1;
        Push(_stack[_stack.Count - 1 - index]);
    }

    private void Roll() {
        int shift = ToInteger(Pop(), "CFF roll shift");
        int count = ToInteger(Pop(), "CFF roll count");
        if (count < 0 || count > _stack.Count) throw new InvalidDataException("The CFF roll count is invalid.");
        if (count <= 1) return;
        shift %= count;
        if (shift < 0) shift += count;
        if (shift == 0) return;
        int start = _stack.Count - count;
        double[] values = _stack.GetRange(start, count).ToArray();
        for (int index = 0; index < count; index++) _stack[start + ((index + shift) % count)] = values[index];
    }

    private void Binary(Func<double, double, double> operation) {
        double right = Pop();
        double left = Pop();
        Push(operation(left, right));
    }

    private void Push(double value) {
        if (_stack.Count >= MaximumStack || double.IsNaN(value) || double.IsInfinity(value)) {
            throw new InvalidDataException("The CFF CharString operand stack is invalid.");
        }
        _stack.Add(value);
    }

    private double Pop() {
        if (_stack.Count == 0) throw new InvalidDataException("The CFF CharString operand stack underflowed.");
        int index = _stack.Count - 1;
        double value = _stack[index];
        _stack.RemoveAt(index);
        return value;
    }

    private double Peek() {
        if (_stack.Count == 0) throw new InvalidDataException("The CFF CharString operand stack is empty.");
        return _stack[_stack.Count - 1];
    }

    private void RequireCount(int count) {
        if (_stack.Count != count) throw new InvalidDataException("A CFF CharString operator has an invalid operand count.");
    }

    private void RequireMultiple(int divisor, int minimum) {
        if (_stack.Count < minimum || _stack.Count % divisor != 0) throw new InvalidDataException("A CFF CharString operator has an invalid operand count.");
    }

    private static int ToInteger(double value, string description) {
        if (value < int.MinValue || value > int.MaxValue || value != Math.Truncate(value)) {
            throw new InvalidDataException(description + " is not an integer.");
        }
        return checked((int)value);
    }

    private static int SubroutineBias(int count) => count < 1240 ? 107 : count < 33900 ? 1131 : 32768;
}
