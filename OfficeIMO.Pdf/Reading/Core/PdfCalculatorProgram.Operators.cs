namespace OfficeIMO.Pdf;

internal sealed partial class PdfCalculatorProgram {
    private static bool TryReadOperator(string token, out CalculatorOperator calculatorOperator) {
        switch (token) {
            case "abs": calculatorOperator = CalculatorOperator.Abs; return true;
            case "add": calculatorOperator = CalculatorOperator.Add; return true;
            case "and": calculatorOperator = CalculatorOperator.And; return true;
            case "atan": calculatorOperator = CalculatorOperator.Atan; return true;
            case "bitshift": calculatorOperator = CalculatorOperator.Bitshift; return true;
            case "ceiling": calculatorOperator = CalculatorOperator.Ceiling; return true;
            case "copy": calculatorOperator = CalculatorOperator.Copy; return true;
            case "cos": calculatorOperator = CalculatorOperator.Cos; return true;
            case "cvi": calculatorOperator = CalculatorOperator.Cvi; return true;
            case "cvr": calculatorOperator = CalculatorOperator.Cvr; return true;
            case "div": calculatorOperator = CalculatorOperator.Div; return true;
            case "dup": calculatorOperator = CalculatorOperator.Dup; return true;
            case "eq": calculatorOperator = CalculatorOperator.Eq; return true;
            case "exch": calculatorOperator = CalculatorOperator.Exch; return true;
            case "exp": calculatorOperator = CalculatorOperator.Exp; return true;
            case "floor": calculatorOperator = CalculatorOperator.Floor; return true;
            case "ge": calculatorOperator = CalculatorOperator.Ge; return true;
            case "gt": calculatorOperator = CalculatorOperator.Gt; return true;
            case "idiv": calculatorOperator = CalculatorOperator.Idiv; return true;
            case "index": calculatorOperator = CalculatorOperator.Index; return true;
            case "le": calculatorOperator = CalculatorOperator.Le; return true;
            case "ln": calculatorOperator = CalculatorOperator.Ln; return true;
            case "log": calculatorOperator = CalculatorOperator.Log; return true;
            case "lt": calculatorOperator = CalculatorOperator.Lt; return true;
            case "mod": calculatorOperator = CalculatorOperator.Mod; return true;
            case "mul": calculatorOperator = CalculatorOperator.Mul; return true;
            case "ne": calculatorOperator = CalculatorOperator.Ne; return true;
            case "neg": calculatorOperator = CalculatorOperator.Neg; return true;
            case "not": calculatorOperator = CalculatorOperator.Not; return true;
            case "or": calculatorOperator = CalculatorOperator.Or; return true;
            case "pop": calculatorOperator = CalculatorOperator.Pop; return true;
            case "roll": calculatorOperator = CalculatorOperator.Roll; return true;
            case "round": calculatorOperator = CalculatorOperator.Round; return true;
            case "sin": calculatorOperator = CalculatorOperator.Sin; return true;
            case "sqrt": calculatorOperator = CalculatorOperator.Sqrt; return true;
            case "sub": calculatorOperator = CalculatorOperator.Sub; return true;
            case "truncate": calculatorOperator = CalculatorOperator.Truncate; return true;
            case "xor": calculatorOperator = CalculatorOperator.Xor; return true;
            default:
                calculatorOperator = default;
                return false;
        }
    }

    private static bool TryExecuteOperator(CalculatorOperator calculatorOperator, CalculatorStack stack) {
        switch (calculatorOperator) {
            case CalculatorOperator.Abs: return TryAbsolute(stack);
            case CalculatorOperator.Add: return TryIntegerAwareBinary(stack, BinaryArithmetic.Add);
            case CalculatorOperator.And: return TryBooleanOrIntegerBinary(stack, BooleanIntegerOperation.And);
            case CalculatorOperator.Atan: return TryAtan(stack);
            case CalculatorOperator.Bitshift: return TryBitshift(stack);
            case CalculatorOperator.Ceiling: return TryIntegralRealOperator(stack, IntegralRealOperation.Ceiling);
            case CalculatorOperator.Copy:
                return stack.TryPopInteger(out int copyCount) && stack.TryCopy(copyCount);
            case CalculatorOperator.Cos: return TryTrigonometric(stack, cosine: true);
            case CalculatorOperator.Cvi: return TryConvertToInteger(stack);
            case CalculatorOperator.Cvr: return TryConvertToReal(stack);
            case CalculatorOperator.Div: return TryDivision(stack);
            case CalculatorOperator.Dup:
                return stack.TryPeek(0, out Value duplicate) && stack.TryPush(duplicate);
            case CalculatorOperator.Eq: return TryEquality(stack, equal: true);
            case CalculatorOperator.Exch: return stack.TryExchange();
            case CalculatorOperator.Exp: return TryExponentiation(stack);
            case CalculatorOperator.Floor: return TryIntegralRealOperator(stack, IntegralRealOperation.Floor);
            case CalculatorOperator.Ge: return TryRelational(stack, RelationalOperation.GreaterThanOrEqual);
            case CalculatorOperator.Gt: return TryRelational(stack, RelationalOperation.GreaterThan);
            case CalculatorOperator.Idiv: return TryIntegerDivision(stack, remainder: false);
            case CalculatorOperator.Index:
                return stack.TryPopInteger(out int depth) && depth >= 0 && stack.TryPeek(depth, out Value indexed) && stack.TryPush(indexed);
            case CalculatorOperator.Le: return TryRelational(stack, RelationalOperation.LessThanOrEqual);
            case CalculatorOperator.Ln: return TryLogarithm(stack, natural: true);
            case CalculatorOperator.Log: return TryLogarithm(stack, natural: false);
            case CalculatorOperator.Lt: return TryRelational(stack, RelationalOperation.LessThan);
            case CalculatorOperator.Mod: return TryIntegerDivision(stack, remainder: true);
            case CalculatorOperator.Mul: return TryIntegerAwareBinary(stack, BinaryArithmetic.Multiply);
            case CalculatorOperator.Ne: return TryEquality(stack, equal: false);
            case CalculatorOperator.Neg: return TryNegate(stack);
            case CalculatorOperator.Not: return TryNot(stack);
            case CalculatorOperator.Or: return TryBooleanOrIntegerBinary(stack, BooleanIntegerOperation.Or);
            case CalculatorOperator.Pop: return stack.TryPopValue();
            case CalculatorOperator.Roll:
                return stack.TryPopInteger(out int shift) && stack.TryPopInteger(out int rollCount) && stack.TryRoll(rollCount, shift);
            case CalculatorOperator.Round: return TryIntegralRealOperator(stack, IntegralRealOperation.Round);
            case CalculatorOperator.Sin: return TryTrigonometric(stack, cosine: false);
            case CalculatorOperator.Sqrt: return TrySquareRoot(stack);
            case CalculatorOperator.Sub: return TryIntegerAwareBinary(stack, BinaryArithmetic.Subtract);
            case CalculatorOperator.Truncate: return TryIntegralRealOperator(stack, IntegralRealOperation.Truncate);
            case CalculatorOperator.Xor: return TryBooleanOrIntegerBinary(stack, BooleanIntegerOperation.Xor);
            default: return false;
        }
    }

    private static bool TryAbsolute(CalculatorStack stack) {
        if (!stack.TryPop(out Value value)) return false;
        if (value.Kind == ValueKind.Integer) {
            if (value.IntegerValue == int.MinValue) return stack.TryPush(Value.Real(-(double)int.MinValue));
            return stack.TryPush(Value.Integer(Math.Abs(value.IntegerValue)));
        }
        return value.Kind == ValueKind.Real && TryPushFiniteReal(stack, Math.Abs(value.RealValue));
    }

    private static bool TryNegate(CalculatorStack stack) {
        if (!stack.TryPop(out Value value)) return false;
        if (value.Kind == ValueKind.Integer) {
            if (value.IntegerValue == int.MinValue) return stack.TryPush(Value.Real(-(double)int.MinValue));
            return stack.TryPush(Value.Integer(-value.IntegerValue));
        }
        return value.Kind == ValueKind.Real && TryPushFiniteReal(stack, -value.RealValue);
    }

    private static bool TryIntegerAwareBinary(CalculatorStack stack, BinaryArithmetic operation) {
        if (!stack.TryPop(out Value right) || !stack.TryPop(out Value left) ||
            !left.TryGetNumber(out double leftNumber) || !right.TryGetNumber(out double rightNumber)) return false;
        if (left.Kind == ValueKind.Integer && right.Kind == ValueKind.Integer) {
            long integerResult = operation switch {
                BinaryArithmetic.Add => (long)left.IntegerValue + right.IntegerValue,
                BinaryArithmetic.Subtract => (long)left.IntegerValue - right.IntegerValue,
                _ => (long)left.IntegerValue * right.IntegerValue
            };
            if (integerResult >= int.MinValue && integerResult <= int.MaxValue) {
                return stack.TryPush(Value.Integer((int)integerResult));
            }
        }
        double result = operation switch {
            BinaryArithmetic.Add => leftNumber + rightNumber,
            BinaryArithmetic.Subtract => leftNumber - rightNumber,
            _ => leftNumber * rightNumber
        };
        return TryPushFiniteReal(stack, result);
    }

    private static bool TryDivision(CalculatorStack stack) {
        if (!TryPopTwoNumbers(stack, out double left, out double right) || right == 0D) return false;
        return TryPushFiniteReal(stack, left / right);
    }

    private static bool TryIntegerDivision(CalculatorStack stack, bool remainder) {
        if (!stack.TryPopInteger(out int right) || !stack.TryPopInteger(out int left) || right == 0 ||
            (left == int.MinValue && right == -1)) return false;
        return stack.TryPush(Value.Integer(remainder ? left % right : left / right));
    }

    private static bool TryExponentiation(CalculatorStack stack) {
        if (!TryPopTwoNumbers(stack, out double basis, out double exponent)) return false;
        return TryPushFiniteReal(stack, Math.Pow(basis, exponent));
    }

    private static bool TrySquareRoot(CalculatorStack stack) {
        if (!TryPopNumber(stack, out double value) || value < 0D) return false;
        return TryPushFiniteReal(stack, Math.Sqrt(value));
    }

    private static bool TryLogarithm(CalculatorStack stack, bool natural) {
        if (!TryPopNumber(stack, out double value) || value <= 0D) return false;
        return TryPushFiniteReal(stack, natural ? Math.Log(value) : Math.Log10(value));
    }

    private static bool TryAtan(CalculatorStack stack) {
        if (!TryPopTwoNumbers(stack, out double numerator, out double denominator) ||
            (numerator == 0D && denominator == 0D)) return false;
        double angle = Math.Atan2(numerator, denominator) * (180D / Math.PI);
        if (angle < 0D) angle += 360D;
        return TryPushFiniteReal(stack, angle);
    }

    private static bool TryTrigonometric(CalculatorStack stack, bool cosine) {
        if (!TryPopNumber(stack, out double degrees)) return false;
        double radians = degrees * (Math.PI / 180D);
        return TryPushFiniteReal(stack, cosine ? Math.Cos(radians) : Math.Sin(radians));
    }

    private static bool TryIntegralRealOperator(CalculatorStack stack, IntegralRealOperation operation) {
        if (!stack.TryPop(out Value value)) return false;
        if (value.Kind == ValueKind.Integer) return stack.TryPush(value);
        if (value.Kind != ValueKind.Real) return false;
        double result = operation switch {
            IntegralRealOperation.Ceiling => Math.Ceiling(value.RealValue),
            IntegralRealOperation.Floor => Math.Floor(value.RealValue),
            IntegralRealOperation.Round => RoundToNearestGreaterInteger(value.RealValue),
            _ => Math.Truncate(value.RealValue)
        };
        return TryPushFiniteReal(stack, result);
    }

    private static double RoundToNearestGreaterInteger(double value) {
        double truncated = Math.Truncate(value);
        double remainder = value - truncated;
        if (remainder >= 0.5D) return truncated + 1D;
        if (remainder < -0.5D) return truncated - 1D;
        return truncated;
    }

    private static bool TryConvertToInteger(CalculatorStack stack) {
        if (!stack.TryPop(out Value value)) return false;
        if (value.Kind == ValueKind.Integer) return stack.TryPush(value);
        if (value.Kind != ValueKind.Real) return false;
        double truncated = Math.Truncate(value.RealValue);
        if (!IsFinite(truncated) || truncated < int.MinValue || truncated > int.MaxValue) return false;
        return stack.TryPush(Value.Integer((int)truncated));
    }

    private static bool TryConvertToReal(CalculatorStack stack) {
        if (!stack.TryPop(out Value value) || !value.TryGetNumber(out double number)) return false;
        return TryPushFiniteReal(stack, number);
    }

    private static bool TryEquality(CalculatorStack stack, bool equal) {
        if (!stack.TryPop(out Value right) || !stack.TryPop(out Value left)) return false;
        bool result;
        if (left.Kind == ValueKind.Boolean && right.Kind == ValueKind.Boolean) {
            result = left.BooleanValue == right.BooleanValue;
        } else if (left.TryGetNumber(out double leftNumber) && right.TryGetNumber(out double rightNumber)) {
            result = leftNumber == rightNumber;
        } else {
            result = false;
        }
        return stack.TryPush(Value.Boolean(equal ? result : !result));
    }

    private static bool TryRelational(CalculatorStack stack, RelationalOperation operation) {
        if (!TryPopTwoNumbers(stack, out double left, out double right)) return false;
        bool result = operation switch {
            RelationalOperation.GreaterThan => left > right,
            RelationalOperation.GreaterThanOrEqual => left >= right,
            RelationalOperation.LessThan => left < right,
            _ => left <= right
        };
        return stack.TryPush(Value.Boolean(result));
    }

    private static bool TryBooleanOrIntegerBinary(CalculatorStack stack, BooleanIntegerOperation operation) {
        if (!stack.TryPop(out Value right) || !stack.TryPop(out Value left)) return false;
        if (left.Kind == ValueKind.Boolean && right.Kind == ValueKind.Boolean) {
            bool result = operation switch {
                BooleanIntegerOperation.And => left.BooleanValue && right.BooleanValue,
                BooleanIntegerOperation.Or => left.BooleanValue || right.BooleanValue,
                _ => left.BooleanValue ^ right.BooleanValue
            };
            return stack.TryPush(Value.Boolean(result));
        }
        if (left.Kind != ValueKind.Integer || right.Kind != ValueKind.Integer) return false;
        int integerResult = operation switch {
            BooleanIntegerOperation.And => left.IntegerValue & right.IntegerValue,
            BooleanIntegerOperation.Or => left.IntegerValue | right.IntegerValue,
            _ => left.IntegerValue ^ right.IntegerValue
        };
        return stack.TryPush(Value.Integer(integerResult));
    }

    private static bool TryNot(CalculatorStack stack) {
        if (!stack.TryPop(out Value value)) return false;
        if (value.Kind == ValueKind.Boolean) return stack.TryPush(Value.Boolean(!value.BooleanValue));
        return value.Kind == ValueKind.Integer && stack.TryPush(Value.Integer(~value.IntegerValue));
    }

    private static bool TryBitshift(CalculatorStack stack) {
        if (!stack.TryPopInteger(out int shift) || !stack.TryPopInteger(out int value)) return false;
        int result;
        if (shift >= 32 || shift <= -32) {
            result = 0;
        } else if (shift >= 0) {
            result = unchecked(value << shift);
        } else {
            result = unchecked((int)((uint)value >> -shift));
        }
        return stack.TryPush(Value.Integer(result));
    }

    private static bool TryPopNumber(CalculatorStack stack, out double value) {
        value = 0D;
        return stack.TryPop(out Value item) && item.TryGetNumber(out value);
    }

    private static bool TryPopTwoNumbers(CalculatorStack stack, out double left, out double right) {
        left = 0D;
        right = 0D;
        return TryPopNumber(stack, out right) && TryPopNumber(stack, out left);
    }

    private static bool TryPushFiniteReal(CalculatorStack stack, double value) =>
        IsFinite(value) && stack.TryPush(Value.Real(value));

    private enum BinaryArithmetic { Add, Subtract, Multiply }
    private enum BooleanIntegerOperation { And, Or, Xor }
    private enum IntegralRealOperation { Ceiling, Floor, Round, Truncate }
    private enum RelationalOperation { GreaterThan, GreaterThanOrEqual, LessThan, LessThanOrEqual }
}
