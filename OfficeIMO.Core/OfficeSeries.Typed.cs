using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
using System.Numerics;

namespace OfficeIMO;

public sealed partial class OfficeSeries {
    /// <summary>Creates chart values from a numeric selector. A direct property selector supplies its own series name.</summary>
    /// <remarks>Computed selectors require an explicit name. Expression interpretation supports NativeAOT without compiling dynamic code.</remarks>
    public static OfficeSeries From<T, TValue>(IEnumerable<T> data, Expression<Func<T, TValue>> value, string? name = null) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        if (value == null) throw new ArgumentNullException(nameof(value));
        if (!IsNumeric(typeof(TValue))) throw new ArgumentException("A chart series selector must return a numeric type.", nameof(value));
        if (name == null) {
            Expression body = value.Body;
            while (body is UnaryExpression unary && (unary.NodeType == ExpressionType.Convert || unary.NodeType == ExpressionType.ConvertChecked)) body = unary.Operand;
            if (body is not MemberExpression member)
                throw new ArgumentException("Supply a series name for a computed selector.", nameof(name));
            if (member.Member.Name == "Value" && member.Expression is MemberExpression parent && Nullable.GetUnderlyingType(parent.Type) != null)
                member = parent;
            name = member.Member.Name;
        }
        var select = value.Compile(preferInterpretation: true);
        return new OfficeSeries(name, data.Select(row => NumericValue(select(row))));
    }

    internal static bool IsNumeric(Type type) {
        type = Nullable.GetUnderlyingType(type) ?? type;
        if (type.IsEnum) return false;
        if (type == typeof(BigInteger)) return true;
#if NET8_0_OR_GREATER
        if (type == typeof(Half) || type == typeof(Int128) || type == typeof(UInt128)) return true;
#endif
        switch (Type.GetTypeCode(type)) {
            case TypeCode.Byte: case TypeCode.SByte: case TypeCode.Int16: case TypeCode.UInt16:
            case TypeCode.Int32: case TypeCode.UInt32: case TypeCode.Int64: case TypeCode.UInt64:
            case TypeCode.Single: case TypeCode.Double: case TypeCode.Decimal: return true;
            default: return false;
        }
    }

    internal static double NumericValue(object? value) {
        if (value == null || !IsNumeric(value.GetType()))
            throw new ArgumentException("Chart values must be numeric and non-null.", nameof(value));
        if (value is BigInteger integer) return (double)integer;
#if NET8_0_OR_GREATER
        if (value is Half half) return (double)half;
        if (value is Int128 signed) return (double)signed;
        if (value is UInt128 unsigned) return (double)unsigned;
#endif
        return Convert.ToDouble(value, CultureInfo.InvariantCulture);
    }
}
