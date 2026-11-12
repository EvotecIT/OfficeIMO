namespace OfficeIMO.Yaml;

/// <summary>
/// Typed reads of scalars through the document schema, sharing the resolution tables with <see cref="YamlReader"/>.
/// Generic dispatch compares <c>typeof(T)</c>, which the JIT folds per instantiation, so value types are not boxed.
/// </summary>
internal static class YamlValueConverter {
    public static bool TryConvert<T>(YamlScalar scalar, out T value) {
        Type type = typeof(T);
        Type? underlying = Nullable.GetUnderlyingType(type);
        if (type == typeof(string) || underlying is not null) {
            if (scalar.ResolvesToNull()) {
                value = default!;
                return true;
            }
            if (type == typeof(string)) {
                value = (T)(object)scalar.Value;
                return true;
            }
            type = underlying!;
        }

        if (type == typeof(bool)) {
            bool ok = TryBoolean(scalar, out bool result);
            value = ok ? (T)(object)result : default!;
            return ok;
        }
        if (type == typeof(long) || type == typeof(int)) {
            bool ok = TryInt64(scalar, out long result);
            if (type == typeof(int)) {
                ok = ok && result >= int.MinValue && result <= int.MaxValue;
                value = ok ? (T)(object)(int)result : default!;
            } else {
                value = ok ? (T)(object)result : default!;
            }
            return ok;
        }
        if (type == typeof(double) || type == typeof(float)) {
            bool ok = TryDouble(scalar, out double result);
            if (type == typeof(float)) value = ok ? (T)(object)(float)result : default!;
            else value = ok ? (T)(object)result : default!;
            return ok;
        }
        if (type == typeof(decimal)) {
            bool ok = TryDecimal(scalar, out decimal result);
            value = ok ? (T)(object)result : default!;
            return ok;
        }
        throw Unsupported(typeof(T));
    }

    public static void RequireSupported<T>() {
        Type type = Nullable.GetUnderlyingType(typeof(T)) ?? typeof(T);
        if (type == typeof(string) || type == typeof(bool) || type == typeof(int) || type == typeof(long)
            || type == typeof(double) || type == typeof(float) || type == typeof(decimal)) return;
        throw Unsupported(typeof(T));
    }

    /// <summary>Returns a short form of the scalar's text for error messages.</summary>
    public static string Describe(YamlScalar scalar) {
        string text = scalar.Value;
        return text.Length <= 40 ? text : text.Substring(0, 37) + "...";
    }

    private static bool TryBoolean(YamlScalar scalar, out bool value) {
        value = false;
        return scalar.TryGetKind(out YamlScalarKind kind) && kind == YamlScalarKind.Boolean
            && YamlSchemaResolver.TryParseBoolean(scalar.ValueSpan, scalar.NumericSchema, out value);
    }

    private static bool TryInt64(YamlScalar scalar, out long value) {
        value = 0;
        return scalar.TryGetKind(out YamlScalarKind kind) && kind == YamlScalarKind.Integer
            && YamlSchemaResolver.TryParseInt64(scalar.ValueSpan, scalar.NumericSchema, out value);
    }

    private static bool TryDouble(YamlScalar scalar, out double value) {
        value = 0;
        return scalar.TryGetKind(out YamlScalarKind kind) && kind is YamlScalarKind.Integer or YamlScalarKind.Float
            && YamlSchemaResolver.TryParseDouble(scalar.ValueSpan, scalar.NumericSchema, out value);
    }

    private static bool TryDecimal(YamlScalar scalar, out decimal value) {
        value = 0;
        return scalar.TryGetKind(out YamlScalarKind kind) && kind is YamlScalarKind.Integer or YamlScalarKind.Float
            && YamlSchemaResolver.TryParseDecimal(scalar.ValueSpan, scalar.NumericSchema, out value);
    }

    private static NotSupportedException Unsupported(Type type) =>
        new NotSupportedException("YAML scalars read as string, bool, int, long, double, float, decimal or their nullable forms, not " + type.Name + ".");
}
