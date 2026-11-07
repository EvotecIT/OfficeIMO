#if NET8_0_OR_GREATER
using System.Diagnostics.CodeAnalysis;
using System.Linq.Expressions;
using System.Reflection;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.CSV;

public sealed partial class CsvRowWriter
{
    private Func<object, CsvRowWriter, CultureInfo, bool>? _bufferedPropertyProjector;
    private (int Start, int Length)[]? _propertyFormattedRanges;
    private bool _propertyBufferFallback;

    /// <summary>Creates a typed row projector after the output schema is established.</summary>
    [RequiresUnreferencedCode("This compatibility method discovers properties from runtime objects. For NativeAOT, use WriteDataReader or write explicit CSV rows.")]
    private bool TryCreateProjectedObjectWriter(object item, IReadOnlyList<string> columns)
    {
        if (!_useDefaultWritePath || _delimiter != ',' || _stringWriterBuffer != null ||
            !ObjectDataHelpers.TryCreatePropertyWriter<CsvRowWriter>(item, columns,
                CreatePropertyWriteExpression, out var projector)) return false;
        _bufferedPropertyProjector = projector;
        _propertyTextValues = new string?[columns.Count];
        _propertyFormattedRanges = new (int Start, int Length)[columns.Count];
        return TryWriteProjectedObject(item);
    }

    /// <summary>Stages every property before delivering its row to the output writer.</summary>
    private bool TryWriteProjectedObject(object item)
    {
        if (_bufferedPropertyProjector == null) return false;
        _rowBuffer.Clear();
        _propertyBufferFallback = false;
        if (!_bufferedPropertyProjector(item, this, _options.Culture)) return false;
        if (_propertyBufferFallback)
        {
            for (int index = 0; index < _propertyTextValues!.Length; index++)
            {
                var range = _propertyFormattedRanges![index];
                if (range.Length > 0)
                    _propertyTextValues[index] = _rowBuffer.ToString(range.Start, range.Length);
            }
            WriteTextBuffered(_propertyTextValues);
        }
        else
        {
            CsvWriter.WriteBufferedRecordLine(_writer, _rowBuffer, _options.NewLine);
        }
        return true;
    }

    /// <summary>Retains owned text and bounds its speculative row-buffer copy.</summary>
    private void StageProjectedText(int ordinal, string? text)
    {
        _propertyTextValues![ordinal] = text;
        _propertyFormattedRanges![ordinal] = default;
        if (_propertyBufferFallback) return;
        long worstCaseLength = (long)_rowBuffer.Length + (text?.Length ?? 0) * 2L + 3;
        if (worstCaseLength > _rowBuffer.Capacity)
        {
            _propertyBufferFallback = true;
            return;
        }
        if (ordinal > 0) _rowBuffer.Append(_delimiter);
        if (text != null) CsvWriter.WriteEscapedDefault(_rowBuffer, text, _delimiter);
    }

    /// <summary>Formats an invariant built-in scalar through the existing constrained formatter.</summary>
    private static void StageProjectedScalar<T>(CsvRowWriter writer, int ordinal, T value)
        where T : ISpanFormattable
    {
        writer._propertyTextValues![ordinal] = null;
        if (!writer._propertyBufferFallback && ordinal > 0) writer._rowBuffer.Append(writer._delimiter);
        int start = writer._rowBuffer.Length;
        CsvWriter.AppendKnownValueDefault(writer._rowBuffer, value, writer._delimiter, CultureInfo.InvariantCulture);
        writer._propertyFormattedRanges![ordinal] = (start, writer._rowBuffer.Length - start);
    }

    internal static MethodInfo GetPropertyScalarStageMethod<T>() where T : ISpanFormattable
    {
        Action<CsvRowWriter, int, T> action = StageProjectedScalar<T>;
        return action.Method;
    }

    [RequiresUnreferencedCode("Builds property expressions from runtime schema metadata, including nullable property accessors.")]
    private static Expression CreatePropertyWriteExpression(ParameterExpression value,
        Expression formattedText, ParameterExpression writer, ParameterExpression culture, int ordinal, out bool usesTypedValue)
    {
        MethodInfo textMethod = typeof(CsvRowWriter).GetMethod(nameof(StageProjectedText),
            BindingFlags.NonPublic | BindingFlags.Instance)!;
        Expression text = Expression.Call(writer, textMethod, Expression.Constant(ordinal), formattedText);
        Type scalarType = Nullable.GetUnderlyingType(value.Type) ?? value.Type;
        MethodInfo? scalarMethod = CsvWriter.GetPropertyScalarStageMethod(scalarType);
        usesTypedValue = scalarMethod != null;
        if (scalarMethod == null) return text;
        Expression scalarValue = value;
        if (scalarType != value.Type) scalarValue = Expression.Property(value, "Value");
        Expression scalar = Expression.Call(scalarMethod, writer, Expression.Constant(ordinal), scalarValue);
        if (scalarType != value.Type)
        {
            Expression empty = Expression.Call(writer, textMethod, Expression.Constant(ordinal),
                Expression.Constant(null, typeof(string)));
            scalar = Expression.IfThenElse(Expression.Property(value, "HasValue"), scalar, empty);
        }
        Expression invariant = Expression.ReferenceEqual(culture,
            Expression.Property(null, typeof(CultureInfo), nameof(CultureInfo.InvariantCulture)));
        return Expression.IfThenElse(invariant, scalar, text);
    }
}
#endif
