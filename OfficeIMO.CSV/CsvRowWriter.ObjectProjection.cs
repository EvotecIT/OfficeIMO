#if NET8_0_OR_GREATER
using System.Diagnostics.CodeAnalysis;
using System.Linq.Expressions;
using System.Reflection;
using OfficeIMO.Core.Internal;
using OfficeIMO.SharedSource.IO;

namespace OfficeIMO.CSV;

public sealed partial class CsvRowWriter
{
    private Func<object, CsvRowWriter, CultureInfo, bool>? _bufferedPropertyProjector;
    private (int Start, int Length)[]? _propertyFormattedRanges;
    private int _propertyFallbackColumn = -1;
    private int _propertyFallbackPrefixLength;

    /// <summary>Creates a typed row projector after the output schema is established.</summary>
    [RequiresUnreferencedCode("This compatibility method discovers properties from runtime objects. For NativeAOT, use WriteDataReader or write explicit CSV rows.")]
    private bool TryCreateProjectedObjectWriter(object item, IReadOnlyList<string> columns)
    {
        if (_delimiter != ',') return false;
        bool useSpanBuffer = _writer.GetType() == typeof(StreamWriter) || _writer is PooledUtf8TextWriter || _stringWriterBuffer != null;
        ObjectDataHelpers.PropertyWriteExpressionFactory factory;
        if (_useDefaultWritePath)
            factory = useSpanBuffer ? CreatePropertySpanWriteExpression : CreatePropertyWriteExpression;
        else if (_useFormattedDefaultWritePath && useSpanBuffer)
            factory = CreateFormattedPropertySpanWriteExpression;
        else
            return false;
        if (!ObjectDataHelpers.TryCreatePropertyWriter<CsvRowWriter>(item, columns, factory, out var projector)) return false;
        if (useSpanBuffer) _propertyCharacters ??= new char[1024];
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
        _propertyCharacterCount = 0;
        _propertyTextValues.AsSpan().Clear();
        _propertyFallbackColumn = -1;
        if (!_bufferedPropertyProjector(item, this, _options.Culture)) return false;
        if (_propertyFallbackColumn >= 0)
        {
            WriteProjectedFallback();
        }
        else if (_propertyCharacters != null)
        {
            string newLine = _options.NewLine;
            ReadOnlySpan<char> terminator = newLine.AsSpan();
            if (terminator.TryCopyTo(_propertyCharacters.AsSpan(_propertyCharacterCount)))
            {
                _writer.Write(_propertyCharacters.AsSpan(0, _propertyCharacterCount + terminator.Length));
            }
            else
            {
                _writer.Write(_propertyCharacters.AsSpan(0, _propertyCharacterCount));
                _writer.Write(newLine);
            }
        }
        else
        {
            CsvWriter.WriteBufferedRecordLine(_writer, _rowBuffer, _options.NewLine);
        }
        return true;
    }

    /// <summary>Delivers the staged prefix and remaining owned values after all getters complete.</summary>
    private void WriteProjectedFallback()
    {
        for (int index = _propertyFallbackColumn; index < _propertyTextValues!.Length; index++)
        {
            var range = _propertyFormattedRanges![index];
            if (range.Length > 0)
                _propertyTextValues[index] = _rowBuffer.ToString(range.Start, range.Length);
        }
        string newLine = _options.NewLine;
        if (_propertyFallbackPrefixLength > 0)
        {
            if (_propertyCharacters != null)
                _writer.Write(_propertyCharacters.AsSpan(0, _propertyFallbackPrefixLength));
            else
                _writer.Write(_rowBuffer.ToString(0, _propertyFallbackPrefixLength));
        }
        if (_propertyFallbackColumn > 0) _writer.Write(_delimiter);
        var remaining = new ArraySegment<string?>(_propertyTextValues, _propertyFallbackColumn,
            _propertyTextValues.Length - _propertyFallbackColumn);
        CsvWriter.WriteRecord(_writer, remaining, _delimiter, newLine, CultureInfo.InvariantCulture,
            CsvFormulaInjectionPolicy.Preserve, CsvQuoteMode.AsNeeded, null, null);
    }

    /// <summary>Bounds the staged prefix and retains owned text after overflow.</summary>
    private void StageProjectedText(int ordinal, string? text)
    {
        if (_propertyFallbackColumn >= 0)
        {
            _propertyTextValues![ordinal] = text;
            _propertyFormattedRanges![ordinal] = default;
            return;
        }
        long worstCaseLength = (long)_rowBuffer.Length + (text?.Length ?? 0) * 2L + 3;
        if (worstCaseLength > _rowBuffer.Capacity)
        {
            _propertyFallbackColumn = ordinal;
            _propertyFallbackPrefixLength = _rowBuffer.Length;
            _propertyTextValues![ordinal] = text;
            _propertyFormattedRanges![ordinal] = default;
            return;
        }
        if (ordinal > 0) _rowBuffer.Append(_delimiter);
        if (text != null) CsvWriter.WriteEscapedDefault(_rowBuffer, text, _delimiter);
    }

    /// <summary>Formats an invariant built-in scalar through the existing constrained formatter.</summary>
    private static void StageProjectedScalar<T>(CsvRowWriter writer, int ordinal, T value)
        where T : ISpanFormattable
    {
        if (writer._propertyFallbackColumn < 0)
        {
            if (ordinal > 0) writer._rowBuffer.Append(writer._delimiter);
            CsvWriter.AppendKnownValueDefault(writer._rowBuffer, value, writer._delimiter, CultureInfo.InvariantCulture);
            return;
        }
        writer._propertyTextValues![ordinal] = null;
        int start = writer._rowBuffer.Length;
        CsvWriter.AppendKnownValueDefault(writer._rowBuffer, value, writer._delimiter, CultureInfo.InvariantCulture);
        writer._propertyFormattedRanges![ordinal] = (start, writer._rowBuffer.Length - start);
    }

    /// <summary>Roots the concrete scalar staging method for the selected row representation.</summary>
    internal static MethodInfo GetPropertyScalarStageMethod<T>(bool useSpanBuffer = false) where T : ISpanFormattable
    {
        Action<CsvRowWriter, int, T> action = useSpanBuffer ? StageProjectedScalarSpan<T> : StageProjectedScalar<T>;
        return action.Method;
    }

    /// <summary>Builds the existing staged property writes for a supplied text writer.</summary>
    [RequiresUnreferencedCode("Builds property expressions from runtime schema metadata, including nullable property accessors.")]
    private static Expression CreatePropertyWriteExpression(ParameterExpression value,
        Expression formattedText, ParameterExpression writer, ParameterExpression culture, int ordinal, out bool usesTypedValue) =>
        CreatePropertyWriteExpressionCore(value, formattedText, writer, culture, ordinal, false, out usesTypedValue);

    /// <summary>Builds property writes into the bounded character row buffer.</summary>
    [RequiresUnreferencedCode("Builds typed property expressions from runtime schema metadata.")]
    private static Expression CreatePropertySpanWriteExpression(ParameterExpression value,
        Expression formattedText, ParameterExpression writer, ParameterExpression culture, int ordinal, out bool usesTypedValue) =>
        CreatePropertyWriteExpressionCore(value, formattedText, writer, culture, ordinal, true, out usesTypedValue);

    /// <summary>Reuses the scalar classification and captured-culture rules for both staged representations.</summary>
    [RequiresUnreferencedCode("Builds property expressions from runtime schema metadata, including nullable property accessors.")]
    private static Expression CreatePropertyWriteExpressionCore(ParameterExpression value,
        Expression formattedText, ParameterExpression writer, ParameterExpression culture, int ordinal,
        bool useSpanBuffer, out bool usesTypedValue)
    {
        MethodInfo textMethod = typeof(CsvRowWriter).GetMethod(useSpanBuffer ? nameof(StageProjectedSpanText) : nameof(StageProjectedText),
            BindingFlags.NonPublic | BindingFlags.Instance)!;
        Expression text = Expression.Call(writer, textMethod, Expression.Constant(ordinal), formattedText);
        Type scalarType = Nullable.GetUnderlyingType(value.Type) ?? value.Type;
        MethodInfo? scalarMethod = CsvWriter.GetPropertyScalarStageMethod(scalarType, useSpanBuffer);
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
