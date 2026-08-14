using System.Data;
using System.Linq.Expressions;
using System.Reflection;
#if NET8_0_OR_GREATER
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;
#endif

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static partial class DirectDataSetWorkbookWriter {
            private const int CompactDataReaderPlanCacheCapacity = 64;
            private static readonly object CompactDataReaderPlanCacheLock = new();
            private static readonly Dictionary<string, CompactDataReaderRowWriter> CompactDataReaderPlanCache = new(StringComparer.Ordinal);
            private static readonly Queue<string> CompactDataReaderPlanCacheOrder = new();

            private delegate void CompactDataReaderRowWriter(
                TextWriter writer,
                IDataRecord reader,
                string?[]? styleAttributes,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                ExcelDateSystem dateSystem);

            private static CompactDataReaderRowWriter? GetCompactDataReaderRowWriter(
                DirectCellValueKind[] cellValueKinds,
                bool[] nullableColumns) {
#if NET8_0_OR_GREATER
                if (!RuntimeFeature.IsDynamicCodeSupported) {
                    return null;
                }
#endif
                string key = CreateCompactDataReaderPlanKey(cellValueKinds, nullableColumns);
                lock (CompactDataReaderPlanCacheLock) {
                    if (CompactDataReaderPlanCache.TryGetValue(key, out CompactDataReaderRowWriter? cached)) {
                        return cached;
                    }
                }

                CompactDataReaderRowWriter? created = CreateCompactDataReaderRowWriter(cellValueKinds, nullableColumns);
                if (created == null) {
                    return null;
                }

                lock (CompactDataReaderPlanCacheLock) {
                    if (CompactDataReaderPlanCache.TryGetValue(key, out CompactDataReaderRowWriter? cached)) {
                        return cached;
                    }

                    if (CompactDataReaderPlanCache.Count >= CompactDataReaderPlanCacheCapacity) {
                        string oldestKey = CompactDataReaderPlanCacheOrder.Dequeue();
                        CompactDataReaderPlanCache.Remove(oldestKey);
                    }

                    CompactDataReaderPlanCache.Add(key, created);
                    CompactDataReaderPlanCacheOrder.Enqueue(key);
                    return created;
                }
            }

            private static string CreateCompactDataReaderPlanKey(
                DirectCellValueKind[] cellValueKinds,
                bool[] nullableColumns) {
                var key = new char[cellValueKinds.Length];
                for (int index = 0; index < cellValueKinds.Length; index++) {
                    key[index] = (char)((int)cellValueKinds[index] | (nullableColumns[index] ? 0x100 : 0));
                }

                return new string(key);
            }

#if NET8_0_OR_GREATER
            [UnconditionalSuppressMessage(
                "Trimming",
                "IL2026",
                Justification = "This schema compiler is reached only when RuntimeFeature.IsDynamicCodeSupported is true; NativeAOT uses the non-compiled writer.")]
            [UnconditionalSuppressMessage(
                "AOT",
                "IL3050",
                Justification = "This schema compiler is reached only when RuntimeFeature.IsDynamicCodeSupported is true; NativeAOT uses the non-compiled writer.")]
#endif
            private static CompactDataReaderRowWriter? CreateCompactDataReaderRowWriter(
                DirectCellValueKind[] cellValueKinds,
                bool[] nullableColumns) {
                try {
                    ParameterExpression writer = Expression.Parameter(typeof(TextWriter), "writer");
                    ParameterExpression reader = Expression.Parameter(typeof(IDataRecord), "reader");
                    ParameterExpression styleAttributes = Expression.Parameter(typeof(string[]), "styleAttributes");
                    ParameterExpression dateTimeOffsetWriteStrategy = Expression.Parameter(
                        typeof(Func<DateTimeOffset, DateTime>),
                        "dateTimeOffsetWriteStrategy");
                    ParameterExpression dateSystem = Expression.Parameter(typeof(ExcelDateSystem), "dateSystem");
                    MethodInfo writeString = typeof(TextWriter).GetMethod(
                        nameof(TextWriter.Write),
                        BindingFlags.Instance | BindingFlags.Public,
                        binder: null,
                        types: [typeof(string)],
                        modifiers: null)!;
                    MethodCallExpression WriteText(Expression text) => Expression.Call(writer, writeString, text);
                    var body = new List<Expression>(2 + (cellValueKinds.Length * 4)) {
                        WriteText(Expression.Constant("<row>"))
                    };
                    var boxedValues = new List<ParameterExpression>(cellValueKinds.Length);

                    for (int columnIndex = 0; columnIndex < cellValueKinds.Length; columnIndex++) {
                        ConstantExpression ordinal = Expression.Constant(columnIndex);
                        Expression? typedValue = CreateCompactDataReaderTypedValue(reader, ordinal, cellValueKinds[columnIndex]);
                        ParameterExpression? boxedValue = typedValue == null
                            ? Expression.Variable(typeof(object), "value" + columnIndex)
                            : null;
                        Expression value = typedValue ?? boxedValue!;
                        if (boxedValue != null) {
                            boxedValues.Add(boxedValue);
                        }

                        Expression? valueWrite = CreateCompactDataReaderValueWrite(
                            writer,
                            value,
                            cellValueKinds[columnIndex],
                            dateTimeOffsetWriteStrategy,
                            dateSystem);
                        if (valueWrite == null) {
                            return null;
                        }

                        BinaryExpression style = Expression.ArrayIndex(styleAttributes, ordinal);
                        if (boxedValue != null) {
                            body.Add(Expression.Assign(
                                boxedValue,
                                Expression.Call(reader, nameof(IDataRecord.GetValue), Type.EmptyTypes, ordinal)));
                        }
                        body.Add(WriteText(Expression.Constant("<c")));
                        body.Add(Expression.IfThen(
                            Expression.AndAlso(
                                Expression.NotEqual(styleAttributes, Expression.Constant(null, typeof(string[]))),
                                Expression.NotEqual(style, Expression.Constant(null, typeof(string)))),
                            WriteText(style)));
                        Expression nullTest = typedValue != null
                            ? nullableColumns[columnIndex]
                                ? Expression.Call(reader, nameof(IDataRecord.IsDBNull), Type.EmptyTypes, ordinal)
                                : Expression.Constant(false)
                            : Expression.OrElse(
                                Expression.Equal(value, Expression.Constant(null, typeof(object))),
                                Expression.ReferenceEqual(value, Expression.Constant(DBNull.Value, typeof(object))));
                        body.Add(nullableColumns[columnIndex] || typedValue == null
                            ? Expression.IfThenElse(
                                nullTest,
                                WriteText(Expression.Constant(" t=\"str\"><v/></c>")),
                                valueWrite)
                            : valueWrite);
                    }

                    body.Add(WriteText(Expression.Constant("</row>")));
                    return Expression.Lambda<CompactDataReaderRowWriter>(
                        Expression.Block(boxedValues, body),
                        writer,
                        reader,
                        styleAttributes,
                        dateTimeOffsetWriteStrategy,
                        dateSystem).Compile();
                } catch (PlatformNotSupportedException) {
                    return null;
                } catch (NotSupportedException) {
                    return null;
                } catch (ArgumentException) {
                    return null;
                } catch (InvalidOperationException) {
                    return null;
                } catch (MemberAccessException) {
                    return null;
                }
            }

            private static bool[] GetDataReaderNullableColumns(IDataReader reader, int columnCount) {
                var nullableColumns = new bool[columnCount];
                for (int index = 0; index < nullableColumns.Length; index++) {
                    nullableColumns[index] = true;
                }
                try {
                    DataTable? schema = reader.GetSchemaTable();
                    if (schema == null || !schema.Columns.Contains("AllowDBNull")) {
                        return nullableColumns;
                    }

                    int availableColumns = Math.Min(columnCount, schema.Rows.Count);
                    for (int index = 0; index < availableColumns; index++) {
                        if (schema.Rows[index]["AllowDBNull"] is bool allowDBNull) {
                            nullableColumns[index] = allowDBNull;
                        }
                    }
                } catch (NotSupportedException) {
                    // Optional provider metadata; conservative null checks remain enabled.
                } catch (NotImplementedException) {
                    // Some valid IDataReader providers expose the API but do not implement
                    // optional schema metadata. Keep conservative null checks enabled.
                }

                return nullableColumns;
            }

#if NET8_0_OR_GREATER
            [UnconditionalSuppressMessage(
                "Trimming",
                "IL2026",
                Justification = "This helper is reached only from the dynamically supported schema compiler; NativeAOT uses the non-compiled writer.")]
            [UnconditionalSuppressMessage(
                "AOT",
                "IL3050",
                Justification = "This helper is reached only from the dynamically supported schema compiler; NativeAOT uses the non-compiled writer.")]
#endif
            private static Expression? CreateCompactDataReaderTypedValue(
                ParameterExpression reader,
                ConstantExpression ordinal,
                DirectCellValueKind cellValueKind) {
                string? getterName = cellValueKind switch {
                    DirectCellValueKind.String => nameof(IDataRecord.GetString),
                    DirectCellValueKind.Boolean => nameof(IDataRecord.GetBoolean),
                    DirectCellValueKind.DateTime => nameof(IDataRecord.GetDateTime),
                    DirectCellValueKind.Double => nameof(IDataRecord.GetDouble),
                    DirectCellValueKind.Float => nameof(IDataRecord.GetFloat),
                    DirectCellValueKind.Decimal => nameof(IDataRecord.GetDecimal),
                    DirectCellValueKind.Byte => nameof(IDataRecord.GetByte),
                    DirectCellValueKind.Int16 => nameof(IDataRecord.GetInt16),
                    DirectCellValueKind.Int32 => nameof(IDataRecord.GetInt32),
                    DirectCellValueKind.Int64 => nameof(IDataRecord.GetInt64),
                    _ => null
                };
                if (getterName == null) {
                    return null;
                }

                MethodCallExpression typedGetter = Expression.Call(
                    reader,
                    getterName,
                    Type.EmptyTypes,
                    ordinal);
                UnaryExpression boxedFallback = Expression.Convert(
                    Expression.Call(reader, nameof(IDataRecord.GetValue), Type.EmptyTypes, ordinal),
                    typedGetter.Type);
                return Expression.TryCatch(
                    typedGetter,
                    Expression.Catch(typeof(NotSupportedException), boxedFallback),
                    Expression.Catch(typeof(NotImplementedException), boxedFallback));
            }

#if NET8_0_OR_GREATER
            [UnconditionalSuppressMessage(
                "Trimming",
                "IL2026",
                Justification = "This helper is reachable only from the dynamically supported schema compiler, and all referenced cell writers are statically preserved.")]
            [UnconditionalSuppressMessage(
                "AOT",
                "IL3050",
                Justification = "This helper is reachable only from the dynamically supported schema compiler; NativeAOT uses the non-compiled writer.")]
#endif
            private static Expression? CreateCompactDataReaderValueWrite(
                ParameterExpression writer,
                Expression value,
                DirectCellValueKind cellValueKind,
                ParameterExpression dateTimeOffsetWriteStrategy,
                ParameterExpression dateSystem) {
                Expression Value(Type valueType) => value.Type == valueType
                    ? value
                    : Expression.Convert(value, valueType);
                Expression Raw(Expression value, Type parameterType) => Expression.Call(
                    typeof(DirectDataSetWorkbookWriter).GetMethod(
                        nameof(WriteRawValueCell),
                        BindingFlags.Static | BindingFlags.NonPublic,
                        binder: null,
                        types: [typeof(TextWriter), parameterType],
                        modifiers: null)!,
                    writer,
                    value.Type == parameterType ? value : Expression.Convert(value, parameterType));

                return cellValueKind switch {
                    DirectCellValueKind.String => Expression.Call(
                        typeof(DirectDataSetWorkbookWriter).GetMethod(
                            nameof(WriteStringCellValue),
                            BindingFlags.Static | BindingFlags.NonPublic)!,
                        writer,
                        Value(typeof(string)),
                        Expression.Constant(null, typeof(DirectSharedStringTable))),
                    DirectCellValueKind.Boolean => Expression.Call(
                        writer,
                        typeof(TextWriter).GetMethod(
                            nameof(TextWriter.Write),
                            BindingFlags.Instance | BindingFlags.Public,
                            binder: null,
                            types: [typeof(string)],
                            modifiers: null)!,
                        Expression.Condition(
                            Value(typeof(bool)),
                            Expression.Constant(" t=\"b\"><v>1</v></c>"),
                            Expression.Constant(" t=\"b\"><v>0</v></c>"))),
                    DirectCellValueKind.DateTime => Expression.Call(
                        typeof(DirectDataSetWorkbookWriter).GetMethod(
                            nameof(WriteDateTimeSerialCell),
                            BindingFlags.Static | BindingFlags.NonPublic)!,
                        writer,
                        Value(typeof(DateTime)),
                        dateSystem),
                    DirectCellValueKind.DateTimeOffset => Expression.Call(
                        typeof(DirectDataSetWorkbookWriter).GetMethod(
                            nameof(WriteDateTimeOffsetCellValue),
                            BindingFlags.Static | BindingFlags.NonPublic)!,
                        writer,
                        Value(typeof(DateTimeOffset)),
                        dateTimeOffsetWriteStrategy,
                        dateSystem),
                    DirectCellValueKind.TimeSpan => Raw(
                        Expression.Property(
                            Value(typeof(TimeSpan)),
                            nameof(TimeSpan.TotalDays)),
                        typeof(double)),
                    DirectCellValueKind.Double => Raw(Value(typeof(double)), typeof(double)),
                    DirectCellValueKind.Float => Raw(Value(typeof(float)), typeof(float)),
                    DirectCellValueKind.Decimal => Raw(Value(typeof(decimal)), typeof(decimal)),
                    DirectCellValueKind.SByte => Raw(Value(typeof(sbyte)), typeof(int)),
                    DirectCellValueKind.Byte => Raw(Value(typeof(byte)), typeof(int)),
                    DirectCellValueKind.Int16 => Raw(Value(typeof(short)), typeof(int)),
                    DirectCellValueKind.UInt16 => Raw(Value(typeof(ushort)), typeof(int)),
                    DirectCellValueKind.Int32 => Raw(Value(typeof(int)), typeof(int)),
                    DirectCellValueKind.UInt32 => Raw(Value(typeof(uint)), typeof(long)),
                    DirectCellValueKind.Int64 => Raw(Value(typeof(long)), typeof(long)),
                    DirectCellValueKind.UInt64 => Raw(Value(typeof(ulong)), typeof(ulong)),
#if NET6_0_OR_GREATER
                    DirectCellValueKind.DateOnly => Expression.Call(
                        typeof(DirectDataSetWorkbookWriter).GetMethod(
                            nameof(WriteDateTimeSerialCell),
                            BindingFlags.Static | BindingFlags.NonPublic)!,
                        writer,
                        Expression.Call(
                            Value(typeof(DateOnly)),
                            nameof(DateOnly.ToDateTime),
                            Type.EmptyTypes,
                            Expression.Property(null, typeof(TimeOnly), nameof(TimeOnly.MinValue))),
                        dateSystem),
                    DirectCellValueKind.TimeOnly => Raw(
                        Expression.Property(
                            Expression.Call(
                                Value(typeof(TimeOnly)),
                                nameof(TimeOnly.ToTimeSpan),
                                Type.EmptyTypes),
                            nameof(TimeSpan.TotalDays)),
                        typeof(double)),
#endif
                    _ => null
                };
            }
        }
    }
}
