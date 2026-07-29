#nullable enable

using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.Serialization;

namespace OfficeIMO.Tabular;

public sealed partial class TabularReader {
    /// <summary>
    /// Streams the current result into objects whose public writable properties match column names.
    /// Column matching is case-insensitive and unmatched columns or properties are ignored.
    /// </summary>
    /// <typeparam name="T">Reference type with a public parameterless constructor.</typeparam>
    /// <returns>Objects created as the current result is traversed.</returns>
#if NET8_0_OR_GREATER
    [RequiresDynamicCode("Object binding compiles typed property assignments. Use typed getters for NativeAOT applications.")]
    [RequiresUnreferencedCode("Object binding discovers writable properties at runtime. Use typed getters for trimmed applications.")]
#endif
    public IEnumerable<T> ReadRecords<T>() where T : class, new() {
        Action<DbDataReader, T>[] bindings = CreateRecordBindings<T>();
        while (Read()) {
            var record = new T();
            for (int index = 0; index < bindings.Length; index++) {
                bindings[index](this, record);
            }

            yield return record;
        }
    }

#if NET8_0_OR_GREATER
    [RequiresDynamicCode("Object binding compiles typed property assignments.")]
    [RequiresUnreferencedCode("Object binding discovers writable properties at runtime.")]
#endif
    private Action<DbDataReader, T>[] CreateRecordBindings<T>() where T : class, new() {
        var ordinals = new Dictionary<string, int>(FieldCount, StringComparer.OrdinalIgnoreCase);
        for (int ordinal = 0; ordinal < FieldCount; ordinal++) {
            ordinals[GetName(ordinal)] = ordinal;
        }

        var bindings = new List<Action<DbDataReader, T>>();
        foreach (PropertyInfo property in typeof(T).GetProperties(BindingFlags.Instance | BindingFlags.Public)) {
            MethodInfo? setter = property.SetMethod;
            string columnName = property.GetCustomAttribute<DataMemberAttribute>()?.Name
                ?? property.Name;
            if (setter == null
                || !setter.IsPublic
                || setter.IsStatic
                || property.GetIndexParameters().Length != 0
                || !ordinals.TryGetValue(columnName, out int ordinal)) {
                continue;
            }

            bindings.Add(CreateRecordBinding<T>(property, ordinal));
        }

        return bindings.ToArray();
    }

#if NET8_0_OR_GREATER
    [RequiresDynamicCode("Object binding compiles typed property assignments.")]
    [RequiresUnreferencedCode("Object binding creates expressions from reflected properties.")]
#endif
    private static Action<DbDataReader, T> CreateRecordBinding<T>(
        PropertyInfo property,
        int ordinal) where T : class {
        var reader = Expression.Parameter(typeof(DbDataReader), "reader");
        var record = Expression.Parameter(typeof(T), "record");
        Type propertyType = property.PropertyType;
        Type valueType = Nullable.GetUnderlyingType(propertyType) ?? propertyType;
        Expression value = CreateRecordValueExpression(reader, valueType, ordinal);
        if (value.Type != propertyType) {
            value = Expression.Convert(value, propertyType);
        }

        if (!propertyType.IsValueType || Nullable.GetUnderlyingType(propertyType) != null) {
            MethodCallExpression isDbNull = Expression.Call(
                reader,
                nameof(DbDataReader.IsDBNull),
                Type.EmptyTypes,
                Expression.Constant(ordinal));
            value = Expression.Condition(isDbNull, Expression.Default(propertyType), value);
        }

        BinaryExpression assign = Expression.Assign(Expression.Property(record, property), value);
        return Expression.Lambda<Action<DbDataReader, T>>(assign, reader, record).Compile();
    }

#if NET8_0_OR_GREATER
    [RequiresDynamicCode("Object binding creates generic getter expressions at runtime.")]
    [RequiresUnreferencedCode("Object binding creates expressions from reflected members.")]
#endif
    private static Expression CreateRecordValueExpression(
        ParameterExpression reader,
        Type valueType,
        int ordinal) {
        string? getter = valueType == typeof(string) ? nameof(DbDataReader.GetString)
            : valueType == typeof(bool) ? nameof(DbDataReader.GetBoolean)
            : valueType == typeof(byte) ? nameof(DbDataReader.GetByte)
            : valueType == typeof(char) ? nameof(DbDataReader.GetChar)
            : valueType == typeof(DateTime) ? nameof(DbDataReader.GetDateTime)
            : valueType == typeof(decimal) ? nameof(DbDataReader.GetDecimal)
            : valueType == typeof(double) ? nameof(DbDataReader.GetDouble)
            : valueType == typeof(float) ? nameof(DbDataReader.GetFloat)
            : valueType == typeof(Guid) ? nameof(DbDataReader.GetGuid)
            : valueType == typeof(short) ? nameof(DbDataReader.GetInt16)
            : valueType == typeof(int) ? nameof(DbDataReader.GetInt32)
            : valueType == typeof(long) ? nameof(DbDataReader.GetInt64)
            : null;

        if (getter != null) {
            return Expression.Call(
                reader,
                getter,
                Type.EmptyTypes,
                Expression.Constant(ordinal));
        }

        MethodCallExpression rawValue = Expression.Call(
            reader,
            nameof(DbDataReader.GetValue),
            Type.EmptyTypes,
            Expression.Constant(ordinal));
        MethodInfo converter = typeof(TabularReader)
            .GetMethod(nameof(ConvertRecordValue), BindingFlags.NonPublic | BindingFlags.Static)!
            .MakeGenericMethod(valueType);
        return Expression.Call(converter, rawValue);
    }

    private static T ConvertRecordValue<T>(object value) {
        if (value is T typed) {
            return typed;
        }

        Type destinationType = typeof(T);
        if (destinationType.IsEnum) {
            return value is string text
                ? (T)Enum.Parse(destinationType, text, ignoreCase: true)
                : (T)Enum.ToObject(destinationType, value);
        }

        return (T)Convert.ChangeType(value, destinationType, CultureInfo.InvariantCulture);
    }
}
