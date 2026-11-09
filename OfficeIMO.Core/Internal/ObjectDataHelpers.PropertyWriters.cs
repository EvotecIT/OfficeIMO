#if NET8_0_OR_GREATER
using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Core.Internal;

internal static partial class ObjectDataHelpers
{
    internal delegate Expression PropertyWriteExpressionFactory(ParameterExpression value,
        Expression formattedText, ParameterExpression writer, ParameterExpression culture, int ordinal, out bool usesTypedValue);

    /// <summary>Creates a typed property writer using the shared schema and formatting rules.</summary>
    [RequiresUnreferencedCode("Uses reflection over arbitrary object graphs. For AOT-safe usage, map values explicitly or pre-flatten items.")]
    internal static bool TryCreatePropertyWriter<TWriter>(object item, IReadOnlyList<string> columns,
        PropertyWriteExpressionFactory factory, out Func<object, TWriter, CultureInfo, bool>? projector)
        where TWriter : class
    {
        ArgumentNullException.ThrowIfNull(item);
        ArgumentNullException.ThrowIfNull(columns);
        ArgumentNullException.ThrowIfNull(factory);
        projector = null;
        if (!RuntimeFeature.IsDynamicCodeSupported || IsDictionaryLike(item)) return false;
        return GetPropertyPlan(item.GetType()).TryCreatePropertyWriter(columns, factory, out projector);
    }

    private sealed partial class ObjectPropertyPlan
    {
        private ConcurrentDictionary<Tuple<Type, PropertyWriteExpressionFactory>, Delegate?>? _propertyWriters;

        internal bool TryCreatePropertyWriter<TWriter>(IReadOnlyList<string> columns,
            PropertyWriteExpressionFactory factory, out Func<object, TWriter, CultureInfo, bool>? projector)
            where TWriter : class
        {
            if (IsDefaultColumnOrder(columns))
            {
                var cache = _propertyWriters ??= new ConcurrentDictionary<Tuple<Type, PropertyWriteExpressionFactory>, Delegate?>();
                var key = Tuple.Create(typeof(TWriter), factory);
                projector = (Func<object, TWriter, CultureInfo, bool>?)cache.GetOrAdd(key,
                    _ => CreateCompiledPropertyWriter<TWriter>(_type, _properties, factory));
                return projector != null;
            }

            var properties = new PropertyInfo[columns.Count];
            for (int index = 0; index < columns.Count; index++)
            {
                if (!_propertiesByName.TryGetValue(columns[index], out var property))
                {
                    projector = null;
                    return false;
                }
                properties[index] = property;
            }
            projector = CreateCompiledPropertyWriter<TWriter>(_type, properties, factory);
            return projector != null;
        }

        private static Func<object, TWriter, CultureInfo, bool>? CreateCompiledPropertyWriter<TWriter>(
            Type type, PropertyInfo[] properties, PropertyWriteExpressionFactory factory) where TWriter : class
        {
            var item = Expression.Parameter(typeof(object), "item");
            var writer = Expression.Parameter(typeof(TWriter), "writer");
            var culture = Expression.Parameter(typeof(CultureInfo), "culture");
            var typedItem = Expression.Variable(type, "typedItem");
            var returnTarget = Expression.Label(typeof(bool));
            var variables = new List<ParameterExpression> { typedItem };
            var expressions = new List<Expression>();
            expressions.Add(Expression.IfThen(
                Expression.OrElse(Expression.Equal(item, Expression.Constant(null, typeof(object))),
                    Expression.NotEqual(Expression.Call(item, typeof(object).GetMethod(nameof(GetType))!),
                        Expression.Constant(type))),
                Expression.Return(returnTarget, Expression.Constant(false))));
            expressions.Add(Expression.IfThen(Expression.Equal(writer, Expression.Constant(null, typeof(TWriter))),
                Expression.Throw(Expression.New(typeof(ArgumentNullException).GetConstructor(new[] { typeof(string) })!,
                    Expression.Constant("writer")))));
            expressions.Add(Expression.IfThen(Expression.Equal(culture, Expression.Constant(null, typeof(CultureInfo))),
                Expression.Throw(Expression.New(typeof(ArgumentNullException).GetConstructor(new[] { typeof(string) })!,
                    Expression.Constant("culture")))));
            expressions.Add(Expression.Assign(typedItem, Expression.Convert(item, type)));
            MethodInfo format = typeof(ObjectPropertyPlan).GetMethod(nameof(FormatProjectedValue),
                BindingFlags.NonPublic | BindingFlags.Static)!;
            bool hasTypedValues = false;
            for (int index = 0; index < properties.Length; index++)
            {
                var value = Expression.Variable(properties[index].PropertyType, "value" + index);
                variables.Add(value);
                expressions.Add(Expression.Assign(value, Expression.Property(typedItem, properties[index])));
                Expression text = Expression.Call(format, Expression.Convert(value, typeof(object)), culture);
                expressions.Add(factory(value, text, writer, culture, index, out bool usesTypedValue));
                hasTypedValues |= usesTypedValue;
            }
            if (!hasTypedValues) return null;
            expressions.Add(Expression.Return(returnTarget, Expression.Constant(true)));
            expressions.Add(Expression.Label(returnTarget, Expression.Constant(false)));
            return Expression.Lambda<Func<object, TWriter, CultureInfo, bool>>(
                Expression.Block(variables, expressions), item, writer, culture).Compile();
        }
    }
}
#endif
