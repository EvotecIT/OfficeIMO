using System.Diagnostics.CodeAnalysis;
using System.Numerics;
using System.Reflection;
using OfficeIMO.Serialization;

namespace OfficeIMO.Toml;

/// <summary>
/// Creates metadata by reflection, the fallback when no <see cref="TomlSerializerContext"/> is given. It applies the
/// same shape rules as the source generator and builds metadata through <see cref="TomlMetadataServices"/>, so both
/// write the same TOML. It needs unreferenced members and runtime code generation, so trimmed and NativeAOT
/// applications use generated metadata instead.
/// </summary>
internal sealed class TomlReflectionResolver : ITomlTypeInfoResolver {
    internal const string UnreferencedCodeMessage =
        "Reflection-based TOML serialization needs members that trimming can remove. Use a TomlSerializerContext ([TomlSerializable]) and the overloads that take TomlTypeInfo<T> or a context.";
    internal const string DynamicCodeMessage =
        "Reflection-based TOML serialization creates generic types at run time, which NativeAOT cannot. Use a TomlSerializerContext ([TomlSerializable]) and the overloads that take TomlTypeInfo<T> or a context.";

    private const BindingFlags Declared = BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly;
    private const BindingFlags Private = BindingFlags.NonPublic | BindingFlags.Static;

    private static TomlReflectionResolver? _instance;

    [RequiresUnreferencedCode(UnreferencedCodeMessage)]
    [RequiresDynamicCode(DynamicCodeMessage)]
    private TomlReflectionResolver() {
    }

    [RequiresUnreferencedCode(UnreferencedCodeMessage)]
    [RequiresDynamicCode(DynamicCodeMessage)]
    internal static TomlReflectionResolver GetInstance() => _instance ??= new TomlReflectionResolver();

    // The constructor carries the annotations: nothing reaches these members without having created the resolver.
    [UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "The resolver is created only through an annotated path.")]
    [UnconditionalSuppressMessage("Trimming", "IL2060", Justification = "The resolver is created only through an annotated path.")]
    [UnconditionalSuppressMessage("Trimming", "IL2070", Justification = "The resolver is created only through an annotated path.")]
    [UnconditionalSuppressMessage("Trimming", "IL2075", Justification = "The resolver is created only through an annotated path.")]
    [UnconditionalSuppressMessage("AOT", "IL3050", Justification = "The resolver is created only through an annotated path.")]
    public TomlTypeInfo? GetTypeInfo(Type type, TomlSerializerOptions options) {
        if (type is null) throw new ArgumentNullException(nameof(type));
        if (options is null) throw new ArgumentNullException(nameof(options));
        string? reason = UnsupportedReason(type);
        if (reason is not null) throw Unsupported(type, reason);

        Type? underlying = Nullable.GetUnderlyingType(type);
        if (underlying is not null) return Invoke(nameof(CreateNullable), new[] { underlying }, options);
        if (type.IsEnum) return Invoke(nameof(CreateEnum), new[] { type }, options);
        if (type.IsArray) {
            if (type.GetArrayRank() != 1 || type != type.GetElementType()!.MakeArrayType()) throw Unsupported(type, "only one-dimensional, zero-based arrays are supported");
            return Invoke(nameof(CreateArray), new[] { type.GetElementType()! }, options);
        }

        if (SerializerTypeShapes.TryGetDictionary(type, out Type? key, out Type? value, out SerializerCollectionKind dictionaryKind)) {
            if (key != typeof(TomlDateTime) && !SerializerTypeShapes.IsKeyType(key!)) throw Unsupported(type, "dictionary keys must be strings, numbers, enums or other scalar types");
            string method = dictionaryKind switch {
                SerializerCollectionKind.Identity => nameof(CreateDictionary),
                SerializerCollectionKind.Custom => nameof(CreateCustomDictionary),
                _ => throw Unsupported(type, "it implements a dictionary interface but has no public parameterless constructor")
            };
            return Invoke(method, new[] { type, key!, value! }, options);
        }

        if (SerializerTypeShapes.TryGetCollection(type, out Type? element, out SerializerCollectionKind collectionKind)) {
            string method = collectionKind switch {
                SerializerCollectionKind.Identity => nameof(CreateList),
                SerializerCollectionKind.Set => nameof(CreateSet),
                SerializerCollectionKind.Custom => nameof(CreateCustomCollection),
                _ => throw Unsupported(type, "it is a collection without a public parameterless constructor, or implements IEnumerable<T> more than once")
            };
            return Invoke(method, new[] { type, element! }, options);
        }
        if (typeof(System.Collections.IEnumerable).IsAssignableFrom(type)) throw Unsupported(type, "non-generic collections are not supported");

        bool polymorphic = type.GetCustomAttributes(typeof(OfficeDerivedTypeAttribute), inherit: false).Length > 0;
        if ((type.IsAbstract || type.IsInterface) && !polymorphic) throw Unsupported(type, "it is abstract and registers no derived types with [OfficeDerivedType]");
        return Invoke(nameof(CreateObject), new[] { type }, options);
    }

    [UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "Reached only from the annotated resolver.")]
    [UnconditionalSuppressMessage("Trimming", "IL2060", Justification = "Reached only from the annotated resolver.")]
    [UnconditionalSuppressMessage("Trimming", "IL2091", Justification = "Reached only from the annotated resolver.")]
    [UnconditionalSuppressMessage("AOT", "IL3050", Justification = "Reached only from the annotated resolver.")]
    private static TomlTypeInfo Invoke(string method, Type[] arguments, TomlSerializerOptions options) {
        MethodInfo generic = typeof(TomlReflectionResolver).GetMethod(method, Private)!.MakeGenericMethod(arguments);
        try {
            return (TomlTypeInfo)generic.Invoke(null, new object[] { options })!;
        } catch (TargetInvocationException exception) when (exception.InnerException is not null) {
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(exception.InnerException).Throw();
            throw;
        }
    }

    private static string? UnsupportedReason(Type type) {
        if (type.IsPointer || type.IsByRef || IsByRefLike(type)) return "pointers, by-reference and ref struct types cannot be serialized";
        if (type.ContainsGenericParameters) return "open generic types cannot be serialized";
        if (typeof(Delegate).IsAssignableFrom(type)) return "delegates cannot be serialized";
        if (typeof(MemberInfo).IsAssignableFrom(type) || type == typeof(IntPtr) || type == typeof(UIntPtr)) return "the type has no TOML representation";
        return null;
    }

    private static bool IsByRefLike(Type type) {
#if NET5_0_OR_GREATER
        return type.IsByRefLike;
#else
        foreach (CustomAttributeData attribute in type.CustomAttributes) {
            if (attribute.AttributeType.FullName == "System.Runtime.CompilerServices.IsByRefLikeAttribute") return true;
        }
        return false;
#endif
    }

    private static NotSupportedException Unsupported(Type type, string reason) =>
        new NotSupportedException("The type " + TomlSerializationErrors.TypeName(type) + " cannot be serialized: " + reason + ".");

    // ------------------------------------------------------------------ generic factories

    private static TomlTypeInfo CreateNullable<T>(TomlSerializerOptions options) where T : struct => TomlMetadataServices.CreateNullableInfo<T>(options);

    private static TomlTypeInfo CreateEnum<T>(TomlSerializerOptions options) where T : struct, Enum => TomlMetadataServices.CreateEnumInfo<T>(options);

    private static TomlTypeInfo CreateArray<T>(TomlSerializerOptions options) =>
        TomlMetadataServices.CreateCollectionInfo<T[], T>(options, static list => list.ToArray());

    private static TomlTypeInfo CreateList<TCollection, T>(TomlSerializerOptions options) where TCollection : IEnumerable<T> =>
        TomlMetadataServices.CreateCollectionInfo<TCollection, T>(options, static list => (TCollection)(object)list);

    private static TomlTypeInfo CreateSet<TCollection, T>(TomlSerializerOptions options) where TCollection : IEnumerable<T> =>
        TomlMetadataServices.CreateCollectionInfo<TCollection, T>(options, static list => (TCollection)(object)new HashSet<T>(list));

    [UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "Reached only from the annotated resolver.")]
    [UnconditionalSuppressMessage("Trimming", "IL2060", Justification = "Reached only from the annotated resolver.")]
    [UnconditionalSuppressMessage("Trimming", "IL2091", Justification = "Reached only from the annotated resolver.")]
    [UnconditionalSuppressMessage("AOT", "IL3050", Justification = "Reached only from the annotated resolver.")]
    private static TomlTypeInfo CreateCustomCollection<TCollection, T>(TomlSerializerOptions options) where TCollection : IEnumerable<T> =>
        TomlMetadataServices.CreateCollectionInfo<TCollection, T>(options, static list => {
            TCollection collection = Activator.CreateInstance<TCollection>();
            var target = (ICollection<T>)collection;
            foreach (T item in list) target.Add(item);
            return collection;
        });

    private static TomlTypeInfo CreateDictionary<TDictionary, TKey, TValue>(TomlSerializerOptions options)
        where TDictionary : IEnumerable<KeyValuePair<TKey, TValue>> where TKey : notnull =>
        TomlMetadataServices.CreateDictionaryInfo<TDictionary, TKey, TValue>(options, static dictionary => (TDictionary)(object)dictionary);

    [UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "Reached only from the annotated resolver.")]
    [UnconditionalSuppressMessage("Trimming", "IL2060", Justification = "Reached only from the annotated resolver.")]
    [UnconditionalSuppressMessage("Trimming", "IL2091", Justification = "Reached only from the annotated resolver.")]
    [UnconditionalSuppressMessage("AOT", "IL3050", Justification = "Reached only from the annotated resolver.")]
    private static TomlTypeInfo CreateCustomDictionary<TDictionary, TKey, TValue>(TomlSerializerOptions options)
        where TDictionary : IEnumerable<KeyValuePair<TKey, TValue>> where TKey : notnull =>
        TomlMetadataServices.CreateDictionaryInfo<TDictionary, TKey, TValue>(options, static dictionary => {
            TDictionary result = Activator.CreateInstance<TDictionary>();
            var target = (IDictionary<TKey, TValue>)result;
            foreach (KeyValuePair<TKey, TValue> entry in dictionary) target.Add(entry.Key, entry.Value);
            return result;
        });

    [UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "Reached only from the annotated resolver.")]
    [UnconditionalSuppressMessage("Trimming", "IL2060", Justification = "Reached only from the annotated resolver.")]
    [UnconditionalSuppressMessage("Trimming", "IL2070", Justification = "Reached only from the annotated resolver.")]
    [UnconditionalSuppressMessage("Trimming", "IL2072", Justification = "Reached only from the annotated resolver.")]
    [UnconditionalSuppressMessage("Trimming", "IL2075", Justification = "Reached only from the annotated resolver.")]
    [UnconditionalSuppressMessage("Trimming", "IL2087", Justification = "Reached only from the annotated resolver.")]
    [UnconditionalSuppressMessage("Trimming", "IL2090", Justification = "Reached only from the annotated resolver.")]
    [UnconditionalSuppressMessage("Trimming", "IL2091", Justification = "Reached only from the annotated resolver.")]
    [UnconditionalSuppressMessage("AOT", "IL3050", Justification = "Reached only from the annotated resolver.")]
    private static TomlTypeInfo CreateObject<T>(TomlSerializerOptions options) {
        Type type = typeof(T);
        var values = new TomlObjectInfoValues<T>();
        List<PropertyInfo> properties = SerializerTypeShapes.GetProperties(type);

        if (!type.IsAbstract && !type.IsInterface) {
            ConstructorInfo? constructor = SerializerTypeShapes.SelectConstructor(type, out bool useDefault);
            if (useDefault) {
                values.ObjectCreator = static () => Activator.CreateInstance<T>();
            } else if (constructor is not null) {
                ParameterInfo[] parameters = constructor.GetParameters();
                if (parameters.Length == 0) {
                    values.ObjectCreator = static () => Activator.CreateInstance<T>();
                } else {
                    values.ObjectWithParameterizedConstructorCreator = arguments => (T)constructor.Invoke(arguments);
                    values.ConstructorParameterMetadataInitializer = () => Describe(parameters);
                }
            }
        }

        values.PropertyMetadataInitializer = propertyOptions => {
            var infos = new TomlPropertyInfo[properties.Count];
            MethodInfo factory = typeof(TomlReflectionResolver).GetMethod(nameof(CreateProperty), Private)!;
            for (int i = 0; i < properties.Count; i++) {
                PropertyInfo property = properties[i];
                infos[i] = (TomlPropertyInfo)factory.MakeGenericMethod(type, property.PropertyType).Invoke(null, new object[] { propertyOptions, property })!;
            }
            return infos;
        };

        object[] derived = type.GetCustomAttributes(typeof(OfficeDerivedTypeAttribute), inherit: false);
        if (derived.Length > 0) {
            values.DerivedTypes = Array.ConvertAll(derived, static attribute => (OfficeDerivedTypeAttribute)attribute);
            values.DiscriminatorPropertyName = type.GetCustomAttribute<OfficePolymorphicAttribute>(inherit: false)?.DiscriminatorPropertyName;
        }
        values.IsInlineTable = type.IsDefined(typeof(TomlInlineTableAttribute), inherit: false);
        return TomlMetadataServices.CreateObjectInfo(options, values);
    }

    [UnconditionalSuppressMessage("Trimming", "IL2072", Justification = "Reached only from the annotated resolver.")]
    private static TomlParameterInfoValues[] Describe(ParameterInfo[] parameters) {
        var result = new TomlParameterInfoValues[parameters.Length];
        for (int i = 0; i < parameters.Length; i++) {
            ParameterInfo parameter = parameters[i];
            object? defaultValue = parameter.HasDefaultValue && parameter.DefaultValue is not DBNull ? parameter.DefaultValue : null;
            if (defaultValue is null && parameter.ParameterType.IsValueType && Nullable.GetUnderlyingType(parameter.ParameterType) is null) {
                defaultValue = Activator.CreateInstance(parameter.ParameterType);
            } else if (defaultValue is not null && parameter.ParameterType.IsEnum && defaultValue.GetType() != parameter.ParameterType) {
                defaultValue = Enum.ToObject(parameter.ParameterType, defaultValue);
            }
            result[i] = new TomlParameterInfoValues { Name = parameter.Name ?? string.Empty, ParameterType = parameter.ParameterType, Position = i, DefaultValue = defaultValue };
        }
        return result;
    }

    [UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "Reached only from the annotated resolver.")]
    private static TomlPropertyInfo CreateProperty<TDeclaring, TProperty>(TomlSerializerOptions options, PropertyInfo property) {
        var values = new TomlPropertyInfoValues<TProperty> {
            Name = property.Name,
            TomlName = property.GetCustomAttribute<OfficePropertyNameAttribute>()?.Name,
            IsRequired = SerializerTypeShapes.IsRequired(property),
            IgnoreCondition = property.GetCustomAttribute<OfficeIgnoreAttribute>()?.Condition,
            IsInlineTable = property.IsDefined(typeof(TomlInlineTableAttribute), inherit: false)
        };
        MethodInfo getter = property.GetMethod!;
        MethodInfo? setter = property.SetMethod is { IsPublic: true } set ? set : null;
        if (typeof(TDeclaring).IsValueType) {
            values.Getter = owner => (TProperty)property.GetValue(owner)!;
            if (setter is not null) values.Setter = (owner, value) => property.SetValue(owner, value);
        } else {
            values.Getter = CreateGetter<TDeclaring, TProperty>(getter, property);
            if (setter is not null) values.Setter = CreateSetter<TDeclaring, TProperty>(setter, property);
        }
        return TomlMetadataServices.CreatePropertyInfo(options, values);
    }

    private static Func<object, TProperty> CreateGetter<TDeclaring, TProperty>(MethodInfo getter, PropertyInfo property) {
        try {
            var typed = (Func<TDeclaring, TProperty>)Delegate.CreateDelegate(typeof(Func<TDeclaring, TProperty>), getter);
            return owner => typed((TDeclaring)owner);
        } catch (ArgumentException) {
            return owner => (TProperty)property.GetValue(owner)!;
        }
    }

    private static Action<object, TProperty> CreateSetter<TDeclaring, TProperty>(MethodInfo setter, PropertyInfo property) {
        try {
            var typed = (Action<TDeclaring, TProperty>)Delegate.CreateDelegate(typeof(Action<TDeclaring, TProperty>), setter);
            return (owner, value) => typed((TDeclaring)owner, value);
        } catch (ArgumentException) {
            return (owner, value) => property.SetValue(owner, value);
        }
    }
}
