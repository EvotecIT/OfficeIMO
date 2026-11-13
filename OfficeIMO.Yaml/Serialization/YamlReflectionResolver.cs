using System.Diagnostics.CodeAnalysis;
using System.Numerics;
using System.Reflection;

namespace OfficeIMO.Yaml;

/// <summary>
/// Creates metadata by reflection, the fallback when no <see cref="YamlSerializerContext"/> is given. It applies the
/// same shape rules as the source generator and builds metadata through <see cref="YamlMetadataServices"/>, so both
/// write the same YAML. It needs unreferenced members and runtime code generation, so trimmed and NativeAOT
/// applications use generated metadata instead.
/// </summary>
internal sealed class YamlReflectionResolver : IYamlTypeInfoResolver {
    internal const string UnreferencedCodeMessage =
        "Reflection-based YAML serialization needs members that trimming can remove. Use a YamlSerializerContext ([YamlSerializable]) and the overloads that take YamlTypeInfo<T> or a context.";
    internal const string DynamicCodeMessage =
        "Reflection-based YAML serialization creates generic types at run time, which NativeAOT cannot. Use a YamlSerializerContext ([YamlSerializable]) and the overloads that take YamlTypeInfo<T> or a context.";

    private const BindingFlags Declared = BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly;
    private const BindingFlags Private = BindingFlags.NonPublic | BindingFlags.Static;

    private static YamlReflectionResolver? _instance;

    [RequiresUnreferencedCode(UnreferencedCodeMessage)]
    [RequiresDynamicCode(DynamicCodeMessage)]
    private YamlReflectionResolver() {
    }

    [RequiresUnreferencedCode(UnreferencedCodeMessage)]
    [RequiresDynamicCode(DynamicCodeMessage)]
    internal static YamlReflectionResolver GetInstance() => _instance ??= new YamlReflectionResolver();

    // The constructor carries the annotations: nothing reaches these members without having created the resolver.
    [UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "The resolver is created only through an annotated path.")]
    [UnconditionalSuppressMessage("Trimming", "IL2060", Justification = "The resolver is created only through an annotated path.")]
    [UnconditionalSuppressMessage("Trimming", "IL2070", Justification = "The resolver is created only through an annotated path.")]
    [UnconditionalSuppressMessage("Trimming", "IL2075", Justification = "The resolver is created only through an annotated path.")]
    [UnconditionalSuppressMessage("AOT", "IL3050", Justification = "The resolver is created only through an annotated path.")]
    public YamlTypeInfo? GetTypeInfo(Type type, YamlSerializerOptions options) {
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

        if (YamlTypeShapes.TryGetDictionary(type, out Type? key, out Type? value, out YamlCollectionKind dictionaryKind)) {
            if (!YamlTypeShapes.IsKeyType(key!)) throw Unsupported(type, "dictionary keys must be strings, numbers, enums or other scalar types");
            string method = dictionaryKind switch {
                YamlCollectionKind.Identity => nameof(CreateDictionary),
                YamlCollectionKind.Custom => nameof(CreateCustomDictionary),
                _ => throw Unsupported(type, "it implements a dictionary interface but has no public parameterless constructor")
            };
            return Invoke(method, new[] { type, key!, value! }, options);
        }

        if (YamlTypeShapes.TryGetCollection(type, out Type? element, out YamlCollectionKind collectionKind)) {
            string method = collectionKind switch {
                YamlCollectionKind.Identity => nameof(CreateList),
                YamlCollectionKind.Set => nameof(CreateSet),
                YamlCollectionKind.Custom => nameof(CreateCustomCollection),
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
    private static YamlTypeInfo Invoke(string method, Type[] arguments, YamlSerializerOptions options) {
        MethodInfo generic = typeof(YamlReflectionResolver).GetMethod(method, Private)!.MakeGenericMethod(arguments);
        try {
            return (YamlTypeInfo)generic.Invoke(null, new object[] { options })!;
        } catch (TargetInvocationException exception) when (exception.InnerException is not null) {
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(exception.InnerException).Throw();
            throw;
        }
    }

    private static string? UnsupportedReason(Type type) {
        if (type.IsPointer || type.IsByRef || IsByRefLike(type)) return "pointers, by-reference and ref struct types cannot be serialized";
        if (type.ContainsGenericParameters) return "open generic types cannot be serialized";
        if (typeof(Delegate).IsAssignableFrom(type)) return "delegates cannot be serialized";
        if (typeof(MemberInfo).IsAssignableFrom(type) || type == typeof(IntPtr) || type == typeof(UIntPtr)) return "the type has no YAML representation";
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
        new NotSupportedException("The type " + YamlSerializationErrors.TypeName(type) + " cannot be serialized: " + reason + ".");

    // ------------------------------------------------------------------ generic factories

    private static YamlTypeInfo CreateNullable<T>(YamlSerializerOptions options) where T : struct => YamlMetadataServices.CreateNullableInfo<T>(options);

    private static YamlTypeInfo CreateEnum<T>(YamlSerializerOptions options) where T : struct, Enum => YamlMetadataServices.CreateEnumInfo<T>(options);

    private static YamlTypeInfo CreateArray<T>(YamlSerializerOptions options) =>
        YamlMetadataServices.CreateCollectionInfo<T[], T>(options, static list => list.ToArray());

    private static YamlTypeInfo CreateList<TCollection, T>(YamlSerializerOptions options) where TCollection : IEnumerable<T> =>
        YamlMetadataServices.CreateCollectionInfo<TCollection, T>(options, static list => (TCollection)(object)list);

    private static YamlTypeInfo CreateSet<TCollection, T>(YamlSerializerOptions options) where TCollection : IEnumerable<T> =>
        YamlMetadataServices.CreateCollectionInfo<TCollection, T>(options, static list => (TCollection)(object)new HashSet<T>(list));

    [UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "Reached only from the annotated resolver.")]
    [UnconditionalSuppressMessage("Trimming", "IL2060", Justification = "Reached only from the annotated resolver.")]
    [UnconditionalSuppressMessage("Trimming", "IL2091", Justification = "Reached only from the annotated resolver.")]
    [UnconditionalSuppressMessage("AOT", "IL3050", Justification = "Reached only from the annotated resolver.")]
    private static YamlTypeInfo CreateCustomCollection<TCollection, T>(YamlSerializerOptions options) where TCollection : IEnumerable<T> =>
        YamlMetadataServices.CreateCollectionInfo<TCollection, T>(options, static list => {
            TCollection collection = Activator.CreateInstance<TCollection>();
            var target = (ICollection<T>)collection;
            foreach (T item in list) target.Add(item);
            return collection;
        });

    private static YamlTypeInfo CreateDictionary<TDictionary, TKey, TValue>(YamlSerializerOptions options)
        where TDictionary : IEnumerable<KeyValuePair<TKey, TValue>> where TKey : notnull =>
        YamlMetadataServices.CreateDictionaryInfo<TDictionary, TKey, TValue>(options, static dictionary => (TDictionary)(object)dictionary);

    [UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "Reached only from the annotated resolver.")]
    [UnconditionalSuppressMessage("Trimming", "IL2060", Justification = "Reached only from the annotated resolver.")]
    [UnconditionalSuppressMessage("Trimming", "IL2091", Justification = "Reached only from the annotated resolver.")]
    [UnconditionalSuppressMessage("AOT", "IL3050", Justification = "Reached only from the annotated resolver.")]
    private static YamlTypeInfo CreateCustomDictionary<TDictionary, TKey, TValue>(YamlSerializerOptions options)
        where TDictionary : IEnumerable<KeyValuePair<TKey, TValue>> where TKey : notnull =>
        YamlMetadataServices.CreateDictionaryInfo<TDictionary, TKey, TValue>(options, static dictionary => {
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
    private static YamlTypeInfo CreateObject<T>(YamlSerializerOptions options) {
        Type type = typeof(T);
        var values = new YamlObjectInfoValues<T>();
        List<PropertyInfo> properties = YamlTypeShapes.GetProperties(type);

        if (!type.IsAbstract && !type.IsInterface) {
            ConstructorInfo? constructor = YamlTypeShapes.SelectConstructor(type, out bool useDefault);
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
            var infos = new YamlPropertyInfo[properties.Count];
            MethodInfo factory = typeof(YamlReflectionResolver).GetMethod(nameof(CreateProperty), Private)!;
            for (int i = 0; i < properties.Count; i++) {
                PropertyInfo property = properties[i];
                infos[i] = (YamlPropertyInfo)factory.MakeGenericMethod(type, property.PropertyType).Invoke(null, new object[] { propertyOptions, property })!;
            }
            return infos;
        };

        object[] derived = type.GetCustomAttributes(typeof(OfficeDerivedTypeAttribute), inherit: false);
        if (derived.Length > 0) {
            values.DerivedTypes = Array.ConvertAll(derived, static attribute => (OfficeDerivedTypeAttribute)attribute);
            values.DiscriminatorPropertyName = type.GetCustomAttribute<OfficePolymorphicAttribute>(inherit: false)?.DiscriminatorPropertyName;
        }
        return YamlMetadataServices.CreateObjectInfo(options, values);
    }

    [UnconditionalSuppressMessage("Trimming", "IL2072", Justification = "Reached only from the annotated resolver.")]
    private static YamlParameterInfoValues[] Describe(ParameterInfo[] parameters) {
        var result = new YamlParameterInfoValues[parameters.Length];
        for (int i = 0; i < parameters.Length; i++) {
            ParameterInfo parameter = parameters[i];
            object? defaultValue = parameter.HasDefaultValue && parameter.DefaultValue is not DBNull ? parameter.DefaultValue : null;
            if (defaultValue is null && parameter.ParameterType.IsValueType && Nullable.GetUnderlyingType(parameter.ParameterType) is null) {
                defaultValue = Activator.CreateInstance(parameter.ParameterType);
            } else if (defaultValue is not null && parameter.ParameterType.IsEnum && defaultValue.GetType() != parameter.ParameterType) {
                defaultValue = Enum.ToObject(parameter.ParameterType, defaultValue);
            }
            result[i] = new YamlParameterInfoValues { Name = parameter.Name ?? string.Empty, ParameterType = parameter.ParameterType, Position = i, DefaultValue = defaultValue };
        }
        return result;
    }

    [UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "Reached only from the annotated resolver.")]
    private static YamlPropertyInfo CreateProperty<TDeclaring, TProperty>(YamlSerializerOptions options, PropertyInfo property) {
        var values = new YamlPropertyInfoValues<TProperty> {
            Name = property.Name,
            YamlName = property.GetCustomAttribute<OfficePropertyNameAttribute>()?.Name,
            IsRequired = YamlTypeShapes.IsRequired(property),
            IgnoreCondition = property.GetCustomAttribute<OfficeIgnoreAttribute>()?.Condition
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
        return YamlMetadataServices.CreatePropertyInfo(options, values);
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

/// <summary>How a collection or dictionary type is created from the list or dictionary the serializer reads into.</summary>
internal enum YamlCollectionKind {
    /// <summary>The list or dictionary is the result: <see cref="List{T}"/>, <see cref="Dictionary{TKey, TValue}"/> or an interface they implement.</summary>
    Identity,
    /// <summary>A <see cref="HashSet{T}"/> is built from the list.</summary>
    Set,
    /// <summary>A concrete type with a public parameterless constructor is filled with <c>Add</c>.</summary>
    Custom,
    /// <summary>The type cannot be created.</summary>
    Unsupported
}

/// <summary>
/// The shape rules shared by reflection metadata, mirrored by the source generator: which types are collections and
/// dictionaries, which properties an object has and in what order, and which constructor reading uses.
/// </summary>
[RequiresUnreferencedCode(YamlReflectionResolver.UnreferencedCodeMessage)]
internal static class YamlTypeShapes {
    private const BindingFlags Declared = BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly;

    public static bool TryGetDictionary(Type type, out Type? key, out Type? value, out YamlCollectionKind kind) {
        key = value = null;
        kind = YamlCollectionKind.Unsupported;
        Type? found = null;
        if (type.IsGenericType) {
            Type definition = type.GetGenericTypeDefinition();
            if (definition == typeof(Dictionary<,>) || definition == typeof(IDictionary<,>) || definition == typeof(IReadOnlyDictionary<,>)) {
                found = type;
                kind = YamlCollectionKind.Identity;
            }
        }
        if (found is null) {
            found = FindInterface(type, typeof(IDictionary<,>)) ?? FindInterface(type, typeof(IReadOnlyDictionary<,>));
            if (found is null) return false;
            kind = !type.IsAbstract && !type.IsInterface && FindInterface(type, typeof(IDictionary<,>)) is not null && HasPublicParameterlessConstructor(type)
                ? YamlCollectionKind.Custom
                : YamlCollectionKind.Unsupported;
        }
        Type[] arguments = found.GetGenericArguments();
        key = arguments[0];
        value = arguments[1];
        return true;
    }

    public static bool TryGetCollection(Type type, out Type? element, out YamlCollectionKind kind) {
        element = null;
        kind = YamlCollectionKind.Unsupported;
        if (type == typeof(string)) return false;
        if (type.IsGenericType) {
            Type definition = type.GetGenericTypeDefinition();
            Type argument = type.GetGenericArguments()[0];
            if (definition == typeof(List<>) || definition == typeof(IEnumerable<>) || definition == typeof(ICollection<>) || definition == typeof(IList<>)
                || definition == typeof(IReadOnlyCollection<>) || definition == typeof(IReadOnlyList<>)) {
                element = argument;
                kind = YamlCollectionKind.Identity;
                return true;
            }
            if (definition == typeof(HashSet<>) || definition == typeof(ISet<>) || definition.FullName == "System.Collections.Generic.IReadOnlySet`1") {
                element = argument;
                kind = YamlCollectionKind.Set;
                return true;
            }
        }
        Type? enumerable = null;
        foreach (Type candidate in type.GetInterfaces()) {
            if (!candidate.IsGenericType || candidate.GetGenericTypeDefinition() != typeof(IEnumerable<>)) continue;
            if (enumerable is not null) return true;
            enumerable = candidate;
        }
        if (enumerable is null) return false;
        element = enumerable.GetGenericArguments()[0];
        if (!type.IsAbstract && !type.IsInterface && FindInterface(type, typeof(ICollection<>)) is Type collection
            && collection.GetGenericArguments()[0] == element && HasPublicParameterlessConstructor(type)) {
            kind = YamlCollectionKind.Custom;
        }
        return true;
    }

    /// <summary>Whether a type can be a mapping key: a string or another type written as a single scalar.</summary>
    public static bool IsKeyType(Type type) =>
        type == typeof(string) || type.IsEnum || type == typeof(bool) || type == typeof(char)
        || type == typeof(byte) || type == typeof(sbyte) || type == typeof(short) || type == typeof(ushort)
        || type == typeof(int) || type == typeof(uint) || type == typeof(long) || type == typeof(ulong)
        || type == typeof(float) || type == typeof(double) || type == typeof(decimal) || type == typeof(BigInteger)
        || type == typeof(Guid) || type == typeof(DateTime) || type == typeof(DateTimeOffset) || type == typeof(TimeSpan)
#if NET6_0_OR_GREATER
        || type == typeof(DateOnly) || type == typeof(TimeOnly)
#endif
        || type == typeof(Uri);

    private static Type? FindInterface(Type type, Type definition) {
        if (type.IsGenericType && type.GetGenericTypeDefinition() == definition) return type;
        foreach (Type candidate in type.GetInterfaces()) {
            if (candidate.IsGenericType && candidate.GetGenericTypeDefinition() == definition) return candidate;
        }
        return null;
    }

    private static bool HasPublicParameterlessConstructor(Type type) => type.IsValueType || type.GetConstructor(Type.EmptyTypes) is not null;

    /// <summary>
    /// Returns the serialized properties: public instance properties with a public getter, base types first in
    /// declaration order. A redeclared property keeps its base position with the most derived declaration, and
    /// <see cref="OfficeIgnoreAttribute"/> with <see cref="OfficeIgnoreCondition.Always"/> removes it.
    /// </summary>
    public static List<PropertyInfo> GetProperties(Type type) {
        var chain = new List<Type>();
        if (type.IsInterface) {
            chain.Add(type);
        } else {
            for (Type? current = type; current is not null && current != typeof(object) && current != typeof(ValueType); current = current.BaseType) chain.Insert(0, current);
        }
        var result = new List<PropertyInfo?>();
        var positions = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (Type declaring in chain) {
            foreach (PropertyInfo property in declaring.GetProperties(Declared)) {
                if (property.GetIndexParameters().Length > 0) continue;
                bool ignored = property.GetCustomAttribute<OfficeIgnoreAttribute>() is { Condition: OfficeIgnoreCondition.Always };
                bool readable = property.GetMethod is { IsPublic: true, IsStatic: false };
                PropertyInfo? entry = ignored || !readable ? null : property;
                if (positions.TryGetValue(property.Name, out int position)) {
                    result[position] = entry;
                } else {
                    positions.Add(property.Name, result.Count);
                    result.Add(entry);
                }
            }
        }
        var properties = new List<PropertyInfo>(result.Count);
        foreach (PropertyInfo? property in result) {
            if (property is not null) properties.Add(property);
        }
        return properties;
    }

    /// <summary>Whether a property is required: <see cref="OfficeRequiredAttribute"/> or a C# <c>required</c> member.</summary>
    public static bool IsRequired(PropertyInfo property) {
        if (Attribute.IsDefined(property, typeof(OfficeRequiredAttribute), inherit: true)) return true;
        foreach (CustomAttributeData attribute in property.CustomAttributes) {
            if (attribute.AttributeType.FullName == "System.Runtime.CompilerServices.RequiredMemberAttribute") return true;
        }
        return false;
    }

    /// <summary>
    /// Selects the constructor reading uses: the one marked <see cref="OfficeConstructorAttribute"/>; for a struct, its
    /// parameterless form; the public parameterless constructor; or the only public constructor. Returns null when
    /// none qualifies. <paramref name="useDefault"/> is set for a struct without a marked constructor.
    /// </summary>
    public static ConstructorInfo? SelectConstructor(Type type, out bool useDefault) {
        useDefault = false;
        ConstructorInfo[] constructors = type.GetConstructors(BindingFlags.Public | BindingFlags.Instance);
        ConstructorInfo? marked = null;
        foreach (ConstructorInfo constructor in constructors) {
            if (!constructor.IsDefined(typeof(OfficeConstructorAttribute), inherit: false)) continue;
            if (marked is not null) throw new NotSupportedException("The type " + YamlSerializationErrors.TypeName(type) + " marks more than one constructor with [OfficeConstructor].");
            marked = constructor;
        }
        if (marked is not null) return marked;
        if (type.IsValueType) {
            useDefault = true;
            return null;
        }
        foreach (ConstructorInfo constructor in constructors) {
            if (constructor.GetParameters().Length == 0) return constructor;
        }
        return constructors.Length == 1 ? constructors[0] : null;
    }
}
