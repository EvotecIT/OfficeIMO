using System.Diagnostics.CodeAnalysis;
using System.Numerics;
using System.Reflection;

namespace OfficeIMO.Serialization;

/// <summary>How a collection or dictionary type is created from the list or dictionary the serializer reads into.</summary>
internal enum SerializerCollectionKind {
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
/// The shape rules of reflection metadata for the YAML and TOML serializers, mirrored by their shared source generator
/// parser: which types are collections and
/// dictionaries, which properties an object has and in what order, and which constructor reading uses.
/// </summary>
[RequiresUnreferencedCode("The shape rules read members that trimming can remove; generated metadata applies them at build time.")]
internal static class SerializerTypeShapes {
    private const BindingFlags Declared = BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly;

    public static bool TryGetDictionary(Type type, out Type? key, out Type? value, out SerializerCollectionKind kind) {
        key = value = null;
        kind = SerializerCollectionKind.Unsupported;
        Type? found = null;
        if (type.IsGenericType) {
            Type definition = type.GetGenericTypeDefinition();
            if (definition == typeof(Dictionary<,>) || definition == typeof(IDictionary<,>) || definition == typeof(IReadOnlyDictionary<,>)) {
                found = type;
                kind = SerializerCollectionKind.Identity;
            }
        }
        if (found is null) {
            found = FindInterface(type, typeof(IDictionary<,>)) ?? FindInterface(type, typeof(IReadOnlyDictionary<,>));
            if (found is null) return false;
            kind = !type.IsAbstract && !type.IsInterface && FindInterface(type, typeof(IDictionary<,>)) is not null && HasPublicParameterlessConstructor(type)
                ? SerializerCollectionKind.Custom
                : SerializerCollectionKind.Unsupported;
        }
        Type[] arguments = found.GetGenericArguments();
        key = arguments[0];
        value = arguments[1];
        return true;
    }

    public static bool TryGetCollection(Type type, out Type? element, out SerializerCollectionKind kind) {
        element = null;
        kind = SerializerCollectionKind.Unsupported;
        if (type == typeof(string)) return false;
        if (type.IsGenericType) {
            Type definition = type.GetGenericTypeDefinition();
            Type argument = type.GetGenericArguments()[0];
            if (definition == typeof(List<>) || definition == typeof(IEnumerable<>) || definition == typeof(ICollection<>) || definition == typeof(IList<>)
                || definition == typeof(IReadOnlyCollection<>) || definition == typeof(IReadOnlyList<>)) {
                element = argument;
                kind = SerializerCollectionKind.Identity;
                return true;
            }
            if (definition == typeof(HashSet<>) || definition == typeof(ISet<>) || definition.FullName == "System.Collections.Generic.IReadOnlySet`1") {
                element = argument;
                kind = SerializerCollectionKind.Set;
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
            kind = SerializerCollectionKind.Custom;
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
            if (marked is not null) throw new NotSupportedException("The type " + TypeName(type) + " marks more than one constructor with [OfficeConstructor].");
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

    private static string TypeName(Type type) {
        Type? underlying = Nullable.GetUnderlyingType(type);
        if (underlying is not null) return TypeName(underlying) + "?";
        if (!type.IsGenericType) return type.Name;
        int tick = type.Name.IndexOf('`');
        return (tick > 0 ? type.Name.Substring(0, tick) : type.Name) + "<" + string.Join(", ", Array.ConvertAll(type.GetGenericArguments(), TypeName)) + ">";
    }
}
