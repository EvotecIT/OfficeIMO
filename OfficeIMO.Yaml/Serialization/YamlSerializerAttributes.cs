namespace OfficeIMO.Yaml;

/// <summary>
/// Sets the mapping key a property is written and read under. It wins over
/// <see cref="YamlSerializerOptions.PropertyNamingPolicy"/>.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class YamlPropertyNameAttribute : Attribute {
    /// <summary>Creates the attribute.</summary>
    /// <param name="name">The mapping key, used as written.</param>
    public YamlPropertyNameAttribute(string name) {
        Name = name ?? throw new ArgumentNullException(nameof(name));
    }

    /// <summary>Gets the mapping key.</summary>
    public string Name { get; }
}

/// <summary>
/// Leaves a property out of serialization: always by default, or only when its value is <see langword="null"/> or the
/// type's default when <see cref="Condition"/> says so.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class YamlIgnoreAttribute : Attribute {
    /// <summary>
    /// Gets or sets when the property is left out. Defaults to <see cref="YamlIgnoreCondition.Always"/>, which also stops
    /// reading it. The other conditions affect writing only and win over
    /// <see cref="YamlSerializerOptions.DefaultIgnoreCondition"/>.
    /// </summary>
    public YamlIgnoreCondition Condition { get; set; } = YamlIgnoreCondition.Always;
}

/// <summary>
/// Makes a property required when reading: a mapping without its key fails with
/// <see cref="YamlErrorCodes.MissingRequiredMember"/>. C# <c>required</c> members are required without the attribute.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class YamlRequiredAttribute : Attribute {
}

/// <summary>
/// Selects the constructor that reading uses when a type has several public constructors. Its parameters bind to the
/// properties with the same names, ignoring case.
/// </summary>
[AttributeUsage(AttributeTargets.Constructor, AllowMultiple = false)]
public sealed class YamlConstructorAttribute : Attribute {
}

/// <summary>
/// Registers a derived type of the annotated class or interface for polymorphic serialization. A value of the derived
/// type is written with its discriminator, and reading creates the derived type only for that discriminator: a YAML tag
/// never names a .NET type by itself.
/// </summary>
/// <remarks>
/// The discriminator is written as the local tag <c>!discriminator</c> by default, or as the first mapping entry when
/// <see cref="YamlPolymorphicAttribute.DiscriminatorPropertyName"/> is set on the base type.
/// </remarks>
[AttributeUsage(AttributeTargets.Class | AttributeTargets.Interface, AllowMultiple = true, Inherited = false)]
public sealed class YamlDerivedTypeAttribute : Attribute {
    /// <summary>Creates the attribute.</summary>
    /// <param name="derivedType">A concrete type that derives from or implements the annotated type.</param>
    /// <param name="discriminator">The text that identifies <paramref name="derivedType"/> in YAML.</param>
    public YamlDerivedTypeAttribute(Type derivedType, string discriminator) {
        DerivedType = derivedType ?? throw new ArgumentNullException(nameof(derivedType));
        Discriminator = discriminator ?? throw new ArgumentNullException(nameof(discriminator));
    }

    /// <summary>Gets the derived type.</summary>
    public Type DerivedType { get; }

    /// <summary>Gets the discriminator.</summary>
    public string Discriminator { get; }
}

/// <summary>
/// Configures how a type with <see cref="YamlDerivedTypeAttribute"/> registrations writes its discriminator.
/// </summary>
[AttributeUsage(AttributeTargets.Class | AttributeTargets.Interface, AllowMultiple = false, Inherited = false)]
public sealed class YamlPolymorphicAttribute : Attribute {
    /// <summary>
    /// Gets or sets the mapping key that holds the discriminator, such as <c>kind</c>. When <see langword="null"/>
    /// (the default) the discriminator is a local tag (<c>!cat</c>). The key is written first and may appear anywhere
    /// when reading.
    /// </summary>
    public string? DiscriminatorPropertyName { get; set; }
}

/// <summary>
/// Adds a root type to a <see cref="YamlSerializerContext"/>. The source generator produces metadata for the type and
/// every type it reaches through properties, collections and derived-type registrations.
/// </summary>
[AttributeUsage(AttributeTargets.Class, AllowMultiple = true)]
public sealed class YamlSerializableAttribute : Attribute {
    /// <summary>Creates the attribute.</summary>
    /// <param name="type">The type to generate metadata for.</param>
    public YamlSerializableAttribute(Type type) {
        Type = type ?? throw new ArgumentNullException(nameof(type));
    }

    /// <summary>Gets the type.</summary>
    public Type Type { get; }
}

/// <summary>When a property is left out of the YAML a serializer writes.</summary>
public enum YamlIgnoreCondition : byte {
    /// <summary>The property is always written.</summary>
    Never = 0,
    /// <summary>The property is never written or read. Valid only on <see cref="YamlIgnoreAttribute"/>.</summary>
    Always,
    /// <summary>The property is not written when its value is <see langword="null"/>.</summary>
    WhenWritingNull,
    /// <summary>The property is not written when its value is <see langword="null"/> or the default of a value type.</summary>
    WhenWritingDefault
}

/// <summary>How the serializer treats an object that is reached more than once.</summary>
public enum YamlReferenceHandling : byte {
    /// <summary>
    /// Each reference is written in full, and a cycle fails with <see cref="YamlErrorCodes.ReferenceCycle"/>. An alias in
    /// the input read as a typed value is read again from its anchored node; one read as <see cref="object"/> returns
    /// the value read for its anchor, unless <see cref="YamlSerializerOptions.CopyAliases"/> is set. This is the default.
    /// </summary>
    None = 0,
    /// <summary>
    /// An object or collection reached again is written as an alias to an anchor on its first occurrence, so shared
    /// references and cycles round-trip. Reading maps each alias back to the instance created for its anchor.
    /// </summary>
    Preserve
}
