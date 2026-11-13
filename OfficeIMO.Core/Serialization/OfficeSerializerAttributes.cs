namespace OfficeIMO;

// The object-mapping attributes the OfficeIMO serializers share. YamlSerializer and TomlSerializer read the same
// attributes, and their source generators look them up by these full names, so one model type maps to both formats.

/// <summary>
/// Sets the key a property is written and read under. It wins over the serializer options' property naming policy.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class OfficePropertyNameAttribute : Attribute {
    /// <summary>Creates the attribute.</summary>
    /// <param name="name">The key, used as written.</param>
    public OfficePropertyNameAttribute(string name) {
        Name = name ?? throw new ArgumentNullException(nameof(name));
    }

    /// <summary>Gets the key.</summary>
    public string Name { get; }
}

/// <summary>
/// Leaves a property out of serialization: always by default, or only when its value is <see langword="null"/> or the
/// type's default when <see cref="Condition"/> says so.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class OfficeIgnoreAttribute : Attribute {
    /// <summary>
    /// Gets or sets when the property is left out. Defaults to <see cref="OfficeIgnoreCondition.Always"/>, which also stops
    /// reading it. The other conditions affect writing only and win over the options' default ignore condition.
    /// </summary>
    public OfficeIgnoreCondition Condition { get; set; } = OfficeIgnoreCondition.Always;
}

/// <summary>
/// Makes a property required when reading: a mapping or table without its key fails with the format's
/// <c>MissingRequiredMember</c> error code. C# <c>required</c> members are required without the attribute.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class OfficeRequiredAttribute : Attribute {
}

/// <summary>
/// Selects the constructor that reading uses when a type has several public constructors. Its parameters bind to the
/// properties with the same names, ignoring case.
/// </summary>
[AttributeUsage(AttributeTargets.Constructor, AllowMultiple = false)]
public sealed class OfficeConstructorAttribute : Attribute {
}

/// <summary>
/// Registers a derived type of the annotated class or interface for polymorphic serialization. A value of the derived
/// type is written with its discriminator, and reading creates the derived type only for that discriminator: the input
/// never names a .NET type by itself.
/// </summary>
/// <remarks>
/// YAML writes the discriminator as the local tag <c>!discriminator</c> by default. TOML has no tags, so it writes the
/// discriminator as a key, <c>$type</c> by default. <see cref="OfficePolymorphicAttribute.DiscriminatorPropertyName"/>
/// on the base type makes both formats use a key with that name.
/// </remarks>
[AttributeUsage(AttributeTargets.Class | AttributeTargets.Interface, AllowMultiple = true, Inherited = false)]
public sealed class OfficeDerivedTypeAttribute : Attribute {
    /// <summary>Creates the attribute.</summary>
    /// <param name="derivedType">A concrete type that derives from or implements the annotated type.</param>
    /// <param name="discriminator">The text that identifies <paramref name="derivedType"/> in the serialized form.</param>
    public OfficeDerivedTypeAttribute(Type derivedType, string discriminator) {
        DerivedType = derivedType ?? throw new ArgumentNullException(nameof(derivedType));
        Discriminator = discriminator ?? throw new ArgumentNullException(nameof(discriminator));
    }

    /// <summary>Gets the derived type.</summary>
    public Type DerivedType { get; }

    /// <summary>Gets the discriminator.</summary>
    public string Discriminator { get; }
}

/// <summary>
/// Configures how a type with <see cref="OfficeDerivedTypeAttribute"/> registrations writes its discriminator.
/// </summary>
[AttributeUsage(AttributeTargets.Class | AttributeTargets.Interface, AllowMultiple = false, Inherited = false)]
public sealed class OfficePolymorphicAttribute : Attribute {
    /// <summary>
    /// Gets or sets the key that holds the discriminator, such as <c>kind</c>. The key is written first and may appear
    /// anywhere when reading. When <see langword="null"/> (the default), YAML uses a local tag (<c>!cat</c>) and TOML
    /// the key <c>$type</c>.
    /// </summary>
    public string? DiscriminatorPropertyName { get; set; }
}

/// <summary>When a property is left out of the output a serializer writes.</summary>
public enum OfficeIgnoreCondition : byte {
    /// <summary>The property is always written.</summary>
    Never = 0,
    /// <summary>The property is never written or read. Valid only on <see cref="OfficeIgnoreAttribute"/>.</summary>
    Always,
    /// <summary>The property is not written when its value is <see langword="null"/>.</summary>
    WhenWritingNull,
    /// <summary>The property is not written when its value is <see langword="null"/> or the default of a value type.</summary>
    WhenWritingDefault
}
