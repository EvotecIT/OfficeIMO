namespace OfficeIMO.Toml;

/// <summary>
/// Adds a root type to a <see cref="TomlSerializerContext"/>. The source generator produces metadata for the type and
/// every type it reaches through properties, collections and derived-type registrations.
/// </summary>
[AttributeUsage(AttributeTargets.Class, AllowMultiple = true)]
public sealed class TomlSerializableAttribute : Attribute {
    /// <summary>Creates the attribute.</summary>
    /// <param name="type">The type to generate metadata for.</param>
    public TomlSerializableAttribute(Type type) {
        Type = type ?? throw new ArgumentNullException(nameof(type));
    }

    /// <summary>Gets the type.</summary>
    public Type Type { get; }
}

/// <summary>
/// Writes a table as an inline table on its key's line (<c>serde = { version = "1", features = ["derive"] }</c>)
/// instead of a <c>[table]</c> section. On a class or struct it applies to every value of the type, as Cargo
/// dependencies are written; on a property it applies to that property's value, and to the items of a collection or
/// the values of a dictionary. Reading accepts both forms whatever the attribute says.
/// </summary>
[AttributeUsage(AttributeTargets.Class | AttributeTargets.Struct | AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
public sealed class TomlInlineTableAttribute : Attribute {
}

/// <summary>What <see cref="TomlSerializer"/> does with a <see langword="null"/> property or dictionary value, which TOML cannot hold.</summary>
public enum TomlNullHandling : byte {
    /// <summary>The key is left out, so reading it back leaves the property at its default. This is the default.</summary>
    Omit = 0,
    /// <summary>Writing fails with <see cref="TomlErrorCodes.UnrepresentableValue"/>.</summary>
    Error
}
