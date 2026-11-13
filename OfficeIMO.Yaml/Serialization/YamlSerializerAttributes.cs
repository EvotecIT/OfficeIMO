namespace OfficeIMO.Yaml;

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
