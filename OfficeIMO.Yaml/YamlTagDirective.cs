namespace OfficeIMO.Yaml;

/// <summary>
/// A <c>%TAG</c> directive of one document: shorthand tags that start with <see cref="Handle"/> expand to
/// <see cref="Prefix"/> followed by the tag suffix.
/// </summary>
public readonly struct YamlTagDirective {
    internal YamlTagDirective(string handle, string prefix) {
        Handle = handle;
        Prefix = prefix;
    }

    /// <summary>Gets the tag handle, such as <c>!e!</c>.</summary>
    public string Handle { get; }

    /// <summary>Gets the prefix the handle expands to, with percent escapes decoded.</summary>
    public string Prefix { get; }
}
