namespace OfficeIMO.Yaml;

/// <summary>
/// An alias node (<c>*name</c>). The tree keeps aliases as nodes so saving and editing preserve them; resolution is a
/// view: <see cref="Target"/> and the typed reads on <see cref="YamlNode"/> look through to the anchored node.
/// </summary>
public sealed class YamlAlias : YamlNode {
    internal YamlAlias(YamlDocumentPart part, int start, int end, string name, YamlNode target) : base(part, YamlStyle.None, start, end) {
        Name = name;
        Target = target;
    }

    /// <summary>Gets the anchor name the alias refers to, without <c>*</c>.</summary>
    public string Name { get; }

    /// <summary>
    /// Gets the anchored node: the most recent node before the alias in the same document whose anchor is
    /// <see cref="Name"/>. It is never itself an alias.
    /// </summary>
    public YamlNode Target { get; }
}
