namespace OfficeIMO.Yaml;

/// <summary>
/// An alias node (<c>*name</c>). The tree keeps aliases as nodes so saving and editing preserve them; resolution is a
/// view: <see cref="Target"/> and the typed reads on <see cref="YamlNode"/> look through to the anchored node.
/// </summary>
public sealed class YamlAlias : YamlNode {
    private readonly string _name;
    private YamlNode? _target;
    /// <summary>The anchored node's record for a loaded alias, found on first use.</summary>
    private readonly int _targetRecord = -1;

    internal YamlAlias(YamlDocumentPart part, int start, int end, string name, int targetRecord) : base(part, YamlStyle.None, start, end) {
        _name = name;
        _targetRecord = targetRecord;
    }

    /// <summary>
    /// Creates an alias to an anchored node, for example to reuse shared settings with a <c>&lt;&lt;</c> merge key. The
    /// alias must come after its anchor in the document when it is saved.
    /// </summary>
    /// <exception cref="ArgumentException"><paramref name="anchored"/> has no anchor, or is itself an alias.</exception>
    public YamlAlias(YamlNode anchored) : base(YamlStyle.None) {
        if (anchored is null) throw new ArgumentNullException(nameof(anchored));
        if (anchored is YamlAlias) throw new ArgumentException("An alias refers to an anchored node, not to another alias.", nameof(anchored));
        _name = anchored.Anchor ?? throw new ArgumentException("The node has no anchor; set one with SetAnchor first.", nameof(anchored));
        _target = anchored;
    }

    /// <summary>Gets the anchor name the alias refers to, without <c>*</c>. It follows renames of the anchor.</summary>
    public string Name => Target.Anchor ?? _name;

    /// <summary>
    /// Gets the anchored node: the most recent node before the alias in the same document whose anchor is
    /// <see cref="Name"/>. It is never itself an alias.
    /// </summary>
    public YamlNode Target => Volatile.Read(ref _target) ?? ResolveTarget();

    private YamlNode ResolveTarget() {
        YamlNode target = _part!._document._table!.NodeAt(_part, _targetRecord);
        Volatile.Write(ref _target, target);
        return target;
    }
}
