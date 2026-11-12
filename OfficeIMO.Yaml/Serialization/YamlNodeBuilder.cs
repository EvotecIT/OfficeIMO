namespace OfficeIMO.Yaml;

/// <summary>
/// Builds detached <see cref="YamlNode"/> trees from <see cref="YamlWriter"/> calls, so the serializer can produce model
/// nodes with the same converters that write text. New string scalars choose their style when they join a document.
/// </summary>
internal sealed class YamlNodeBuilder {
    private struct Frame {
        public YamlNode Collection;
        public YamlNode? PendingKey;
    }

    private Frame[] _frames = new Frame[16];
    private int _depth;
    private string? _anchor;
    private string? _tag;
    private Dictionary<string, YamlNode>? _anchors;

    /// <summary>Gets the finished root node, or null before one is written.</summary>
    public YamlNode? Root { get; private set; }

    public bool AtKey => _depth > 0 && _frames[_depth - 1].Collection is YamlMapping && _frames[_depth - 1].PendingKey is null;

    public void StartCollection(bool mapping, YamlStyle style) {
        YamlNode node = mapping ? new YamlMapping(style) : new YamlSequence(style);
        Add(node);
        if (_depth == _frames.Length) Array.Resize(ref _frames, _depth * 2);
        _frames[_depth++] = new Frame { Collection = node };
    }

    public void EndCollection(bool mapping) {
        if (_depth == 0 || (_frames[_depth - 1].Collection is YamlMapping) != mapping) {
            throw new InvalidOperationException("The innermost open collection is not a " + (mapping ? "mapping." : "sequence."));
        }
        if (_frames[_depth - 1].PendingKey is not null) throw new InvalidOperationException("The last key of the mapping has no value.");
        _frames[--_depth] = default;
    }

    public void Scalar(YamlScalar scalar) => Add(scalar);

    public void Alias(string name) {
        if (_anchors is null || !_anchors.TryGetValue(name, out YamlNode? target)) throw new InvalidOperationException("The alias '*" + name + "' refers to no anchor written before it.");
        Add(new YamlAlias(target));
    }

    public void SetAnchor(string name) => _anchor = name;

    public void SetTag(string tag) => _tag = tag;

    private void Add(YamlNode node) {
        if (_anchor is not null) {
            node.SetAnchor(_anchor);
            (_anchors ??= new Dictionary<string, YamlNode>(StringComparer.Ordinal))[_anchor] = node;
            _anchor = null;
        }
        if (_tag is not null) {
            node.SetTag(_tag);
            _tag = null;
        }
        if (_depth == 0) {
            if (Root is not null) throw new InvalidOperationException("A node tree has one root; it is already written.");
            Root = node;
            return;
        }
        ref Frame top = ref _frames[_depth - 1];
        if (top.Collection is YamlMapping mapping) {
            if (top.PendingKey is null) {
                top.PendingKey = node;
                return;
            }
            mapping.Add(top.PendingKey, node);
            top.PendingKey = null;
            return;
        }
        ((YamlSequence)top.Collection).Add(node);
    }
}
