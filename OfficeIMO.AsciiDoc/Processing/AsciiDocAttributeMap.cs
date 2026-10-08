using System.Collections;

namespace OfficeIMO.AsciiDoc;

// Immutable AVL nodes share unchanged branches between source-ordered attribute snapshots.
internal sealed class AsciiDocAttributeMap : IReadOnlyDictionary<string, string> {
    private readonly Node? _root;

    internal AsciiDocAttributeMap(IReadOnlyDictionary<string, string>? values = null) {
        if (values == null) return;
        foreach (KeyValuePair<string, string> pair in values)
            _root = Set(_root, pair.Key, pair.Value);
    }

    private AsciiDocAttributeMap(Node? root) => _root = root;

    internal AsciiDocAttributeMap Set(string key, string value) {
        if (key == null) throw new ArgumentNullException(nameof(key));
        return new AsciiDocAttributeMap(Set(_root, key, value));
    }

    internal AsciiDocAttributeMap Remove(string key) {
        if (key == null) throw new ArgumentNullException(nameof(key));
        return new AsciiDocAttributeMap(Remove(_root, key));
    }

    public int Count => _root?.Count ?? 0;
    public IEnumerable<string> Keys => this.Select(static pair => pair.Key);
    public IEnumerable<string> Values => this.Select(static pair => pair.Value);
    public string this[string key] => TryGetValue(key, out string? value)
        ? value : throw new KeyNotFoundException("The attribute does not exist: " + key);

    public bool ContainsKey(string key) => TryGetValue(key, out _);

    public bool TryGetValue(string key, out string value) {
        if (key == null) throw new ArgumentNullException(nameof(key));
        Node? node = _root;
        while (node != null) {
            int order = StringComparer.OrdinalIgnoreCase.Compare(key, node.Key);
            if (order == 0) { value = node.Value; return true; }
            node = order < 0 ? node.Left : node.Right;
        }
        value = null!;
        return false;
    }

    public IEnumerator<KeyValuePair<string, string>> GetEnumerator() => Enumerate(_root).GetEnumerator();
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    private static IEnumerable<KeyValuePair<string, string>> Enumerate(Node? node) {
        if (node == null) yield break;
        foreach (KeyValuePair<string, string> pair in Enumerate(node.Left)) yield return pair;
        yield return new KeyValuePair<string, string>(node.Key, node.Value);
        foreach (KeyValuePair<string, string> pair in Enumerate(node.Right)) yield return pair;
    }

    private static Node Set(Node? node, string key, string value) {
        if (node == null) return new Node(key, value, null, null);
        int order = StringComparer.OrdinalIgnoreCase.Compare(key, node.Key);
        if (order == 0) return new Node(node.Key, value, node.Left, node.Right);
        return order < 0
            ? Balance(new Node(node.Key, node.Value, Set(node.Left, key, value), node.Right))
            : Balance(new Node(node.Key, node.Value, node.Left, Set(node.Right, key, value)));
    }

    private static Node? Remove(Node? node, string key) {
        if (node == null) return null;
        int order = StringComparer.OrdinalIgnoreCase.Compare(key, node.Key);
        if (order < 0) {
            Node? left = Remove(node.Left, key);
            return ReferenceEquals(left, node.Left) ? node : Balance(new Node(node.Key, node.Value, left, node.Right));
        }
        if (order > 0) {
            Node? right = Remove(node.Right, key);
            return ReferenceEquals(right, node.Right) ? node : Balance(new Node(node.Key, node.Value, node.Left, right));
        }
        if (node.Left == null) return node.Right;
        if (node.Right == null) return node.Left;
        Node successor = Minimum(node.Right);
        return Balance(new Node(successor.Key, successor.Value, node.Left, RemoveMinimum(node.Right)));
    }

    private static Node Minimum(Node node) => node.Left == null ? node : Minimum(node.Left);

    private static Node? RemoveMinimum(Node node) => node.Left == null ? node.Right
        : Balance(new Node(node.Key, node.Value, RemoveMinimum(node.Left), node.Right));

    private static Node Balance(Node node) {
        int difference = Height(node.Left) - Height(node.Right);
        if (difference > 1) {
            Node left = node.Left!;
            if (Height(left.Left) < Height(left.Right))
                node = new Node(node.Key, node.Value, RotateLeft(left), node.Right);
            return RotateRight(node);
        }
        if (difference < -1) {
            Node right = node.Right!;
            if (Height(right.Right) < Height(right.Left))
                node = new Node(node.Key, node.Value, node.Left, RotateRight(right));
            return RotateLeft(node);
        }
        return node;
    }

    private static Node RotateLeft(Node node) {
        Node right = node.Right!;
        return new Node(right.Key, right.Value,
            new Node(node.Key, node.Value, node.Left, right.Left), right.Right);
    }

    private static Node RotateRight(Node node) {
        Node left = node.Left!;
        return new Node(left.Key, left.Value, left.Left,
            new Node(node.Key, node.Value, left.Right, node.Right));
    }

    private static int Height(Node? node) => node?.Height ?? 0;

    private sealed class Node {
        internal Node(string key, string value, Node? left, Node? right) {
            Key = key; Value = value; Left = left; Right = right;
            Height = 1 + Math.Max(AsciiDocAttributeMap.Height(left), AsciiDocAttributeMap.Height(right));
            Count = 1 + (left?.Count ?? 0) + (right?.Count ?? 0);
        }
        internal string Key { get; }
        internal string Value { get; }
        internal Node? Left { get; }
        internal Node? Right { get; }
        internal int Height { get; }
        internal int Count { get; }
    }
}
