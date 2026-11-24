using System.Globalization;

namespace OfficeIMO.Core.Internal.OpenXml;

/// <summary>Reads and writes attribute values with XML Schema lexical forms, for the generated element views.</summary>
/// <remarks>Values that don't parse read as null rather than throwing, so a view can still edit a part another
/// producer wrote loosely; the attribute itself is kept as written until it is set.</remarks>
internal static class OfficeXmlViewValues {
    internal static bool? ReadBoolean(OfficeXmlTree tree, int node, int namespaceId, string localName) {
        if (!tree.TryGetAttribute(node, namespaceId, localName, out string value)) return null;
        return Trim(value) switch {
            "1" or "true" => true,
            "0" or "false" => false,
            _ => null
        };
    }

    internal static int? ReadInt32(OfficeXmlTree tree, int node, int namespaceId, string localName) =>
        tree.TryGetAttribute(node, namespaceId, localName, out string value)
        && int.TryParse(Trim(value), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out int parsed) ? parsed : null;

    internal static uint? ReadUInt32(OfficeXmlTree tree, int node, int namespaceId, string localName) =>
        tree.TryGetAttribute(node, namespaceId, localName, out string value)
        && uint.TryParse(Trim(value), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out uint parsed) ? parsed : null;

    internal static long? ReadInt64(OfficeXmlTree tree, int node, int namespaceId, string localName) =>
        tree.TryGetAttribute(node, namespaceId, localName, out string value)
        && long.TryParse(Trim(value), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out long parsed) ? parsed : null;

    internal static double? ReadDouble(OfficeXmlTree tree, int node, int namespaceId, string localName) =>
        tree.TryGetAttribute(node, namespaceId, localName, out string value)
        && double.TryParse(Trim(value), NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) ? parsed : null;

    internal static void Write(OfficeXmlTree tree, int node, int namespaceId, string localName, string? value) {
        if (value == null) tree.RemoveAttribute(node, namespaceId, localName);
        else tree.SetAttribute(node, namespaceId, localName, value);
    }

    internal static void Write(OfficeXmlTree tree, int node, int namespaceId, string localName, bool? value) =>
        Write(tree, node, namespaceId, localName, value.HasValue ? value.Value ? "1" : "0" : null);

    internal static void Write(OfficeXmlTree tree, int node, int namespaceId, string localName, int? value) =>
        Write(tree, node, namespaceId, localName, value?.ToString(CultureInfo.InvariantCulture));

    internal static void Write(OfficeXmlTree tree, int node, int namespaceId, string localName, uint? value) =>
        Write(tree, node, namespaceId, localName, value?.ToString(CultureInfo.InvariantCulture));

    internal static void Write(OfficeXmlTree tree, int node, int namespaceId, string localName, long? value) =>
        Write(tree, node, namespaceId, localName, value?.ToString(CultureInfo.InvariantCulture));

    internal static void Write(OfficeXmlTree tree, int node, int namespaceId, string localName, double? value) =>
        Write(tree, node, namespaceId, localName, value?.ToString("R", CultureInfo.InvariantCulture));

    /// <summary>The text of a child element with simple content, or null when there is no such child.</summary>
    internal static string? ReadChildText(OfficeXmlTree tree, int node, int namespaceId, string localName) {
        int child = tree.FindChild(node, namespaceId, localName);
        return child == OfficeXmlTree.NoNode ? null : tree.GetInnerText(child);
    }

    /// <summary>Sets the text of a child element, creating it in sequence order; null removes the child.</summary>
    internal static void WriteChildText(OfficeXmlTree tree, int node, int namespaceId, string localName, string? value, OfficeXmlChildOrder order) {
        int child = tree.FindChild(node, namespaceId, localName);
        if (value == null) {
            if (child != OfficeXmlTree.NoNode) tree.Remove(child);
            return;
        }
        if (child == OfficeXmlTree.NoNode) {
            child = tree.CreateElement(namespaceId, localName);
            tree.InsertOrdered(node, child, order);
        }
        tree.SetInnerText(child, value);
    }

    // XML Schema collapses whitespace in boolean and numeric values.
    private static string Trim(string value) => value.Length > 0 && (value[0] <= ' ' || value[value.Length - 1] <= ' ') ? value.Trim() : value;
}

/// <summary>The child elements of one name, enumerated without allocating.</summary>
internal readonly struct OfficeXmlElementList {
    private readonly OfficeXmlTree _tree;
    private readonly int _parent;
    private readonly int _namespaceId;
    private readonly string _localName;

    internal OfficeXmlElementList(OfficeXmlTree tree, int parent, int namespaceId, string localName) {
        _tree = tree;
        _parent = parent;
        _namespaceId = namespaceId;
        _localName = localName;
    }

    public Enumerator GetEnumerator() => new(_tree, _parent, _namespaceId, _localName);

    internal int Count {
        get {
            int count = 0;
            foreach (int _ in this) count++;
            return count;
        }
    }

    internal struct Enumerator {
        private readonly OfficeXmlTree _tree;
        private readonly int _parent;
        private readonly int _namespaceId;
        private readonly string _localName;
        private int _current;
        private bool _started;

        internal Enumerator(OfficeXmlTree tree, int parent, int namespaceId, string localName) {
            _tree = tree;
            _parent = parent;
            _namespaceId = namespaceId;
            _localName = localName;
            _current = OfficeXmlTree.NoNode;
            _started = false;
        }

        public int Current => _current;

        public bool MoveNext() {
            if (_tree == null) return false;
            if (!_started) {
                _started = true;
                _current = _tree.FindChild(_parent, _namespaceId, _localName);
            } else if (_current != OfficeXmlTree.NoNode) {
                _current = _tree.FindNextSibling(_current, _namespaceId, _localName);
            }
            return _current != OfficeXmlTree.NoNode;
        }
    }
}
