using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Word.OpenXml;

/// <summary>
/// The style facts built-in style resolution needs, read from the styles part of an owned package: each identifier's
/// first definition (name, type and the styles it depends on) and the identifier Word uses for each built-in name. It
/// answers as <c>WordDocument.BuiltInStyleIndex</c> does over the SDK's styles DOM. The owned layer doesn't add style
/// definitions; a style the document lacks makes the caller open the SDK, which adds it from the library.
/// </summary>
internal sealed class WordOwnedStyles {
    private const int W = OfficeXmlNamespaces.WordprocessingML;
    private const string StylesRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "styles";

    private readonly Dictionary<string, Definition> _definitions = new(StringComparer.Ordinal);
    private readonly Dictionary<string, string> _idByBuiltInId = new(StringComparer.Ordinal);

    private readonly struct Definition {
        internal Definition(string? name, string? type, string? basedOn, string? linked, string? next) {
            Name = name;
            Type = type;
            BasedOn = basedOn;
            Linked = linked;
            Next = next;
        }

        internal string? Name { get; }
        internal string? Type { get; }
        internal string? BasedOn { get; }
        internal string? Linked { get; }
        internal string? Next { get; }
    }

    private WordOwnedStyles(OfficeXmlTree? styles) {
        if (styles == null) return;
        for (int style = styles.FirstChildOf(styles.Root); style != OfficeXmlTree.NoNode; style = styles.NextSiblingOf(style)) {
            if (!styles.IsElement(style, W, "style")) continue;
            // As the SDK path's index: a style needs an identifier, and only an identifier's first definition counts.
            string? id = styles.GetAttribute(style, W, "styleId");
            if (string.IsNullOrEmpty(id) || _definitions.ContainsKey(id!)) continue;
            string? name = ChildValue(styles, style, "name");
            string? type = styles.GetAttribute(style, W, "type");
            _definitions.Add(id!, new Definition(name, type, ChildValue(styles, style, "basedOn"),
                ChildValue(styles, style, "link"), ChildValue(styles, style, "next")));
            // Word keeps the first definition of a name; a style of another type never stands in.
            if (WordBuiltInStyleNames.GetId(name) is { } builtInId &&
                WordBuiltInStyleNames.Classify(id, name, type) == builtInId && !_idByBuiltInId.ContainsKey(builtInId)) {
                _idByBuiltInId.Add(builtInId, id!);
            }
        }
    }

    /// <summary>Reads the styles part of <paramref name="model"/>; a package without one has no definitions.</summary>
    internal static WordOwnedStyles Read(WordPackageModel model) {
        string? part = WordOwnedSave.RelatedParts(model, StylesRelationshipType).FirstOrDefault();
        return new WordOwnedStyles(part == null ? null : model.Package.GetXml(part));
    }

    private static string? ChildValue(OfficeXmlTree tree, int node, string localName) {
        int child = tree.FindChild(node, W, localName);
        return child == OfficeXmlTree.NoNode ? null : tree.GetAttribute(child, W, "val");
    }

    /// <summary>
    /// The English identifier of the built-in style <paramref name="styleId"/> is, through its definition's name, or the
    /// identifier itself when it is an English built-in identifier without a definition; as <c>GetBuiltInStyleId</c>.
    /// </summary>
    internal string? GetBuiltInStyleId(string? styleId) {
        if (string.IsNullOrEmpty(styleId)) return null;
        if (_definitions.TryGetValue(styleId!, out Definition definition)) {
            return WordBuiltInStyleNames.Classify(styleId, definition.Name, definition.Type);
        }
        return WordBuiltInStyleNames.IsBuiltInId(styleId) ? styleId : null;
    }

    /// <summary>The document's identifier for a built-in style, as <c>FindBuiltInStyleId</c>.</summary>
    internal string? FindBuiltInStyleId(string builtInStyleId) {
        if (_idByBuiltInId.TryGetValue(builtInStyleId, out string? styleId)) return styleId;
        // A definition under the English identifier counts only when it has the built-in style's type.
        return _definitions.TryGetValue(builtInStyleId, out Definition definition) &&
            WordBuiltInStyleNames.Classify(builtInStyleId, null, definition.Type) != null ? builtInStyleId : null;
    }

    /// <summary>
    /// Whether the document defines <paramref name="styleId"/> and every style it is based on, linked to or followed by,
    /// so applying it adds no definition.
    /// </summary>
    internal bool DefinesWithDependencies(string styleId) {
        var visited = new HashSet<string>(StringComparer.Ordinal);
        var pending = new Stack<string>();
        pending.Push(styleId);
        while (pending.Count > 0) {
            string id = pending.Pop();
            if (!visited.Add(id)) continue;
            if (!_definitions.TryGetValue(id, out Definition definition)) return false;
            if (!string.IsNullOrEmpty(definition.BasedOn)) pending.Push(definition.BasedOn!);
            if (!string.IsNullOrEmpty(definition.Linked)) pending.Push(definition.Linked!);
            if (!string.IsNullOrEmpty(definition.Next)) pending.Push(definition.Next!);
        }
        return true;
    }
}
