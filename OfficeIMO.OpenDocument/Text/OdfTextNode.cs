namespace OfficeIMO.OpenDocument;

/// <summary>Kind of native inline text syntax.</summary>
public enum OdfTextNodeKind {
    /// <summary>Text, including native whitespace.</summary>
    Text,
    /// <summary>Styled run.</summary>
    Run,
    /// <summary>Hyperlink.</summary>
    Hyperlink,
    /// <summary>Basic page, date, or time field.</summary>
    Field,
    /// <summary>Preserved native syntax without a typed editor.</summary>
    Other
}

/// <summary>An ordered snapshot of native inline syntax with live wrappers for targeted edits.</summary>
public sealed class OdfTextNode {
    private readonly string? _text;
    private OdfTextNode(OdfTextNodeKind kind, string? text, XElement? element = null,
        OdfTextRun? run = null, OdfTextHyperlink? hyperlink = null, OdfTextField? field = null, IReadOnlyList<OdfTextNode>? children = null) {
        Kind = kind; _text = text; QualifiedName = element?.Name.ToString(); Run = run; Hyperlink = hyperlink; Field = field;
        Children = children ?? Array.Empty<OdfTextNode>();
    }
    /// <summary>Native syntax kind.</summary>
    public OdfTextNodeKind Kind { get; }
    /// <summary>Decoded text at the time this view was read.</summary>
    public string Text => _text ?? OdfTextCodec.JoinBounded(Children.Select(c => c.Text), string.Empty);
    /// <summary>Expanded XML name for an element node.</summary>
    public string? QualifiedName { get; }
    /// <summary>Live run wrapper, when present.</summary>
    public OdfTextRun? Run { get; }
    /// <summary>Live hyperlink wrapper, when present.</summary>
    public OdfTextHyperlink? Hyperlink { get; }
    /// <summary>Live field wrapper, when present.</summary>
    public OdfTextField? Field { get; }
    /// <summary>Ordered children of a run or hyperlink.</summary>
    public IReadOnlyList<OdfTextNode> Children { get; }

    internal static IReadOnlyList<OdfTextNode> Read(OdfDocument document, XElement parent, XElement graphic) {
        int remaining = OdfTextCodec.MaximumDecodedCharacters;
        return ReadChildren(document, parent, graphic, OdfTextCodec.Snapshot(parent), ref remaining);
    }
    private static IReadOnlyList<OdfTextNode> ReadChildren(OdfDocument document, XElement parent, XElement graphic, OdfTextCodec.TextSnapshot decoded, ref int remaining) {
        var result = new List<OdfTextNode>(); var plain = new List<XNode>();
        void Flush(ref int budget) {
            if (plain.Count == 0) return;
            string text = decoded.ReadNodes(plain, ref budget);
            if (text.Length > 0) result.Add(new OdfTextNode(OdfTextNodeKind.Text, text));
            plain.Clear();
        }
        foreach (XNode node in parent.Nodes()) {
            if (node is XText) { plain.Add(node); continue; }
            if (!(node is XElement element)) continue;
            if (element.Name == OdfNamespaces.Text + "s" || element.Name == OdfNamespaces.Text + "tab" || element.Name == OdfNamespaces.Text + "line-break") { plain.Add(element); continue; }
            Flush(ref remaining);
            if (element.Name == OdfNamespaces.Text + "span") result.Add(new OdfTextNode(OdfTextNodeKind.Run, null, element,
                run: new OdfTextRun(document, element, graphic), children: ReadChildren(document, element, graphic, decoded, ref remaining)));
            else if (element.Name == OdfNamespaces.Text + "a") result.Add(new OdfTextNode(OdfTextNodeKind.Hyperlink, null, element,
                hyperlink: new OdfTextHyperlink(document, element, graphic), children: ReadChildren(document, element, graphic, decoded, ref remaining)));
            else if (OdfTextField.IsField(element.Name)) result.Add(new OdfTextNode(OdfTextNodeKind.Field, decoded.Read(element, ref remaining), element, field: new OdfTextField(document, element)));
            else result.Add(new OdfTextNode(OdfTextNodeKind.Other, OdfTextCodec.IsNonVisibleTextElement(element) ? string.Empty : decoded.Read(element, ref remaining), element));
        }
        Flush(ref remaining); return result;
    }
}
