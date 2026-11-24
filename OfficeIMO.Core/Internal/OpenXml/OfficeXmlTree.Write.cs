namespace OfficeIMO.Core.Internal.OpenXml;

internal sealed partial class OfficeXmlTree {
    private static Span _noSpan;
    private static readonly byte[] Declaration = System.Text.Encoding.ASCII.GetBytes("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");

    /// <summary>
    /// Writes the part as UTF-8. Unchanged subtrees, start tags, text, comments and instructions of a parsed part are
    /// copied as written; a part that was not changed is written back byte for byte.
    /// </summary>
    internal void WriteTo(OfficeXmlUtf8Writer writer) {
        ArgumentNullException.ThrowIfNull(writer);
        if (Root == NoNode) throw new InvalidOperationException("The part has no root element.");
        writer.ResetScope();
        if (_source == null) {
            writer.WriteRaw(Declaration, 0, Declaration.Length);
        } else if (_encodingStart >= 0) {
            // A UTF-16 part is written as UTF-8, so its declaration has to say so.
            writer.WriteRaw(_source, 0, _encodingStart);
            writer.WriteAscii("UTF-8");
            writer.WriteRaw(_source, _encodingEnd, _prologEnd - _encodingEnd);
        } else {
            writer.WriteRaw(_source, 0, _prologEnd);
        }
        WriteNode(writer, Root);
        if (_source != null) writer.WriteRaw(_source, _epilogStart, _sourceLength - _epilogStart);
    }

    private void WriteNode(OfficeXmlUtf8Writer writer, int node) {
        ref Node record = ref _nodes[node];
        bool sourced = node < _sourceNodeCount;
        ref Span span = ref sourced ? ref _spans[node] : ref _noSpan;
        if (record.Kind == OfficeXmlTreeNodeKind.Text) {
            if (sourced && (record.Flags & ContentChanged) == 0 && !FollowsOtherText(node)) writer.WriteRaw(_source!, span.StartTagStart, span.StartTagEnd - span.StartTagStart);
            else writer.WriteEscaped(GetText(node), attribute: false);
            return;
        }
        if (sourced && (record.Flags & SubtreeChanged) == 0) {
            writer.WriteRaw(_source!, span.StartTagStart, span.EndTagEnd - span.StartTagStart);
            return;
        }

        int scope = writer.ScopeMark;
        bool selfClosing = (record.Flags & SelfClosing) != 0;
        bool hasChildren = record.FirstChild != NoNode;
        string prefix;
        if (sourced && (record.Flags & StartTagChanged) == 0) {
            PushDeclarations(writer, node);
            prefix = Namespaces.GetName(record.Name).Prefix!;
            if (selfClosing && hasChildren) {
                // "<a/>" gains content: drop the "/" and close the tag.
                writer.WriteRaw(_source!, span.StartTagStart, span.StartTagEnd - 2 - span.StartTagStart);
                writer.WriteByte((byte)'>');
            } else {
                writer.WriteRaw(_source!, span.StartTagStart, span.StartTagEnd - span.StartTagStart);
                if (selfClosing) {
                    writer.PopScope(scope);
                    return;
                }
            }
        } else {
            prefix = WriteStartTag(writer, node);
            // An element written with an end tag keeps it, even when it has no content now.
            bool keepsEndTag = sourced && !selfClosing;
            if (!hasChildren && !keepsEndTag) {
                writer.WriteByte((byte)'/');
                writer.WriteByte((byte)'>');
                writer.PopScope(scope);
                return;
            }
            writer.WriteByte((byte)'>');
        }

        for (int child = record.FirstChild; child != NoNode; child = _nodes[child].Next) {
            if (sourced && child < _sourceNodeCount && _spans[child].GapStart >= 0) {
                ref Span childSpan = ref _spans[child];
                writer.WriteRaw(_source!, childSpan.GapStart, childSpan.StartTagStart - childSpan.GapStart);
            }
            WriteNode(writer, child);
        }

        if (sourced && !selfClosing) {
            writer.WriteRaw(_source!, span.TrailingGapStart, span.EndTagStart - span.TrailingGapStart);
            writer.WriteRaw(_source!, span.EndTagStart, span.EndTagEnd - span.EndTagStart);
        } else {
            writer.WriteByte((byte)'<');
            writer.WriteByte((byte)'/');
            writer.WriteQualifiedName(prefix, Namespaces.GetLocalNameUtf8(record.Name));
            writer.WriteByte((byte)'>');
        }
        writer.PopScope(scope);
    }

    // Text written as it was is safe next to the markup it was written beside. Next to other text, which an edit
    // brought there, two raw pieces could join into "]]>", so the second is escaped.
    private bool FollowsOtherText(int node) {
        int previous = _nodes[node].Previous;
        if (previous == NoNode || _nodes[previous].Kind != OfficeXmlTreeNodeKind.Text) return false;
        return previous >= _sourceNodeCount || (_nodes[previous].Flags & ContentChanged) != 0 || _spans[previous].StartTagEnd != _spans[node].StartTagStart;
    }

    // Writes "<name attributes" and returns the element's prefix.
    private string WriteStartTag(OfficeXmlUtf8Writer writer, int node) {
        PushDeclarations(writer, node);
        List<CarriedBinding>? carried = PushCarriedBindings(writer, node);
        ref Node record = ref _nodes[node];
        string? declare = null;
        OfficeXmlName name = Namespaces.GetName(record.Name);
        string prefix = name.Prefix ?? ResolvePrefix(writer, name.NamespaceId, name.UriId, preferDefault: true, out declare);
        if (name.Prefix != null) declare = RebindPrefix(writer, prefix, name.NamespaceId, name.UriId);
        writer.WriteByte((byte)'<');
        writer.WriteQualifiedName(prefix, Namespaces.GetLocalNameUtf8(record.Name));
        if (declare != null) WriteDeclaration(writer, prefix, declare);
        if (carried != null) {
            foreach (CarriedBinding binding in carried) {
                WriteDeclaration(writer, binding.Prefix, binding.NamespaceId == OfficeXmlNamespaces.None ? string.Empty : Namespaces.GetUriById(binding.UriId));
            }
        }

        int end = record.AttrFirst + record.AttrCount;
        for (int index = record.AttrFirst; index < end; index++) {
            ref Attr attribute = ref _attrs[index];
            OfficeXmlName attributeName = Namespaces.GetName(attribute.Name);
            writer.WriteByte((byte)' ');
            if (!attribute.Changed && attribute.ValueStart >= 0) {
                if (attributeName.Prefix!.Length > 0 && attributeName.NamespaceId != OfficeXmlNamespaces.Xmlns
                    && RebindPrefix(writer, attributeName.Prefix, attributeName.NamespaceId, attributeName.UriId) is { } attributeUri) {
                    WriteDeclaration(writer, attributeName.Prefix, attributeUri);
                    writer.WriteByte((byte)' ');
                }
                writer.WriteQualifiedName(attributeName.Prefix!, Namespaces.GetLocalNameUtf8(attribute.Name));
                writer.WriteByte((byte)'=');
                writer.WriteByte(attribute.Quote);
                writer.WriteRaw(_source!, attribute.ValueStart, attribute.ValueLength);
                writer.WriteByte(attribute.Quote);
                continue;
            }
            string? attributeDeclare = null;
            string attributePrefix;
            if (attributeName.Prefix == null) {
                attributePrefix = ResolvePrefix(writer, attributeName.NamespaceId, attributeName.UriId, preferDefault: false, out attributeDeclare);
            } else {
                // A parsed attribute keeps its prefix when its value changes; a moved element may have to declare it again.
                attributePrefix = attributeName.Prefix;
                if (attributePrefix.Length > 0 && attributeName.NamespaceId != OfficeXmlNamespaces.Xmlns) {
                    attributeDeclare = RebindPrefix(writer, attributePrefix, attributeName.NamespaceId, attributeName.UriId);
                }
            }
            if (attributeDeclare != null) {
                WriteDeclaration(writer, attributePrefix, attributeDeclare);
                writer.WriteByte((byte)' ');
            }
            writer.WriteQualifiedName(attributePrefix, Namespaces.GetLocalNameUtf8(attribute.Name));
            writer.WriteByte((byte)'=');
            writer.WriteByte((byte)'"');
            writer.WriteEscaped(attribute.Value ?? string.Empty, attribute: true);
            writer.WriteByte((byte)'"');
        }
        return prefix;
    }

    // Bindings the element keeps from where it was written and the current scope binds differently; they are pushed
    // here and declared on the start tag.
    private List<CarriedBinding>? PushCarriedBindings(OfficeXmlUtf8Writer writer, int node) {
        List<CarriedBinding>? bindings = CarriedBindingsOf(node);
        if (bindings == null) return null;
        List<CarriedBinding>? declared = null;
        foreach (CarriedBinding binding in bindings) {
            if (FindAttribute(node, OfficeXmlNamespaces.Xmlns, binding.Prefix.Length == 0 ? "xmlns" : binding.Prefix) >= 0) continue;
            bool bound = writer.TryGetBinding(binding.Prefix, out int current);
            if (bound ? current == binding.UriId : binding.NamespaceId == OfficeXmlNamespaces.None) continue;
            writer.PushScope(binding.Prefix, binding.NamespaceId, binding.UriId);
            (declared ??= new List<CarriedBinding>()).Add(binding);
        }
        return declared;
    }

    private void PushDeclarations(OfficeXmlUtf8Writer writer, int node) {
        ref Node record = ref _nodes[node];
        int end = record.AttrFirst + record.AttrCount;
        for (int index = record.AttrFirst; index < end; index++) {
            ref Attr attribute = ref _attrs[index];
            OfficeXmlName name = Namespaces.GetName(attribute.Name);
            if (name.NamespaceId != OfficeXmlNamespaces.Xmlns) continue;
            string prefix = name.Prefix!.Length == 0 ? string.Empty : name.LocalName;
            writer.PushScope(prefix, attribute.DeclaredNamespace, attribute.DeclaredUri);
        }
    }

    // The innermost prefix bound to the namespace, keeping its exact URI; otherwise the namespace's usual prefix,
    // which the caller declares on the element.
    private string ResolvePrefix(OfficeXmlUtf8Writer writer, int namespaceId, int uriId, bool preferDefault, out string? declare) {
        declare = null;
        if (namespaceId == OfficeXmlNamespaces.None) {
            // An unprefixed element would otherwise land in a default namespace declared above it.
            if (preferDefault && writer.TryFindPrefix(OfficeXmlNamespaces.None, allowDefault: true, out _) == false
                && writer.IsPrefixBound(string.Empty)) {
                declare = string.Empty;
                writer.PushScope(string.Empty, OfficeXmlNamespaces.None, OfficeXmlNamespaces.None);
            }
            return string.Empty;
        }
        if (namespaceId == OfficeXmlNamespaces.Xml) return "xml";
        if (writer.TryFindPrefix(namespaceId, allowDefault: preferDefault, out string prefix)) return prefix;
        prefix = OfficeXmlPrefixes.Canonical(namespaceId);
        if (prefix.Length == 0 && (!preferDefault || writer.IsPrefixBound(string.Empty))) prefix = "ns";
        string candidate = prefix;
        for (int suffix = 1; writer.IsPrefixBound(candidate); suffix++) candidate = prefix + suffix.ToString(System.Globalization.CultureInfo.InvariantCulture);
        declare = Namespaces.GetUriById(uriId);
        writer.PushScope(candidate, namespaceId, uriId);
        return candidate;
    }

    // A parsed name keeps its prefix. Where that prefix is unbound, or bound to another URI (a moved node), the
    // element declares it again; returns the URI to declare, or null when the scope already matches.
    private string? RebindPrefix(OfficeXmlUtf8Writer writer, string prefix, int namespaceId, int uriId) {
        if (namespaceId == OfficeXmlNamespaces.Xml || namespaceId == OfficeXmlNamespaces.Xmlns) return null;
        if (writer.TryGetBinding(prefix, out int boundUri)) {
            if (boundUri == uriId) return null;
        } else if (namespaceId == OfficeXmlNamespaces.None) {
            return null;
        }
        writer.PushScope(prefix, namespaceId, uriId);
        return namespaceId == OfficeXmlNamespaces.None ? string.Empty : Namespaces.GetUriById(uriId);
    }

    private static void WriteDeclaration(OfficeXmlUtf8Writer writer, string prefix, string uri) {
        writer.WriteByte((byte)' ');
        writer.WriteAscii("xmlns");
        if (prefix.Length > 0) {
            writer.WriteByte((byte)':');
            writer.WriteEscaped(prefix, attribute: true);
        }
        writer.WriteByte((byte)'=');
        writer.WriteByte((byte)'"');
        writer.WriteEscaped(uri, attribute: true);
        writer.WriteByte((byte)'"');
    }
}

/// <summary>The prefixes Office applications write for each namespace.</summary>
internal static class OfficeXmlPrefixes {
    internal static string Canonical(int namespaceId) => namespaceId switch {
        OfficeXmlNamespaces.MarkupCompatibility => "mc",
        OfficeXmlNamespaces.Relationships => "r",
        OfficeXmlNamespaces.WordprocessingML => "w",
        OfficeXmlNamespaces.SpreadsheetML => "x",
        OfficeXmlNamespaces.PresentationML => "p",
        OfficeXmlNamespaces.DrawingML => "a",
        OfficeXmlNamespaces.Math => "m",
        OfficeXmlNamespaces.WordprocessingDrawing => "wp",
        OfficeXmlNamespaces.SpreadsheetDrawing => "xdr",
        OfficeXmlNamespaces.Picture => "pic",
        OfficeXmlNamespaces.Chart => "c",
        OfficeXmlNamespaces.Diagram => "dgm",
        OfficeXmlNamespaces.CoreProperties => "cp",
        OfficeXmlNamespaces.DublinCore => "dc",
        OfficeXmlNamespaces.DublinCoreTerms => "dcterms",
        OfficeXmlNamespaces.DocPropsVTypes => "vt",
        OfficeXmlNamespaces.Vml => "v",
        OfficeXmlNamespaces.VmlOffice => "o",
        _ => string.Empty
    };
}
