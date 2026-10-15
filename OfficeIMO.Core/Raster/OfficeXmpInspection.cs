using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using OfficeIMO.Provenance;

namespace OfficeIMO.Drawing;

/// <summary>The RDF representation of a retained XMP value.</summary>
public enum OfficeXmpValueKind {
    /// <summary>A scalar literal.</summary>
    Text,
    /// <summary>An rdf:resource URI.</summary>
    Resource,
    /// <summary>An ordered rdf:Seq item.</summary>
    SequenceItem,
    /// <summary>An unordered rdf:Bag item.</summary>
    BagItem,
    /// <summary>An rdf:Alt item, with its language retained.</summary>
    AlternativeItem
}

/// <summary>A located XMP property value with namespace and RDF collection identity retained.</summary>
public sealed class OfficeXmpProperty {
    internal OfficeXmpProperty(string ns, string name, string value, OfficeXmpValueKind kind, string? language, int? index, string source, bool truncated, bool resource) {
        NamespaceUri = ns; Name = name; Value = value; Kind = kind; Language = language;
        ItemIndex = index; Source = source; IsTruncated = truncated; IsResource = resource;
    }
    /// <summary>Namespace URI; XML prefixes are not identities.</summary>
    public string NamespaceUri { get; }
    /// <summary>Local property name.</summary>
    public string Name { get; }
    /// <summary>Retained literal or resource URI.</summary>
    public string Value { get; }
    /// <summary>RDF value representation.</summary>
    public OfficeXmpValueKind Kind { get; }
    /// <summary>Whether the source value used rdf:resource, including inside collections.</summary>
    public bool IsResource { get; }
    /// <summary>Explicit xml:lang qualifier, or null.</summary>
    public string? Language { get; }
    /// <summary>Zero-based collection index, or null for a scalar.</summary>
    public int? ItemIndex { get; }
    /// <summary>Packet and RDF description location.</summary>
    public string Source { get; }
    /// <summary>Whether the scalar exceeded its retention limit.</summary>
    public bool IsTruncated { get; }
}

/// <summary>Bounded parsed XMP evidence; arbitrary structured RDF remains explicitly unsupported.</summary>
public sealed class OfficeXmpInspection {
    internal OfficeXmpInspection(List<OfficeXmpProperty> properties, List<string> issues) { Properties = properties.AsReadOnly(); Issues = issues.AsReadOnly(); }
    /// <summary>Retained namespace-qualified property values.</summary>
    public IReadOnlyList<OfficeXmpProperty> Properties { get; }
    /// <summary>Stable parse, omission and unsupported-structure diagnostics.</summary>
    public IReadOnlyList<string> Issues { get; }
}

/// <summary>Reads Dublin Core and other XMP namespaces using the canonical secure XML loader.</summary>
public static class OfficeXmpInspector {
    private static readonly XNamespace Rdf = "http://www.w3.org/1999/02/22-rdf-syntax-ns#";
    /// <summary>Inspects a standalone XMP packet without resolving external XML or RDF resources.</summary>
    public static OfficeXmpInspection Inspect(byte[] packet, OfficeImageMetadataInspectionOptions? options = null, CancellationToken cancellationToken = default) {
        if (packet == null) throw new ArgumentNullException(nameof(packet));
        OfficeImageMetadataInspectionOptions policy = (options ?? new OfficeImageMetadataInspectionOptions()).Snapshot();
        var properties = new List<OfficeXmpProperty>(); var issues = new List<string>();
        Read(packet, "XMP", policy, properties, issues, cancellationToken);
        return new OfficeXmpInspection(properties, issues);
    }

    internal static void Read(byte[] packet, string source, OfficeImageMetadataInspectionOptions policy,
        List<OfficeXmpProperty> properties, List<string> issues, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        if (packet.Length > policy.MaximumProfileBytes) { AddIssue(issues, "xmp:profile-byte-limit"); return; }
        var xmlPolicy = new OfficeProvenanceOptions { MaxAssetBytes = policy.MaximumProfileBytes, MaxContainerEntries = policy.MaximumXmlNodes, CancellationToken = token };
        XDocument? document;
        try {
            if (!OfficeProvenanceXml.TryLoadDocument(packet, xmlPolicy, out document) || document == null) { AddIssue(issues, "xmp:invalid-xml"); return; }
        } catch (InvalidDataException) { AddIssue(issues, "xmp:xml-node-limit"); return; }
        int descriptionIndex = 0;
        long retainedCharacters = properties.Sum(property => (long)property.Value.Length + (property.Language?.Length ?? 0));
        foreach (XElement description in document.Descendants(Rdf + "Description")) {
            token.ThrowIfCancellationRequested();
            if (description.Parent?.Name != Rdf + "RDF" ||
                description.Ancestors().Any(node => node.Attribute(Rdf + "parseType")?.Value == "Literal")) continue;
            string location = source + "/rdf:Description[" + descriptionIndex++ + "]";
            foreach (XAttribute attribute in description.Attributes()) {
                if (attribute.IsNamespaceDeclaration || attribute.Name.Namespace == Rdf || attribute.Name.Namespace == XNamespace.Xml) continue;
                Add(attribute.Name, attribute.Value, OfficeXmpValueKind.Text, (string?)description.Attribute(XNamespace.Xml + "lang"), null, location, policy, properties, issues, ref retainedCharacters);
            }
            foreach (XElement property in description.Elements()) {
                token.ThrowIfCancellationRequested();
                if (property.Attribute(Rdf + "parseType") != null || property.Attribute(Rdf + "nodeID") != null) { AddIssue(issues, "xmp:structured-property-omitted"); continue; }
                XElement? collection = property.Elements().SingleOrDefaultSafe();
                if (collection != null && (collection.Name == Rdf + "Seq" || collection.Name == Rdf + "Bag" || collection.Name == Rdf + "Alt")) {
                    int index = 0;
                    foreach (XElement item in collection.Elements()) {
                        token.ThrowIfCancellationRequested();
                        if (item.Name != Rdf + "li" || item.HasElements || item.Attribute(Rdf + "parseType") != null) { AddIssue(issues, "xmp:structured-property-omitted"); index++; continue; }
                        string? resource = (string?)item.Attribute(Rdf + "resource");
                        if (resource != null && !string.IsNullOrWhiteSpace(item.Value)) { AddIssue(issues, "xmp:contradictory-resource"); index++; continue; }
                        OfficeXmpValueKind kind = collection.Name == Rdf + "Seq" ? OfficeXmpValueKind.SequenceItem : collection.Name == Rdf + "Bag" ? OfficeXmpValueKind.BagItem : OfficeXmpValueKind.AlternativeItem;
                        Add(property.Name, resource ?? item.Value, kind, (string?)item.Attribute(XNamespace.Xml + "lang") ?? (string?)property.Attribute(XNamespace.Xml + "lang") ?? (string?)description.Attribute(XNamespace.Xml + "lang"), index++, location, policy, properties, issues, ref retainedCharacters, resource != null);
                    }
                } else if (!property.HasElements) {
                    string? resource = (string?)property.Attribute(Rdf + "resource");
                    if (resource != null && !string.IsNullOrWhiteSpace(property.Value)) { AddIssue(issues, "xmp:contradictory-resource"); continue; }
                    Add(property.Name, resource ?? property.Value, resource == null ? OfficeXmpValueKind.Text : OfficeXmpValueKind.Resource,
                        (string?)property.Attribute(XNamespace.Xml + "lang") ?? (string?)description.Attribute(XNamespace.Xml + "lang"), null, location, policy, properties, issues, ref retainedCharacters);
                } else AddIssue(issues, "xmp:structured-property-omitted");
            }
        }
        if (!document.Descendants(Rdf + "RDF").Any() && document.Root?.Name != Rdf + "RDF") AddIssue(issues, "xmp:rdf-envelope-unsupported");
    }

    private static XElement? SingleOrDefaultSafe(this IEnumerable<XElement> values) {
        using IEnumerator<XElement> enumerator = values.GetEnumerator();
        if (!enumerator.MoveNext()) return null;
        XElement result = enumerator.Current;
        return enumerator.MoveNext() ? null : result;
    }
    private static void Add(XName name, string value, OfficeXmpValueKind kind, string? language, int? index, string source,
        OfficeImageMetadataInspectionOptions policy, List<OfficeXmpProperty> properties, List<string> issues, ref long retainedCharacters, bool resource = false) {
        if (name.NamespaceName.Length == 0 || name.Namespace == Rdf) { AddIssue(issues, "xmp:unqualified-property-omitted"); return; }
        if (properties.Count >= policy.MaximumXmpProperties) { AddIssue(issues, "xmp:property-limit"); return; }
        bool truncated = value.Length > policy.MaximumValueCharacters;
        if (truncated) { value = value.Substring(0, policy.MaximumValueCharacters); AddIssue(issues, "xmp:value-character-limit"); }
        if (language != null && language.Length > policy.MaximumValueCharacters) { language = language.Substring(0, policy.MaximumValueCharacters); AddIssue(issues, "xmp:value-character-limit"); truncated = true; }
        long characters = value.Length + (long)(language?.Length ?? 0);
        if (characters > policy.MaximumRetainedXmpCharacters - retainedCharacters) { AddIssue(issues, "xmp:aggregate-character-limit"); return; }
        retainedCharacters += characters;
        properties.Add(new OfficeXmpProperty(name.NamespaceName, name.LocalName, value, kind, language, index, source, truncated, resource || kind == OfficeXmpValueKind.Resource));
    }
    internal static void AddIssue(List<string> issues, string issue) { if (!issues.Contains(issue)) issues.Add(issue); }
}
