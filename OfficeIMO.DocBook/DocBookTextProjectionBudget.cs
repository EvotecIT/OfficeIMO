using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Xml.Linq;

namespace OfficeIMO.DocBook;

/// <summary>Bounds aggregate text retained by a shared-model projection and reuses each element projection.</summary>
internal sealed class DocBookTextProjectionBudget {
    private static readonly HashSet<string> AuthorNameParts = new HashSet<string>(StringComparer.Ordinal) {
        "honorific", "firstname", "othername", "surname", "lineage"
    };

    private readonly DocBookDiagnosticCollector _diagnostics;
    private readonly XNamespace _docBookNamespace;
    private readonly CancellationToken _cancellationToken;
    private readonly Dictionary<XElement, string> _elementValues = new Dictionary<XElement, string>();
    private readonly Dictionary<XElement, string> _authorValues = new Dictionary<XElement, string>();
    private long _remaining;

    internal DocBookTextProjectionBudget(long maximumCharacters, DocBookDiagnosticCollector diagnostics,
        XNamespace docBookNamespace, CancellationToken cancellationToken) {
        if (maximumCharacters < 1) throw new ArgumentOutOfRangeException(nameof(maximumCharacters));
        _remaining = maximumCharacters;
        _diagnostics = diagnostics ?? throw new ArgumentNullException(nameof(diagnostics));
        _docBookNamespace = docBookNamespace;
        _cancellationToken = cancellationToken;
    }

    internal string GetPrimaryText(XElement element, DocBookNodeKind kind, XNamespace ns, string? path) {
        if (kind == DocBookNodeKind.Section || kind == DocBookNodeKind.Table ||
            kind == DocBookNodeKind.Figure || kind == DocBookNodeKind.Info) {
            return GetElementValue(element.Element(ns + "title"), path);
        }
        if (kind == DocBookNodeKind.Author) return GetAuthorName(element, path);
        if (kind == DocBookNodeKind.Title || kind == DocBookNodeKind.Subtitle || kind == DocBookNodeKind.Link ||
            kind == DocBookNodeKind.Entry || kind == DocBookNodeKind.Caption) return GetElementValue(element, path);
        return element.HasElements && kind != DocBookNodeKind.Paragraph &&
               kind != DocBookNodeKind.ProgramListing && kind != DocBookNodeKind.Screen
            ? string.Empty
            : GetElementValue(element, path);
    }

    internal string GetElementValue(XElement? element, string? path) {
        if (element == null) return string.Empty;
        if (_elementValues.TryGetValue(element, out string? cached)) return cached;
        string value = MaterializeExact(() => element.DescendantNodes().OfType<XText>().Where(text =>
            text.Ancestors().TakeWhile(ancestor => !ReferenceEquals(ancestor, element))
                .All(ancestor => ancestor.Name != _docBookNamespace + "indexterm")), path);
        _elementValues.Add(element, value);
        return value;
    }

    internal string GetElementValueExcluding(XElement element, XName excludedSubtree, string? path) =>
        MaterializeExact(() => element.DescendantNodes().OfType<XText>().Where(text =>
            text.Ancestors().TakeWhile(ancestor => !ReferenceEquals(ancestor, element)).All(ancestor =>
                ancestor.Name != _docBookNamespace + "indexterm" && ancestor.Name != excludedSubtree)), path);

    internal string GetTextValue(XText text, string? path) {
        _cancellationToken.ThrowIfCancellationRequested();
        if (text.Value.Length > _remaining) {
            ReportLimit(path);
            return string.Empty;
        }
        _remaining -= text.Value.Length;
        return text.Value;
    }

    internal string GetAuthorName(XElement author, string? path) {
        if (_authorValues.TryGetValue(author, out string? cached)) return cached;
        IEnumerable<XText> nameNodes = SelectAuthorNameNodes(author);
        var parts = new List<string>();
        foreach (XText node in nameNodes) {
            _cancellationToken.ThrowIfCancellationRequested();
            string part = node.Value.Trim();
            if (part.Length > 0) parts.Add(part);
        }
        long length = parts.Sum(part => (long)part.Length) + Math.Max(0, parts.Count - 1);
        string value;
        if (length > _remaining) {
            ReportLimit(path);
            value = string.Empty;
        } else {
            _remaining -= length;
            value = string.Join(" ", parts);
        }
        _authorValues.Add(author, value);
        return value;
    }

    private string MaterializeExact(Func<IEnumerable<XText>> textFactory, string? path) {
        long length = 0;
        foreach (XText text in textFactory()) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (text.Value.Length > _remaining - length) {
                ReportLimit(path);
                return string.Empty;
            }
            length += text.Value.Length;
        }
        _remaining -= length;
        if (length == 0) return string.Empty;
        var value = new System.Text.StringBuilder((int)Math.Min(length, int.MaxValue));
        foreach (XText text in textFactory()) {
            _cancellationToken.ThrowIfCancellationRequested();
            value.Append(text.Value);
        }
        return value.ToString();
    }

    private static IEnumerable<XText> SelectAuthorNameNodes(XElement author) {
        XNamespace docBookNamespace = author.Name.Namespace;
        XElement? personName = author.Elements().FirstOrDefault(element =>
            element.Name.Namespace == docBookNamespace && element.Name.LocalName == "personname");
        if (personName != null) return personName.DescendantNodes().OfType<XText>().Where(text =>
            text.Ancestors().TakeWhile(element => !ReferenceEquals(element, personName))
                .All(element => element.Name.Namespace == docBookNamespace));

        XElement[] components = author.Elements().Where(element =>
            element.Name.Namespace == docBookNamespace && AuthorNameParts.Contains(element.Name.LocalName)).ToArray();
        if (components.Length > 0) return components.SelectMany(component =>
            component.DescendantNodes().OfType<XText>().Where(text =>
                text.Ancestors().TakeWhile(element => !ReferenceEquals(element, component))
                    .All(element => element.Name.Namespace == docBookNamespace)));

        return author.Nodes().OfType<XText>();
    }

    private void ReportLimit(string? path) {
        _diagnostics.Add(new DocBookDiagnostic("DB123", DocBookDiagnosticSeverity.Warning,
            "Shared-model text exceeded MaxTotalTextCharacters; additional text remains available in the native DocBook XML but was omitted from the projection.", path));
    }
}
