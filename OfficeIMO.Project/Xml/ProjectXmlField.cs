using System.Xml.Linq;

namespace OfficeIMO.Project;

/// <summary>A reflection-free scalar mapping. Binary formats do not use this XML mapping layer.</summary>
internal sealed class ProjectXmlField<T> {
    internal ProjectXmlField(string name, Func<T, ProjectDocument, string?> get, Action<T, string, ProjectDocument, XElement> read) {
        Name = name; Get = get; Read = read;
    }
    internal string Name { get; }
    internal Func<T, ProjectDocument, string?> Get { get; }
    internal Action<T, string, ProjectDocument, XElement> Read { get; }
}

internal static partial class ProjectXmlFields {
    internal static void Read<T>(T model, XElement element, ProjectDocument document, IEnumerable<ProjectXmlField<T>> fields) where T : class {
        foreach (var field in fields) {
            var nodes = element.Elements(element.Name.Namespace + field.Name).ToArray();
            if (nodes.Length > 1) throw new InvalidDataException("Duplicate scalar " + field.Name + " at " + ProjectXmlValue.Location(element));
            if (nodes.Length == 0) continue;
            if (nodes[0].HasElements) throw new InvalidDataException("Scalar " + field.Name + " contains nested XML.");
            try { field.Read(model, nodes[0].Value, document, element); }
            catch (Exception error) when (error is FormatException || error is OverflowException || error is ArgumentException) {
                throw new InvalidDataException("Invalid " + field.Name + " at " + ProjectXmlValue.Location(element) + ".", error);
            }
        }
    }
    internal static void Snapshot<T>(T model, ProjectDocument document, IEnumerable<ProjectXmlField<T>> fields) where T : class {
        foreach (var field in fields) document.Source!.Snapshot(model, field.Name, field.Get(model, document));
    }
    internal static void Write<T>(T model, XElement element, ProjectDocument document, IEnumerable<ProjectXmlField<T>> fields, string[] order) where T : class {
        foreach (var field in fields) Apply(document, model, element, field.Name, field.Get(model, document), order);
    }
    internal static void Apply(ProjectDocument document, object model, XElement element, string name, string? value, string[] order) {
        if (document.Source?.Capturing == true) { document.Source.Snapshot(model, name, value); return; }
        if (document.Source?.Unchanged(model, name, value) == true) return;
        var existing = element.Element(element.Name.Namespace + name);
        if (value == null) { existing?.Remove(); return; }
        if (existing != null) { existing.Value = value; return; }
        Insert(element, new XElement(element.Name.Namespace + name, value), order);
    }
    internal static void Insert(XElement parent, XElement child, string[] order) {
        int rank = Array.IndexOf(order, child.Name.LocalName);
        if (rank >= 0) {
            var after = parent.Elements().FirstOrDefault(e => e.Name.Namespace == parent.Name.Namespace && Array.IndexOf(order, e.Name.LocalName) is int other && other > rank);
            if (after != null) { after.AddBeforeSelf(child); return; }
        }
        parent.Add(child);
    }
}
