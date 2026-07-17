namespace OfficeIMO.Email;

/// <summary>Ordered content-line component with properties and nested components.</summary>
public sealed class ContentLineComponent {
    internal const int MaximumTraversalDepth = 256;
    private readonly List<ContentLineProperty> _properties = new List<ContentLineProperty>();
    private readonly List<ContentLineComponent> _components = new List<ContentLineComponent>();

    /// <summary>Creates a component such as VCALENDAR, VEVENT, VALARM, or VCARD.</summary>
    public ContentLineComponent(string name) {
        Name = ContentLineSyntax.RequireToken(name, nameof(name));
    }

    /// <summary>Component name.</summary>
    public string Name { get; set; }

    /// <summary>Ordered component properties.</summary>
    public IList<ContentLineProperty> Properties => _properties;

    /// <summary>Ordered nested components.</summary>
    public IList<ContentLineComponent> Components => _components;

    /// <summary>Returns all direct properties matching <paramref name="name"/>.</summary>
    public IEnumerable<ContentLineProperty> GetProperties(string name) => _properties.Where(property =>
        string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase));

    /// <summary>Returns the first direct property matching <paramref name="name"/>.</summary>
    public ContentLineProperty? GetFirstProperty(string name) => _properties.FirstOrDefault(property =>
        string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase));

    /// <summary>Adds a property and returns it for further configuration.</summary>
    public ContentLineProperty AddProperty(string name, string value = "") {
        var property = new ContentLineProperty(name, value);
        _properties.Add(property);
        return property;
    }

    /// <summary>Replaces all matching direct properties with one value.</summary>
    public ContentLineProperty SetProperty(string name, string value) {
        for (int index = _properties.Count - 1; index >= 0; index--) {
            if (string.Equals(_properties[index].Name, name, StringComparison.OrdinalIgnoreCase))
                _properties.RemoveAt(index);
        }
        return AddProperty(name, value);
    }

    /// <summary>Adds a nested component and returns it.</summary>
    public ContentLineComponent AddComponent(string name) {
        var component = new ContentLineComponent(name);
        _components.Add(component);
        return component;
    }

    /// <summary>Enumerates matching nested components in document order.</summary>
    public IEnumerable<ContentLineComponent> GetComponents(string name, bool recursive = false) {
        var active = new HashSet<ContentLineComponent> { this };
        foreach (ContentLineComponent component in EnumerateComponents(
            this, name, recursive, active, depth: 1)) {
            yield return component;
        }
    }

    private static IEnumerable<ContentLineComponent> EnumerateComponents(
        ContentLineComponent parent, string name, bool recursive,
        ISet<ContentLineComponent> active, int depth) {
        foreach (ContentLineComponent component in parent._components) {
            if (component == null)
                throw new InvalidDataException("A null content-line component cannot be traversed.");
            if (depth + 1 > MaximumTraversalDepth)
                throw new InvalidDataException("The content-line component graph is too deeply nested.");
            if (!active.Add(component))
                throw new InvalidDataException("The content-line component graph contains a cycle.");
            try {
                if (string.Equals(component.Name, name, StringComparison.OrdinalIgnoreCase)) yield return component;
                if (!recursive) continue;
                foreach (ContentLineComponent descendant in EnumerateComponents(
                    component, name, true, active, depth + 1)) {
                    yield return descendant;
                }
            } finally {
                active.Remove(component);
            }
        }
    }
}
