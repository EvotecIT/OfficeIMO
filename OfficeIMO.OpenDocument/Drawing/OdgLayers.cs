using System.Collections;

namespace OfficeIMO.OpenDocument;

/// <summary>Where shapes assigned to a drawing layer are visible.</summary>
public enum OdgLayerDisplay {
    /// <summary>Visible on screen and in printed output.</summary>
    Always,
    /// <summary>Hidden on screen and in printed output.</summary>
    None,
    /// <summary>Visible only on screen.</summary>
    Screen,
    /// <summary>Visible only in printed output.</summary>
    Printer
}

/// <summary>An XML-backed Draw layer. Protection is an editing hint for native applications.</summary>
public sealed class OdgLayer {
    private readonly OdfDocument _document;
    private readonly XElement _element;
    private readonly string _part;
    internal OdgLayer(OdfDocument document, XElement element, string part) { _document = document; _element = element; _part = part; }
    /// <summary>Layer identifier referenced by shapes.</summary>
    public string Name => (string?)_element.Attribute(OdfNamespaces.Draw + "name") ?? string.Empty;
    /// <summary>Screen and print visibility. Invalid legacy masks leave the layer and saved views unchanged.</summary>
    public OdgLayerDisplay Display {
        get => ((string?)_element.Attribute(OdfNamespaces.Draw + "display")) switch {
            null => LegacyDisplay, "always" => OdgLayerDisplay.Always, "none" => OdgLayerDisplay.None,
            "screen" => OdgLayerDisplay.Screen, "printer" => OdgLayerDisplay.Printer,
            _ => throw new InvalidDataException("Invalid Draw layer display value.")
        };
        set {
            if (!Enum.IsDefined(typeof(OdgLayerDisplay), value)) throw new ArgumentOutOfRangeException(nameof(value));
            OdgLegacyLayerSettings.WriteDisplay(_document, _element, value);
            _element.SetAttributeValue(OdfNamespaces.Draw + "display", value.ToString().ToLowerInvariant()); _document.MarkPartDirty(_part);
        }
    }
    /// <summary>Whether native applications should protect the layer from interactive editing.</summary>
    public bool IsProtected {
        get => _element.Attribute(OdfNamespaces.Draw + "protected") is XAttribute attribute ? attribute.Value is "true" or "1"
            : OdgLegacyLayerSettings.Read(_document, _element, "LockedLayers") ?? false;
        set { OdgLegacyLayerSettings.Write(_document, _element, "LockedLayers", value); _element.SetAttributeValue(OdfNamespaces.Draw + "protected", value ? "true" : "false"); _document.MarkPartDirty(_part); }
    }
    private OdgLayerDisplay LegacyDisplay {
        get {
            bool screen = OdgLegacyLayerSettings.Read(_document, _element, "VisibleLayers") ?? true;
            bool print = OdgLegacyLayerSettings.Read(_document, _element, "PrintableLayers") ?? true;
            return screen ? (print ? OdgLayerDisplay.Always : OdgLayerDisplay.Screen) : (print ? OdgLayerDisplay.Printer : OdgLayerDisplay.None);
        }
    }
    /// <summary>Whether this layer participates in the selected output intent.</summary>
    public bool IsVisible(bool forPrint = false) => Display == OdgLayerDisplay.Always || Display == (forPrint ? OdgLayerDisplay.Printer : OdgLayerDisplay.Screen);
}

/// <summary>Layers declared in one document, master, or page scope. Adding a layer set overrides inherited layers in that scope.</summary>
public sealed class OdgLayers : IReadOnlyList<OdgLayer> {
    private readonly OdfDocument _document;
    private readonly XElement _parent;
    private readonly string _part;
    internal OdgLayers(OdfDocument document, XElement parent, string part) { _document = document; _parent = parent; _part = part; }
    private XElement? Container => _parent.Element(OdfNamespaces.Draw + "layer-set");
    internal bool IsDeclared => Container != null;
    private IEnumerable<XElement> Elements => Container?.Elements(OdfNamespaces.Draw + "layer") ?? Enumerable.Empty<XElement>();
    /// <summary>Number of declared layers.</summary>
    public int Count => Elements.Count();
    /// <summary>A layer at a zero-based position.</summary>
    public OdgLayer this[int index] => Wrap(Elements.ElementAtOrDefault(index) ?? throw new ArgumentOutOfRangeException(nameof(index)));
    /// <summary>Finds a layer by its case-sensitive identifier.</summary>
    public OdgLayer? Find(string name) => Elements.Where(element => (string?)element.Attribute(OdfNamespaces.Draw + "name") == name).Select(Wrap).FirstOrDefault();
    /// <summary>Adds a uniquely named layer. Invalid legacy masks leave the existing layer set and saved views unchanged.</summary>
    public OdgLayer Add(string name, OdgLayerDisplay display = OdgLayerDisplay.Always) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Layer name cannot be empty.", nameof(name));
        if (!Enum.IsDefined(typeof(OdgLayerDisplay), display)) throw new ArgumentOutOfRangeException(nameof(display));
        if (Find(name) != null) throw new ArgumentException("Layer names must be unique in their scope.", nameof(name));
        XElement? container = Container;
        bool createdContainer = container == null;
        if (container == null) {
            container = new XElement(OdfNamespaces.Draw + "layer-set");
            XElement? lastDescription = _parent.Name == OdfNamespaces.Draw + "page"
                ? _parent.Elements().LastOrDefault(element => element.Name == OdfNamespaces.Svg + "title" || element.Name == OdfNamespaces.Svg + "desc") : null;
            if (lastDescription == null) _parent.AddFirst(container); else lastDescription.AddAfterSelf(container);
        }
        var element = new XElement(OdfNamespaces.Draw + "layer", new XAttribute(OdfNamespaces.Draw + "name", name));
        container.Add(element); OdgLayer layer = Wrap(element);
        try { layer.Display = display; }
        catch {
            element.Remove();
            if (createdContainer) container.Remove();
            throw;
        }
        return layer;
    }
    private OdgLayer Wrap(XElement element) => new OdgLayer(_document, element, _part);
    /// <inheritdoc />
    public IEnumerator<OdgLayer> GetEnumerator() => Elements.Select(Wrap).GetEnumerator();
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
