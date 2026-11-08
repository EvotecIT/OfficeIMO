using System.Globalization;
using System.Xml.Linq;

namespace OfficeIMO.IWork.Internal;

/// <summary>Indexes the immutable XML snapshot and resolves scoped native styles without repeated tree scans.</summary>
internal sealed partial class IWorkLegacyXmlCatalog {
    private static readonly XNamespace Sf = IWorkLegacyXmlSource.Sf;
    private static readonly XNamespace Sfa = IWorkLegacyXmlSource.Sfa;
    private readonly IWorkSourceDocument _source;
    private readonly IWorkLegacyXmlSource _legacy;
    private readonly IWorkProjectionBudget _budget;
    private readonly Dictionary<string, XElement?> _ids = new(StringComparer.Ordinal);
    private readonly Dictionary<XElement, XElement> _ownerStylesheets = new();
    private readonly Dictionary<XElement, string> _ownerStylesheetReferences = new();
    private readonly Dictionary<(XElement Scope, string Kind, string Identifier), XElement?> _styles = new();
    private readonly Dictionary<(XElement Style, string Name), XElement?> _properties = new();
    private readonly List<IWorkDiagnostic> _diagnostics = new();
    private readonly HashSet<string> _lossCodes = new(StringComparer.Ordinal);
    private readonly Dictionary<XElement, IWorkObjectIdentity> _identities = new();
    private Dictionary<string, IWorkPackageEntry>? _packageEntries;
    private Dictionary<IWorkPackageEntry, (int Width, int Height, long DecodedBytes)>? _validatedImages;

    internal IWorkLegacyXmlCatalog(IWorkSourceDocument source, IWorkLegacyXmlSource legacy,
        IWorkProjectionBudget budget) {
        _source = source;
        _legacy = legacy;
        _budget = budget;
        foreach (XElement element in legacy.Root.DescendantsAndSelf()) {
            source.CancellationToken.ThrowIfCancellationRequested();
            string? id = (string?)element.Attribute(Sfa + "ID");
            if (id != null && id.Length > 0) {
                if (_ids.ContainsKey(id)) _ids[id] = null;
                else _ids.Add(id, element);
            }
            if (element.Name.LocalName == "stylesheet" && element.Parent != null)
                _ownerStylesheets[element.Parent] = element;
            if (element.Name == Sf + "stylesheet-ref" && element.Parent != null
                && (string?)element.Attribute(Sfa + "IDREF") is string scopeReference)
                _ownerStylesheetReferences[element.Parent] = scopeReference;
            string? identifier = (string?)element.Attribute(Sf + "ident");
            if (element.Name.Namespace != Sf || identifier == null || identifier.Length == 0) continue;
            XElement? scope = element.Ancestors().FirstOrDefault(IsStylesheet);
            if (scope == null) continue;
            var key = (scope, element.Name.LocalName, identifier);
            if (_styles.ContainsKey(key)) _styles[key] = null;
            else _styles.Add(key, element);
        }
    }

    internal IReadOnlyList<IWorkDiagnostic> Diagnostics => _diagnostics;

    internal IWorkObjectIdentity Identity(XElement element) {
        if (!_identities.TryGetValue(element, out IWorkObjectIdentity? identity)) {
            identity = new IWorkObjectIdentity(_legacy.EntryPath, (ulong)_identities.Count + 1);
            _identities.Add(element, identity);
        }
        return identity;
    }

    internal IWorkPackageEntry? PackageEntry(string path) {
        _packageEntries ??= _source.Entries.ToDictionary(entry => entry.Path, StringComparer.Ordinal);
        return _packageEntries.TryGetValue(path, out IWorkPackageEntry? entry) ? entry : null;
    }

    /// <summary>Reuses successful raster validation within this projection; occurrence budgets remain with the reader.</summary>
    internal bool TryGetValidatedImage(IWorkPackageEntry entry, out (int Width, int Height, long DecodedBytes) metadata) {
        _source.CancellationToken.ThrowIfCancellationRequested();
        if (_validatedImages != null && _validatedImages.TryGetValue(entry, out metadata)) return true;
        metadata = default;
        return false;
    }

    /// <summary>Retains only metadata from a fully validated captured raster after its occurrence passed every budget.</summary>
    internal void RememberValidatedImage(IWorkPackageEntry entry, int width, int height, long decodedBytes) {
        _source.CancellationToken.ThrowIfCancellationRequested();
        (_validatedImages ??= new()).Add(entry, (width, height, decodedBytes));
    }

    internal void AddLoss(string code, string message) {
        if (_lossCodes.Add(code))
            _diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, code, message,
                _legacy.EntryPath));
    }

    internal XElement? Resolve(XElement? element, string kind) {
        if (element == null) return null;
        if (element.Name == Sf + kind) return element;
        string? reference = (string?)element.Attribute(Sfa + "IDREF");
        if (element.Name == Sf + (kind + "-ref") && reference != null
            && _ids.TryGetValue(reference, out XElement? value) && value?.Name == Sf + kind) return value;
        AddLoss("IWORK_LEGACY_REFERENCE_UNRESOLVED",
            "An XML object reference is missing, ambiguous, or has an unsupported object type.");
        return null;
    }

    internal XElement? Style(XElement context, string? identifier, string kind) {
        if (identifier == null || identifier.Length == 0) return null;
        if (_ids.TryGetValue(identifier, out XElement? direct)) {
            if (direct?.Name == Sf + kind) return direct;
            AddLoss("IWORK_LEGACY_STYLE_UNRESOLVED",
                "A selected XML style is missing or ambiguous; its unqualified formatting is omitted.");
            return null;
        }
        XElement? style = FindScopedStyle(FindScope(context), identifier, kind);
        if (style == null) AddLoss("IWORK_LEGACY_STYLE_UNRESOLVED",
            "A selected XML style is missing or ambiguous; its unqualified formatting is omitted.");
        return style;
    }

    internal XElement? Property(XElement? style, string name) {
        if (style == null) return null;
        var key = (style, name);
        if (_properties.TryGetValue(key, out XElement? cached)) return cached;
        XElement? current = style;
        var visited = new HashSet<XElement>();
        while (current != null) {
            _source.CancellationToken.ThrowIfCancellationRequested();
            if (!visited.Add(current) || visited.Count > _budget.MaxTextStyleInheritanceDepth)
                throw new InvalidDataException("The iWork '09 style inheritance exceeds its depth limit or contains a cycle.");
            XElement? property = current.Element(Sf + "property-map")?.Element(Sf + name);
            // Empty native property slots inherit. An explicit sf:null child clears the inherited value.
            if (property?.HasElements == true || property?.HasAttributes == true) {
                _properties.Add(key, property);
                return property;
            }
            string? parent = (string?)current.Attribute(Sf + "parent-ident");
            if (parent == null || parent.Length == 0) break;
            XElement? scope = FindScope(current);
            if (parent == (string?)current.Attribute(Sf + "ident")) scope = ParentScope(scope);
            current = FindScopedStyle(scope, parent, current.Name.LocalName);
            if (current == null) AddLoss("IWORK_LEGACY_STYLE_UNRESOLVED",
                "A selected XML style is missing or ambiguous; its unqualified formatting is omitted.");
        }
        _properties.Add(key, null);
        return null;
    }

    private XElement? FindScopedStyle(XElement? scope, string identifier, string kind) {
        var visited = new HashSet<XElement>();
        while (scope != null) {
            _source.CancellationToken.ThrowIfCancellationRequested();
            if (!visited.Add(scope) || visited.Count > _budget.MaxTextStyleInheritanceDepth)
                throw new InvalidDataException("The iWork '09 stylesheet hierarchy exceeds its depth limit or contains a cycle.");
            if (_styles.TryGetValue((scope, kind, identifier), out XElement? style)) return style;
            scope = ParentScope(scope);
        }
        return null;
    }

    private XElement? FindScope(XElement context) {
        foreach (XElement ancestor in context.AncestorsAndSelf()) {
            if (IsStylesheet(ancestor)) return ancestor;
            if (_ownerStylesheetReferences.TryGetValue(ancestor, out string? reference)) {
                if (_ids.TryGetValue(reference, out XElement? selected) && selected != null && IsStylesheet(selected))
                    return selected;
                AddLoss("IWORK_LEGACY_STYLE_UNRESOLVED",
                    "A selected XML stylesheet reference is missing or ambiguous; its unqualified formatting is omitted.");
                return null;
            }
            if (_ownerStylesheets.TryGetValue(ancestor, out XElement? scope)) return scope;
        }
        return null;
    }

    private XElement? ParentScope(XElement? scope) {
        string? reference = (string?)scope?.Element(Sf + "parent-ref")?.Attribute(Sfa + "IDREF");
        return reference != null && _ids.TryGetValue(reference, out XElement? parent)
            && parent != null && IsStylesheet(parent) ? parent : null;
    }

    private static bool IsStylesheet(XElement element) => element.Name.LocalName == "stylesheet";

    internal IWorkColor? Color(XElement? color) {
        color = Resolve(color, "color");
        if (color == null) return null;
        double? red = AttributeNumber(color, Sfa + "r");
        double? green = AttributeNumber(color, Sfa + "g");
        double? blue = AttributeNumber(color, Sfa + "b");
        if (AttributeNumber(color, Sfa + "w") is double white) red = green = blue = white;
        double alpha = AttributeNumber(color, Sfa + "a") ?? 1;
        if (red is >= 0 and <= 1 && green is >= 0 and <= 1 && blue is >= 0 and <= 1
            && alpha is >= 0 and <= 1)
            return new IWorkColor((byte)Math.Round(red.Value * 255), (byte)Math.Round(green.Value * 255),
                (byte)Math.Round(blue.Value * 255), (byte)Math.Round(alpha * 255));
        AddLoss("IWORK_LEGACY_COLOR_UNSUPPORTED", "A selected XML color uses an unqualified color space or component range.");
        return null;
    }

    internal static double? AttributeNumber(XElement? element, XName name) =>
        double.TryParse((string?)element?.Attribute(name), NumberStyles.Float, CultureInfo.InvariantCulture,
            out double value) && !double.IsNaN(value) && !double.IsInfinity(value) ? value : null;
}
