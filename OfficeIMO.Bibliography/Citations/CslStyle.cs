using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Bibliography;

/// <summary>An immutable parsed CSL style. Styles and their parent aliases are resolved from caller-supplied data.</summary>
public sealed class CslStyle {
    internal static readonly XNamespace Namespace = "http://purl.org/net/xbiblio/csl";
    private readonly XElement _root;
    internal readonly IReadOnlyDictionary<string, XElement> Macros;
    internal XElement Root => _root;
    internal int MaximumDepth { get; }

    private CslStyle(XElement root, int maximumDepth, string? dependentLocale, CancellationToken token) {
        _root = root;
        MaximumDepth = maximumDepth;
        Id = root.Element(Namespace + "info")?.Element(Namespace + "id")?.Value ?? string.Empty;
        Title = root.Element(Namespace + "info")?.Element(Namespace + "title")?.Value ?? string.Empty;
        DefaultLocale = dependentLocale ?? (string?)root.Attribute("default-locale") ?? "en-US";
        IsNoteStyle = (string?)root.Attribute("class") == "note";
        CslStyleValidation.Validate(root, token);
        CslElementIdentity.Assign(root, token);
        HasConditionalDisambiguation = root.Descendants().Any(element => element.Attribute("disambiguate") != null);
        HasLocatorConditions = root.Descendants().Where(element => element.Name == Namespace + "if" || element.Name == Namespace + "else-if")
            .Attributes().Any(attribute => attribute.Name.LocalName == "locator" ||
                (attribute.Name.LocalName == "variable" || attribute.Name.LocalName == "is-numeric") &&
                attribute.Value.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Contains("locator", StringComparer.Ordinal));
        HasNearNoteCondition = root.Descendants().Attributes("position").Any(attribute =>
            attribute.Value.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Contains("near-note", StringComparer.Ordinal));
        HasSubsequentForm = root.DescendantsAndSelf().Attributes().Any(attribute => attribute.Name.LocalName == "position" ||
            attribute.Name.LocalName == "et-al-subsequent-min" || attribute.Name.LocalName == "et-al-subsequent-use-first" ||
            (attribute.Name.LocalName == "variable" || attribute.Name.LocalName == "is-numeric") &&
            attribute.Value.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Contains("first-reference-note-number", StringComparer.Ordinal));
        var macros = new Dictionary<string, XElement>(StringComparer.Ordinal);
        foreach (XElement macro in root.Elements(Namespace + "macro")) {
            token.ThrowIfCancellationRequested();
            string name = (string?)macro.Attribute("name") ?? string.Empty;
            if (name.Length == 0 || macros.ContainsKey(name)) throw new InvalidDataException("CSL macro names must be nonempty and unique.");
            macros.Add(name, macro);
        }
        Macros = macros;
        IReadOnlyDictionary<string, bool> suffixes = ValidateMacros(token);
        HasExplicitYearSuffix = root.Elements().Where(section => section.Name == Namespace + "citation" || section.Name == Namespace + "bibliography")
            .Select(section => section.Element(Namespace + "layout")).Where(layout => layout != null).Any(layout => ContainsYearSuffix(layout!, suffixes, token));
    }

    /// <summary>Style identifier from its metadata.</summary>
    public string Id { get; }
    /// <summary>Style title from its metadata.</summary>
    public string Title { get; }
    /// <summary>Effective default language dialect, including a dependent style override.</summary>
    public string DefaultLocale { get; }
    /// <summary>Whether citations are intended for notes.</summary>
    public bool IsNoteStyle { get; }
    internal bool HasExplicitYearSuffix { get; }
    internal bool HasConditionalDisambiguation { get; }
    internal bool HasLocatorConditions { get; }
    internal bool HasNearNoteCondition { get; }
    internal bool HasSubsequentForm { get; }

    private static bool ContainsYearSuffix(XElement element, IReadOnlyDictionary<string, bool> suffixes, CancellationToken token) {
        foreach (XElement child in element.DescendantsAndSelf()) {
            token.ThrowIfCancellationRequested();
            if (child.Name != Namespace + "text") continue;
            if ((string?)child.Attribute("variable") == "year-suffix") return true;
            if (!(child.Attribute("macro") is XAttribute macro)) continue;
            if (suffixes[macro.Value]) return true;
        }
        return false;
    }

    /// <summary>Parses bounded CSL XML, prohibits external entities, resolves local parent styles, and rejects recursive macros.</summary>
    public static CslStyle Parse(string xml, CslStyleLoadOptions? options = null, CancellationToken cancellationToken = default) {
        options ??= new CslStyleLoadOptions();
        if (options.MaximumCharacters < 1 || options.MaximumNestingDepth < 1 || options.MaximumNestingDepth > 256)
            throw new ArgumentOutOfRangeException(nameof(options), "CSL limits must be positive and nesting depth cannot exceed 256.");
        var seen = new HashSet<string>(StringComparer.Ordinal);
        string? locale = null;
        for (int depth = 0; depth < options.MaximumNestingDepth; depth++) {
            XElement root = ReadXml(xml, options.MaximumCharacters, options.MaximumNestingDepth, cancellationToken);
            if (root.Name != Namespace + "style" || (string?)root.Attribute("version") != "1.0")
                throw new InvalidDataException("A CSL 1.0 style in the CSL XML namespace is required.");
            if (root.Element(Namespace + "citation") != null) {
                if (root.Element(Namespace + "citation")!.Element(Namespace + "layout") == null)
                    throw new InvalidDataException("A CSL citation requires a layout.");
                return new CslStyle(root, options.MaximumNestingDepth, locale, cancellationToken);
            }
            locale ??= (string?)root.Attribute("default-locale");
            string? parent = root.Element(Namespace + "info")?.Elements(Namespace + "link")
                .FirstOrDefault(link => (string?)link.Attribute("rel") == "independent-parent")?.Attribute("href")?.Value;
            if (parent == null || !seen.Add(parent)) throw new InvalidDataException("CSL parent style is missing or cyclic.");
            xml = options.IndependentStyleResolver?.Invoke(parent) ?? throw new InvalidDataException("Supply IndependentStyleResolver for CSL parent '" + parent + "'.");
        }
        throw new InvalidDataException("CSL parent resolution exceeds MaximumNestingDepth.");
    }

    internal static XElement ReadXml(string xml, int maximumCharacters, int maximumDepth, CancellationToken token) {
        if (xml == null) throw new ArgumentNullException(nameof(xml));
        if (xml.Length > maximumCharacters) throw new InvalidDataException("CSL XML exceeds MaximumCharacters.");
        token.ThrowIfCancellationRequested();
        var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = maximumCharacters };
        using var input = new StringReader(xml);
        using XmlReader reader = XmlReader.Create(input, settings);
        while (reader.Read()) {
            token.ThrowIfCancellationRequested();
            if (reader.Depth > maximumDepth) throw new InvalidDataException("CSL XML exceeds MaximumNestingDepth.");
        }
        using var secondInput = new StringReader(xml);
        using XmlReader second = XmlReader.Create(secondInput, settings);
        XElement result = XElement.Load(second, LoadOptions.PreserveWhitespace);
        token.ThrowIfCancellationRequested();
        return result;
    }

    private IReadOnlyDictionary<string, bool> ValidateMacros(CancellationToken token) {
        var lengths = new Dictionary<string, int>(StringComparer.Ordinal);
        var suffixes = new Dictionary<string, bool>(StringComparer.Ordinal);
        var active = new HashSet<string>(StringComparer.Ordinal);
        foreach (string name in Macros.Keys) Visit(name, active, lengths, suffixes, 0, token);
        foreach (XAttribute reference in _root.Descendants().Attributes("macro")) {
            token.ThrowIfCancellationRequested();
            if (!Macros.ContainsKey(reference.Value)) throw new InvalidDataException("Undefined CSL macro '" + reference.Value + "'.");
        }
        return suffixes;
    }

    private int Visit(string name, HashSet<string> active, IDictionary<string, int> lengths, IDictionary<string, bool> suffixes, int depth, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        if (lengths.TryGetValue(name, out int cached)) {
            if (depth + cached > MaximumDepth) throw new InvalidDataException("Excessively nested CSL macro '" + name + "'.");
            return cached;
        }
        if (depth >= MaximumDepth || !active.Add(name)) throw new InvalidDataException("Recursive or excessively nested CSL macro '" + name + "'.");
        int length = 1;
        bool hasSuffix = false;
        foreach (XElement child in Macros[name].DescendantsAndSelf()) {
            token.ThrowIfCancellationRequested();
            hasSuffix |= child.Name == Namespace + "text" && (string?)child.Attribute("variable") == "year-suffix";
        }
        foreach (XAttribute reference in Macros[name].Descendants().Attributes("macro")) {
            token.ThrowIfCancellationRequested();
            if (!Macros.ContainsKey(reference.Value)) throw new InvalidDataException("Undefined CSL macro '" + reference.Value + "'.");
            length = Math.Max(length, 1 + Visit(reference.Value, active, lengths, suffixes, depth + 1, token));
            hasSuffix |= suffixes[reference.Value];
        }
        active.Remove(name);
        lengths.Add(name, length);
        suffixes.Add(name, hasSuffix);
        return length;
    }
}
