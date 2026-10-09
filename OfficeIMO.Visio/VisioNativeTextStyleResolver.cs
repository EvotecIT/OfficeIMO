using System.Globalization;
using System.Threading;
using System.Xml.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

/// <summary>Operation-local projection of cached text properties; never edits native source.</summary>
internal sealed class VisioNativeTextStyleResolver {
    internal const string DiagnosticCode = "VISIO_TEXT_STYLE_INHERITANCE";
    private const int MaximumInheritanceDepth = 64;
    private readonly CancellationToken _cancellation;
    private readonly ICollection<OfficeImageExportDiagnostic>? _diagnostics;
    private readonly string? _source;
    private readonly Dictionary<string, XElement> _styles = new(StringComparer.Ordinal);
    private readonly HashSet<string> _ambiguous = new(StringComparer.Ordinal);
    private readonly HashSet<string> _reported = new(StringComparer.Ordinal);
    private readonly Dictionary<string, XElement?> _styleSections = new(StringComparer.Ordinal);

    internal VisioNativeTextStyleResolver(VisioDocument? document, CancellationToken cancellation,
        ICollection<OfficeImageExportDiagnostic>? diagnostics = null, string? source = null) {
        _cancellation = cancellation; _diagnostics = diagnostics; _source = source;
        if (document == null) return;
        foreach (string id in new[] { "0", "1", "2" })
            _styles.Add(id, VisioDocument.CreateGeneratedStyleSheet(VisioDocument.VisioNamespace, id, document.PreservedGeneratedStyleSheets));
        foreach (XElement style in document.PreservedAdditionalStyleSheets) {
            _cancellation.ThrowIfCancellationRequested();
            string? id = Id((string?)style.Attribute("ID"));
            if (id == null) continue;
            if (_styles.ContainsKey(id)) _ambiguous.Add(id);
            else _styles.Add(id, style);
        }
    }

    internal XElement? Resolve(VisioShape shape, bool character, bool replacement = false) => ResolveShape(shape, character, replacement,
        new HashSet<VisioShape>(), 0);

    // Editing needs the same cached inheritance while retaining native font identities.
    internal XElement? ResolveNative(VisioShape shape, bool character, bool replacement = false) => ResolveShape(shape, character, replacement,
        new HashSet<VisioShape>(), 0, nativeFonts: true);

    internal bool HasStyleText(string reference) => Style(reference, true)?.Descendants().Any(e => e.Name.LocalName == "Cell") == true ||
        Style(reference, false)?.Descendants().Any(e => e.Name.LocalName == "Cell") == true;
    internal bool UsesUnformattedBase(string? reference) => Id(reference) == "0" && !HasStyleText(reference!);

    internal XElement? Resolve(VisioConnector connector, bool character, bool replacement = false) {
        XElement? local = VisioDocument.GetRenderTextSection(connector.PreservedNonGeometrySections,
            character ? connector.CharacterSectionSource : connector.ParagraphSectionSource, connector.TextStyle, character);
        XElement? inherited = Style(connector.NativeStyleReferences?.TextStyle, character);
        return Merge(local, inherited, character, replacement ? PlainText(connector.Label) : connector.PreservedTextElement);
    }

    private XElement? ResolveShape(VisioShape shape, bool character, bool replacement, HashSet<VisioShape> seen, int depth, bool nativeFonts = false) {
        _cancellation.ThrowIfCancellationRequested();
        XElement? local = VisioDocument.GetRenderTextSection(shape.PreservedNonGeometrySections,
            character ? shape.CharacterSectionSource : shape.ParagraphSectionSource, shape.TextStyle, character);
        if (!nativeFonts) local = VisioDocument.GetMasterFontRenderSection(shape, local);
        else if (local != null) {
            local = new XElement(local);
            foreach (XElement cell in local.Descendants().Where(element => element.Name.LocalName == "Cell")) cell.AddAnnotation(shape);
        }
        if (depth >= MaximumInheritanceDepth || !seen.Add(shape)) {
            Warn("master-chain", "A cyclic or excessively deep master text chain uses available cached properties and render defaults.");
            return Merge(local, null, character, depth == 0 ? shape.PreservedTextElement : null, nativeRows: nativeFonts);
        }
        string? binding = shape.NativeStyleReferences?.TextStyle;
        // An explicit instance TextStyle selects style properties instead of master text properties.
        XElement? inherited = !string.IsNullOrEmpty(binding) ? Style(binding, character) : shape.MasterShape == null ? null :
            ResolveShape(shape.MasterShape, character, false, seen, depth + 1, nativeFonts);
        seen.Remove(shape);
        // Only the rendered instance's markers need effective fallback rows. An
        // ancestor's synthetic zero row must not displace its first native row.
        return Merge(local, inherited, character, depth == 0 ? replacement ? PlainText(shape.Text) : shape.PreservedTextElement : null, nativeRows: nativeFonts);
    }

    private XElement? Style(string? reference, bool character) {
        if (string.IsNullOrEmpty(reference)) return null;
        string? id = Id(reference);
        string key = (character ? "C:" : "P:") + reference;
        if (_styleSections.TryGetValue(key, out XElement? cached)) return cached;
        var chain = new List<XElement>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        while (id != null) {
            _cancellation.ThrowIfCancellationRequested();
            if (chain.Count >= MaximumInheritanceDepth || !seen.Add(id)) {
                Warn("style-chain:" + reference, "A cyclic or excessively deep text style chain uses available cached properties and render defaults.");
                break;
            }
            if (_ambiguous.Contains(id) || !_styles.TryGetValue(id, out XElement? style)) {
                Warn("style-reference:" + reference, "An unresolved or ambiguous text style reference uses available cached properties and render defaults.");
                break;
            }
            XElement? enabled = style.Elements().FirstOrDefault(e => e.Name.LocalName == "Cell" && (string?)e.Attribute("N") == "EnableTextProps");
            if ((string?)enabled?.Attribute("V") == "0") {
                // Exclusion is specified; continuation through a disabled style's
                // parent has no qualified native oracle. Do not guess silently.
                Warn("disabled-style:" + id, "A style excludes text properties; inheritance through its parent is unqualified and uses render defaults.");
                break;
            }
            chain.Add(style);
            string? parent = (string?)style.Attribute("TextStyle");
            if (string.IsNullOrEmpty(parent)) break;
            // Serialized base style 0 commonly names itself. It is the terminal
            // base, while self references in other style IDs remain malformed.
            if (id == "0" && Id(parent) == "0") break;
            id = Id(parent);
            if (id == null) Warn("style-reference:" + parent, "An invalid text style reference uses available cached properties and render defaults.");
        }
        if (id == null && chain.Count == 0)
            Warn("style-reference:" + reference, "An invalid text style reference uses render defaults.");
        XElement? result = null;
        for (int i = chain.Count - 1; i >= 0; i--) {
            XElement? section = chain[i].Elements().FirstOrDefault(e => e.Name.LocalName == "Section" &&
                (string?)e.Attribute("N") == (character ? "Character" : "Paragraph"));
            result = Merge(section, result, character, null);
        }
        _styleSections[key] = result;
        return result;
    }

    private XElement? Merge(XElement? local, XElement? inherited, bool character, XElement? text, bool nativeRows = false) {
        if (local == null && inherited == null) return null;
        XNamespace ns = local?.Name.Namespace ?? inherited!.Name.Namespace;
        var result = new XElement(ns + "Section", new XAttribute("N", character ? "Character" : "Paragraph"));
        bool sectionDeleted = Deleted(local);
        Dictionary<string, XElement> localRows = Rows(local), inheritedRows = Rows(inherited);
        XElement? first = inheritedRows.Values.FirstOrDefault();
        var indices = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        void AddIndex(string index) { if (seen.Add(index)) indices.Add(index); }
        foreach (string index in localRows.Keys) AddIndex(index);
        foreach (string index in inheritedRows.Keys) AddIndex(index);
        if (text != null) {
            if (!nativeRows || text.Nodes().TakeWhile(node => node is not XElement marker || marker.Name.LocalName != (character ? "cp" : "pp"))
                .OfType<XText>().Any(node => node.Value.Length > 0)) AddIndex("0");
            foreach (XElement marker in text.Descendants()) {
                _cancellation.ThrowIfCancellationRequested();
                if (marker.Name.LocalName == (character ? "cp" : "pp")) AddIndex(RowIndex((string?)marker.Attribute("IX")));
                if (indices.Count > OfficeTextLayoutEngine.MaximumLayoutTextRuns) return null;
            }
        }
        if (indices.Count > OfficeTextLayoutEngine.MaximumLayoutTextRuns) {
            Warn("row-limit", "Text style row limits were exceeded; render defaults are used.");
            return null;
        }
        if (indices.Count == 0) AddIndex("0");
        foreach (string index in indices) {
            _cancellation.ThrowIfCancellationRequested();
            localRows.TryGetValue(index, out XElement? own);
            inheritedRows.TryGetValue(index, out XElement? parent);
            parent ??= first;
            XElement row = new(ns + "Row", new XAttribute("IX", index));
            if (!sectionDeleted && !Deleted(own)) {
                var cells = new Dictionary<string, XElement>(StringComparer.Ordinal);
                AddCells(parent, cells);
                AddCells(own, cells);
                row.Add(cells.Values.Select(cell => {
                    var copy = new XElement(cell);
                    if (nativeRows && cell.Annotation<VisioShape>() is VisioShape owner) copy.AddAnnotation(owner);
                    return copy;
                }));
            }
            result.Add(row);
        }
        return result;
    }

    private Dictionary<string, XElement> Rows(XElement? section) {
        var rows = new Dictionary<string, XElement>(StringComparer.Ordinal);
        int position = 0;
        foreach (XElement row in section?.Elements().Where(e => e.Name.LocalName == "Row") ?? Enumerable.Empty<XElement>()) {
            _cancellation.ThrowIfCancellationRequested();
            if (position >= OfficeTextLayoutEngine.MaximumLayoutTextRuns) {
                Warn("row-limit", "Text style row limits were exceeded; available cached rows are used.");
                break;
            }
            string index = RowIndex((string?)row.Attribute("IX"), position);
            if (rows.ContainsKey(index)) Warn("duplicate-row:" + index, "Duplicate text style row indices use the first cached row.");
            else rows.Add(index, row);
            position++;
        }
        return rows;
    }

    private void AddCells(XElement? row, Dictionary<string, XElement> cells) {
        if (row == null || Deleted(row)) return;
        foreach (XElement cell in row.Elements().Where(e => e.Name.LocalName == "Cell")) {
            _cancellation.ThrowIfCancellationRequested();
            string? name = (string?)cell.Attribute("N"), value = (string?)cell.Attribute("V");
            if (name == null) continue;
            if (value == null) {
                Warn("uncached-cell:" + name, "An uncached text formula uses inherited properties or render defaults; formulas are preserved without evaluation.");
                continue;
            }
            // Concrete V, including F=Inh, is authoritative. Themed is a blocked
            // cached value rather than permission to substitute an ancestor.
            if (string.Equals(value, "themed", StringComparison.OrdinalIgnoreCase))
                Warn("themed-cell:" + name, "A themed text property uses render defaults; dynamic theme evaluation is not implemented.");
            cells[name] = cell;
        }
    }

    internal static string RowIndex(string? value, int position = 0) => value == null ? position.ToString(CultureInfo.InvariantCulture) : Id(value) ?? value;
    /// <summary>Detects native deletion before projecting rows or loading typed text caches.</summary>
    internal static bool HasDeletion(XElement? section) => Deleted(section) ||
        section?.Elements().Any(row => row.Name.LocalName == "Row" && Deleted(row)) == true;
    private static XElement PlainText(string? text) => new(XName.Get("Text", VisioDocument.VisioNamespace), text ?? string.Empty);
    private static bool Deleted(XElement? element) => (string?)element?.Attribute("Del") is "1" or "true";
    private static string? Id(string? value) => uint.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out uint id)
        ? id.ToString(CultureInfo.InvariantCulture) : null;
    private void Warn(string key, string message) {
        if (!_reported.Add(key) || _diagnostics == null) return;
        if (_diagnostics.Any(d => d.Code == DiagnosticCode && d.Message == message && d.Source == _source)) return;
        _diagnostics.Add(new OfficeImageExportDiagnostic(OfficeImageExportDiagnosticSeverity.Warning, DiagnosticCode,
            message, _source, OfficeConversionLossKind.Approximation));
    }
}
