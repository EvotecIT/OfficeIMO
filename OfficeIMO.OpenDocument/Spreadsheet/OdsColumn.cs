namespace OfficeIMO.OpenDocument;

/// <summary>A sparse ODS column definition run.</summary>
public sealed class OdsColumnRun {
    internal OdsColumnRun(OdsDocument document, XElement element, long startColumn, long repeatCount) {
        StartColumn = startColumn; RepeatCount = repeatCount;
        Hidden = (string?)element.Attribute(OdfNamespaces.Table + "visibility") == "collapse";
        StyleName = (string?)element.Attribute(OdfNamespaces.Table + "style-name");
    }
    /// <summary>Zero-based first logical column.</summary>
    public long StartColumn { get; }
    /// <summary>Number of logical columns represented by this run.</summary>
    public long RepeatCount { get; }
    /// <summary>Whether the run is hidden.</summary>
    public bool Hidden { get; }
    /// <summary>Referenced column style name.</summary>
    public string? StyleName { get; }
}

/// <summary>An editable ODS column definition after sparse run splitting.</summary>
public sealed class OdsColumn {
    private readonly OdsDocument _document;
    private readonly XElement _element;
    internal OdsColumn(OdsDocument document, XElement element) { _document = document; _element = element; }
    /// <summary>Whether this column is hidden.</summary>
    public bool Hidden {
        get => (string?)_element.Attribute(OdfNamespaces.Table + "visibility") == "collapse";
        set { _element.SetAttributeValue(OdfNamespaces.Table + "visibility", value ? "collapse" : null); Dirty(); }
    }
    /// <summary>Referenced column style name.</summary>
    public string? StyleName {
        get => (string?)_element.Attribute(OdfNamespaces.Table + "style-name");
        set { _element.SetAttributeValue(OdfNamespaces.Table + "style-name", value); Dirty(); }
    }
    /// <summary>Explicit column width.</summary>
    public OdfLength? Width {
        get {
            OdfStyle? style = StyleName == null ? null : _document.Styles.Find(OdfStyleFamily.TableColumn, StyleName);
            string? lexical = (string?)style?.Element.Element(OdfNamespaces.Style + "table-column-properties")?.Attribute(OdfNamespaces.Style + "column-width");
            return lexical == null ? (OdfLength?)null : OdfLength.Parse(lexical);
        }
        set {
            OdfStyle style = _document.Styles.EnsureAutomaticStyle(_element, OdfNamespaces.Table + "style-name", OdfStyleFamily.TableColumn, "ofC");
            style.SetProperty(OdfNamespaces.Style + "table-column-properties", OdfNamespaces.Style + "column-width", value?.ToString());
        }
    }
    private void Dirty() => _document.MarkPartDirty("content.xml");
}
