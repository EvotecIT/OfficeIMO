namespace OfficeIMO.OpenDocument;

/// <summary>A sparse ODS row run.</summary>
public sealed class OdsRowRun {
    private readonly OdsDocument _document;
    private readonly XElement _element;
    internal OdsRowRun(OdsDocument document, XElement element, long startRow, long repeatCount) {
        _document = document; _element = element; StartRow = startRow; RepeatCount = repeatCount;
    }
    /// <summary>Zero-based first logical row.</summary>
    public long StartRow { get; }
    /// <summary>Number of logical rows represented by the run.</summary>
    public long RepeatCount { get; }
    /// <summary>Prototype cell runs shared by every logical row in this run.</summary>
    public IReadOnlyList<OdsCellRun> CellRuns {
        get {
            var runs = new List<OdsCellRun>();
            long start = 0;
            foreach (XElement cell in OdsSheet.CellElements(_element)) {
                long count = OdsRepeatModel.Read(cell, OdfNamespaces.Table + "number-columns-repeated");
                runs.Add(new OdsCellRun(_document, cell, start, count));
                start = checked(start + count);
            }
            return runs;
        }
    }
    /// <summary>Whether the prototype row is hidden.</summary>
    public bool Hidden => (string?)_element.Attribute(OdfNamespaces.Table + "visibility") == "collapse";
    /// <summary>Referenced prototype row style.</summary>
    public string? StyleName => (string?)_element.Attribute(OdfNamespaces.Table + "style-name");
    /// <summary>Explicit prototype row height.</summary>
    public OdfLength? Height => new OdsRow(_document, _element).Height;
}

/// <summary>An editable ODS row after sparse run splitting.</summary>
public sealed class OdsRow {
    private readonly OdsDocument _document;
    private readonly XElement _element;
    internal OdsRow(OdsDocument document, XElement element) { _document = document; _element = element; }
    /// <summary>Whether this row is hidden.</summary>
    public bool Hidden {
        get => (string?)_element.Attribute(OdfNamespaces.Table + "visibility") == "collapse";
        set { _element.SetAttributeValue(OdfNamespaces.Table + "visibility", value ? "collapse" : null); Dirty(); }
    }
    /// <summary>Referenced row style name.</summary>
    public string? StyleName {
        get => (string?)_element.Attribute(OdfNamespaces.Table + "style-name");
        set { _element.SetAttributeValue(OdfNamespaces.Table + "style-name", value); Dirty(); }
    }
    /// <summary>Explicit row height.</summary>
    public OdfLength? Height {
        get {
            OdfStyle? style = StyleName == null ? null : _document.Styles.Find(OdfStyleFamily.TableRow, StyleName);
            string? lexical = (string?)style?.Element.Element(OdfNamespaces.Style + "table-row-properties")?.Attribute(OdfNamespaces.Style + "row-height");
            return lexical == null ? (OdfLength?)null : OdfLength.Parse(lexical);
        }
        set {
            OdfStyle style = _document.Styles.EnsureAutomaticStyle(_element, OdfNamespaces.Table + "style-name", OdfStyleFamily.TableRow, "ofR");
            style.SetProperty(OdfNamespaces.Style + "table-row-properties", OdfNamespaces.Style + "row-height", value?.ToString());
        }
    }
    private void Dirty() => _document.MarkPartDirty("content.xml");
}
