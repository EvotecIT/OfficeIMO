using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

/// <summary>Chart forms supported by native ODS chart authoring.</summary>
public enum OdsChartType {
    /// <summary>Vertical clustered columns.</summary>
    Column,
    /// <summary>Horizontal clustered bars.</summary>
    Bar,
    /// <summary>Line series.</summary>
    Line,
    /// <summary>One circular pie series.</summary>
    Pie,
    /// <summary>One or more concentric doughnut rings.</summary>
    Doughnut
}

/// <summary>One embedded ODS chart and its source-cell references. Unsupported styling remains preserved package XML.</summary>
public sealed class OdsChart {
    private const int MaximumImportedChartSeries = 256;
    private OdsChart(string name, string chartClass, string? title, string? titleCellRangeAddress,
        string? categoriesAddress,
        IReadOnlyList<OdsChartSeries> series, bool isStacked, bool isPercentage, bool isThreeDimensional,
        bool? verticalBars, OdfRect bounds, long? anchorRow, long? anchorColumn) {
        Name = name;
        ChartClass = chartClass;
        Title = title;
        TitleCellRangeAddress = titleCellRangeAddress;
        CategoriesAddress = categoriesAddress;
        Series = series;
        IsStacked = isStacked;
        IsPercentage = isPercentage;
        IsThreeDimensional = isThreeDimensional;
        VerticalBars = verticalBars;
        Bounds = bounds;
        AnchorRow = anchorRow;
        AnchorColumn = anchorColumn;
    }

    /// <summary>Drawing name in the parent worksheet.</summary>
    public string Name { get; }
    /// <summary>ODF chart class, such as <c>chart:bar</c> or <c>chart:line</c>.</summary>
    public string ChartClass { get; }
    /// <summary>Chart title text when present.</summary>
    public string? Title { get; }
    /// <summary>Cell range supplying the title when the chart uses a referenced title.</summary>
    public string? TitleCellRangeAddress { get; }
    /// <summary>ODF category-cell range referenced by the chart.</summary>
    public string? CategoriesAddress { get; }
    /// <summary>Value ranges and optional label cells in chart order.</summary>
    public IReadOnlyList<OdsChartSeries> Series { get; }
    /// <summary>Whether the plot style requests stacked series.</summary>
    public bool IsStacked { get; }
    /// <summary>Whether the plot style requests percentage stacking.</summary>
    public bool IsPercentage { get; }
    /// <summary>Whether the plot style requests a three-dimensional chart.</summary>
    public bool IsThreeDimensional { get; }
    /// <summary>For bar charts, true means horizontal bars; false means vertical columns.</summary>
    public bool? VerticalBars { get; }
    /// <summary>Frame position and size in the parent worksheet.</summary>
    public OdfRect Bounds { get; }
    /// <summary>Zero-based row containing the chart frame, or null for a sheet-level shape.</summary>
    public long? AnchorRow { get; }
    /// <summary>Zero-based column containing the chart frame, or null for a sheet-level shape.</summary>
    public long? AnchorColumn { get; }

    internal static OdsChart? TryRead(OdsDocument document, XElement frame, long? anchorRow, long? anchorColumn) {
        XElement? objectElement = frame.Element(OdfNamespaces.Draw + "object");
        string? href = (string?)objectElement?.Attribute(OdfNamespaces.XLink + "href");
        if (string.IsNullOrWhiteSpace(href) || href!.Contains(":") || href.Contains("..") || href.StartsWith("/", StringComparison.Ordinal)) return null;
        string directory = OdfPackagePath.NormalizeHref(href);
        if (directory.Length == 0 || directory.Contains("..") || directory.StartsWith("/", StringComparison.Ordinal)) return null;
        if (!directory.EndsWith("/", StringComparison.Ordinal)) directory += "/";
        string partPath = directory + "content.xml";
        if (!document.Package.ContainsEntry(partPath)) return null;
        try {
            XDocument part = document.Package.GetXml(partPath);
            XNamespace chart = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
            XElement? chartElement = part.Root?.Element(OdfNamespaces.Office + "body")?
                .Element(OdfNamespaces.Office + "chart")?.Element(chart + "chart");
            XElement[] plots = chartElement?.Elements(chart + "plot-area").Take(2).ToArray() ?? Array.Empty<XElement>();
            if (chartElement == null || plots.Length != 1) return null;
            XElement plot = plots[0];
            string chartClass = NormalizeChartClass(chartElement,
                (string?)chartElement.Attribute(chart + "class")) ?? string.Empty;
            XElement? titleElement = chartElement.Element(chart + "title");
            string? titleCellRangeAddress = (string?)titleElement?
                .Attribute(OdfNamespaces.Table + "cell-range");
            XElement[] titleParagraphs = titleElement?
                .Elements(OdfNamespaces.Text + "p").ToArray() ?? Array.Empty<XElement>();
            string? title = titleCellRangeAddress != null || titleParagraphs.Length == 0 ? null :
                string.Join("\n", titleParagraphs.Select(OdfTextCodec.Read));
            string[] categoryAddresses = plot.Elements(chart + "axis")
                .Select(axis => (string?)axis.Element(chart + "categories")?.Attribute(OdfNamespaces.Table + "cell-range-address"))
                .Where(value => !string.IsNullOrWhiteSpace(value)).Select(value => value!).ToArray();
            string? categories = categoryAddresses.Length == 1 ? categoryAddresses[0] : null;
            // Bound native series before expanding compact repeated point styles.
            XElement[] seriesElements = plot.Elements(chart + "series")
                .Take(MaximumImportedChartSeries + 1).ToArray();
            if (seriesElements.Length > MaximumImportedChartSeries) return null;
            XDocument? stylesPart = document.Package.ContainsEntry(directory + "styles.xml")
                ? document.Package.GetXml(directory + "styles.xml") : null;
            Dictionary<string, XElement> chartStyles = IndexChartStyles(part, stylesPart);
            XElement? Style(string? name) => name != null && chartStyles.TryGetValue(name, out XElement? found)
                ? found : null;
            XElement? defaultStyle = FindChartDefaultStyle(part)
                ?? (stylesPart == null ? null : FindChartDefaultStyle(stylesPart));
            IReadOnlyDictionary<string, XElement> hatches = OdsChartPointStyles.IndexHatches(part, stylesPart);
            var series = seriesElements.Select(element => {
                IReadOnlyList<OfficeChartPointStyle?>? pointStyles =
                    OdsChartPointStyles.Read(element, Style, defaultStyle, hatches);
                bool unprojectedAppearance = pointStyles == null && element.Elements(chart + "data-point")
                    .Any(point => point.Attribute(chart + "style-name") != null) ||
                    OdsChartPointStyles.HasUnprojectedSeriesPieOffset(element, Style);
                return new OdsChartSeries(
                    (string?)element.Attribute(chart + "values-cell-range-address") ?? string.Empty,
                    (string?)element.Attribute(chart + "label-cell-address"),
                    NormalizeChartClass(element, (string?)element.Attribute(chart + "class")),
                    pointStyles, unprojectedAppearance);
            }).ToArray();
            var allProperties = new List<XElement>();
            bool AddStyleChain(string? name) {
                var visited = new HashSet<string>(StringComparer.Ordinal);
                var effective = new XElement(OdfNamespaces.Style + "chart-properties");
                while (!string.IsNullOrEmpty(name)) {
                    if (!visited.Add(name!) || visited.Count > 32) return false;
                    XElement? current = Style(name);
                    if (current == null || (string?)current.Attribute(OdfNamespaces.Style + "family") != "chart") return false;
                    XElement? properties = current.Element(OdfNamespaces.Style + "chart-properties");
                    if (properties != null) {
                        foreach (XAttribute attribute in properties.Attributes()) {
                            if (effective.Attribute(attribute.Name) == null)
                                effective.SetAttributeValue(attribute.Name, attribute.Value);
                        }
                    }
                    name = (string?)current.Attribute(OdfNamespaces.Style + "parent-style-name");
                }
                XElement? defaults = defaultStyle?.Element(OdfNamespaces.Style + "chart-properties");
                if (defaults != null) {
                    foreach (XAttribute attribute in defaults.Attributes()) {
                        if (effective.Attribute(attribute.Name) == null)
                            effective.SetAttributeValue(attribute.Name, attribute.Value);
                    }
                }
                if (effective.HasAttributes) allProperties.Add(effective);
                return true;
            }
            if (!AddStyleChain((string?)chartElement.Attribute(chart + "style-name"))
                || !AddStyleChain((string?)plot.Attribute(chart + "style-name"))) return null;
            foreach (XElement element in seriesElements) {
                if (!AddStyleChain((string?)element.Attribute(chart + "style-name"))) return null;
            }
            bool UnsafeFlag(string name) => allProperties.Any(element => {
                string? token = (string?)element.Attribute(chart + name);
                return token != null && (!OdfBoolean.TryParseXml(token, out bool value) || value);
            });
            bool stacked = UnsafeFlag("stacked");
            bool percentage = UnsafeFlag("percentage");
            bool threeDimensional = UnsafeFlag("three-dimensional") || UnsafeFlag("deep");
            bool? vertical = null;
            foreach (XElement element in allProperties) {
                string? token = (string?)element.Attribute(chart + "vertical");
                if (token == null) continue;
                if (!OdfBoolean.TryParseXml(token, out bool parsed)
                    || (vertical.HasValue && vertical.Value != parsed)) {
                    return null;
                }
                vertical = parsed;
            }
            OdfRect bounds = new OdfRect(
                OdfLength.Parse((string?)frame.Attribute(OdfNamespaces.Svg + "x") ?? "0cm"),
                OdfLength.Parse((string?)frame.Attribute(OdfNamespaces.Svg + "y") ?? "0cm"),
                OdfLength.Parse((string?)frame.Attribute(OdfNamespaces.Svg + "width") ?? "0cm"),
                OdfLength.Parse((string?)frame.Attribute(OdfNamespaces.Svg + "height") ?? "0cm"));
            return new OdsChart((string?)frame.Attribute(OdfNamespaces.Draw + "name") ?? string.Empty,
                chartClass, title, titleCellRangeAddress, categories, series, stacked, percentage,
                threeDimensional, vertical, bounds,
                anchorRow, anchorColumn);
        } catch (InvalidDataException) {
            return null;
        }
    }

    private static Dictionary<string, XElement> IndexChartStyles(XDocument content, XDocument? styles) {
        var index = new Dictionary<string, XElement>(StringComparer.Ordinal);
        void Add(XDocument part) {
            foreach (XElement definition in part.Root?.Elements()
                .Where(element => element.Name == OdfNamespaces.Office + "automatic-styles"
                    || element.Name == OdfNamespaces.Office + "styles")
                .SelectMany(element => element.Elements(OdfNamespaces.Style + "style"))
                ?? Enumerable.Empty<XElement>()) {
                string? name = (string?)definition.Attribute(OdfNamespaces.Style + "name");
                if ((string?)definition.Attribute(OdfNamespaces.Style + "family") == "chart" &&
                    name != null && !index.ContainsKey(name)) index.Add(name, definition);
            }
        }
        Add(content);
        if (styles != null) Add(styles);
        return index;
    }

    private static XElement? FindChartDefaultStyle(XDocument part) =>
        part.Root?.Elements()
            .Where(element => element.Name == OdfNamespaces.Office + "automatic-styles"
                || element.Name == OdfNamespaces.Office + "styles")
            .SelectMany(element => element.Elements(OdfNamespaces.Style + "default-style"))
            .FirstOrDefault(element => (string?)element.Attribute(OdfNamespaces.Style + "family") == "chart");

    private static string? NormalizeChartClass(XElement owner, string? lexical) {
        if (lexical == null) return null;
        int separator = lexical.IndexOf(':');
        if (separator < 0 && owner.GetDefaultNamespace() == OdfNamespaces.Chart)
            return "chart:" + lexical;
        if (separator <= 0 || separator == lexical.Length - 1 || lexical.IndexOf(':', separator + 1) >= 0)
            return lexical;
        string prefix = lexical.Substring(0, separator);
        return owner.GetNamespaceOfPrefix(prefix) == OdfNamespaces.Chart
            ? "chart:" + lexical.Substring(separator + 1)
            : lexical;
    }
}

/// <summary>Cell references for one embedded ODS chart series.</summary>
public sealed class OdsChartSeries {
    /// <summary>Creates a chart series from an ODF cell range and optional one-cell label address.</summary>
    public OdsChartSeries(string valuesAddress, string? labelAddress = null)
        : this(valuesAddress, labelAddress, null) { }

    internal OdsChartSeries(string valuesAddress, string? labelAddress, string? chartClass,
        IReadOnlyList<OfficeChartPointStyle?>? pointStyles = null, bool hasUnprojectedAppearance = false) {
        ValuesAddress = valuesAddress;
        LabelAddress = labelAddress;
        ChartClass = chartClass;
        PointStyles = pointStyles == null ? null : Array.AsReadOnly(pointStyles.ToArray());
        HasUnprojectedAppearance = hasUnprojectedAppearance;
    }

    /// <summary>ODF range containing the series values.</summary>
    public string ValuesAddress { get; }
    /// <summary>Optional ODF cell containing the series label.</summary>
    public string? LabelAddress { get; }
    /// <summary>Optional series-specific ODF chart class.</summary>
    public string? ChartClass { get; }
    /// <summary>Supported native per-point fill, hatch, and outline overrides, when present.</summary>
    public IReadOnlyList<OfficeChartPointStyle?>? PointStyles { get; }
    internal bool HasUnprojectedAppearance { get; }

    /// <summary>Returns a series with native per-point appearance overrides.</summary>
    public OdsChartSeries WithPointStyles(IReadOnlyList<OfficeChartPointStyle?>? styles) =>
        new OdsChartSeries(ValuesAddress, LabelAddress, ChartClass, styles);
}
