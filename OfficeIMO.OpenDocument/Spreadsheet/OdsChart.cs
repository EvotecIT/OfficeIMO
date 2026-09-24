namespace OfficeIMO.OpenDocument;

/// <summary>One embedded ODS chart and its source-cell references. Chart styling remains preserved package XML.</summary>
public sealed class OdsChart {
    private OdsChart(string name, string chartClass, string? title, string? categoriesAddress,
        IReadOnlyList<OdsChartSeries> series, bool isStacked, bool isPercentage, bool isThreeDimensional,
        bool? verticalBars, OdfRect bounds, long? anchorRow, long? anchorColumn) {
        Name = name;
        ChartClass = chartClass;
        Title = title;
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
            string? title = chartElement.Element(chart + "title")?.Element(OdfNamespaces.Text + "p")?.Value;
            string[] categoryAddresses = plot.Elements(chart + "axis")
                .Select(axis => (string?)axis.Element(chart + "categories")?.Attribute(OdfNamespaces.Table + "cell-range-address"))
                .Where(value => !string.IsNullOrWhiteSpace(value)).Select(value => value!).ToArray();
            string? categories = categoryAddresses.Length == 1 ? categoryAddresses[0] : null;
            XElement[] seriesElements = plot.Elements(chart + "series").ToArray();
            var series = seriesElements.Select(element => new OdsChartSeries(
                (string?)element.Attribute(chart + "values-cell-range-address") ?? string.Empty,
                (string?)element.Attribute(chart + "label-cell-address"),
                NormalizeChartClass(element, (string?)element.Attribute(chart + "class")))).ToArray();
            XDocument? stylesPart = document.Package.ContainsEntry(directory + "styles.xml")
                ? document.Package.GetXml(directory + "styles.xml") : null;
            XElement? Style(string? name) => FindChartStyle(part, name)
                ?? (stylesPart == null ? null : FindChartStyle(stylesPart, name));
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
                    vertical = null;
                    break;
                }
                vertical = parsed;
            }
            OdfRect bounds = new OdfRect(
                OdfLength.Parse((string?)frame.Attribute(OdfNamespaces.Svg + "x") ?? "0cm"),
                OdfLength.Parse((string?)frame.Attribute(OdfNamespaces.Svg + "y") ?? "0cm"),
                OdfLength.Parse((string?)frame.Attribute(OdfNamespaces.Svg + "width") ?? "0cm"),
                OdfLength.Parse((string?)frame.Attribute(OdfNamespaces.Svg + "height") ?? "0cm"));
            return new OdsChart((string?)frame.Attribute(OdfNamespaces.Draw + "name") ?? string.Empty,
                chartClass, title, categories, series, stacked, percentage, threeDimensional, vertical, bounds,
                anchorRow, anchorColumn);
        } catch (InvalidDataException) {
            return null;
        }
    }

    private static XElement? FindChartStyle(XDocument part, string? styleName) => styleName == null ? null :
        part.Root?.Elements()
            .Where(element => element.Name == OdfNamespaces.Office + "automatic-styles"
                || element.Name == OdfNamespaces.Office + "styles")
            .SelectMany(element => element.Elements(OdfNamespaces.Style + "style"))
            .FirstOrDefault(element => string.Equals((string?)element.Attribute(OdfNamespaces.Style + "name"), styleName, StringComparison.Ordinal));

    private static string? NormalizeChartClass(XElement owner, string? lexical) {
        if (lexical == null) return null;
        int separator = lexical.IndexOf(':');
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
    internal OdsChartSeries(string valuesAddress, string? labelAddress, string? chartClass) {
        ValuesAddress = valuesAddress;
        LabelAddress = labelAddress;
        ChartClass = chartClass;
    }

    /// <summary>ODF range containing the series values.</summary>
    public string ValuesAddress { get; }
    /// <summary>Optional ODF cell containing the series label.</summary>
    public string? LabelAddress { get; }
    /// <summary>Optional series-specific ODF chart class.</summary>
    public string? ChartClass { get; }
}
