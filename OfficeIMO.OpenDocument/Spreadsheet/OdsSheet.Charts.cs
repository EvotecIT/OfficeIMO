using OfficeIMO.Spreadsheet;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdsSheet {
    /// <summary>Adds a bounded chart linked to existing ODS cells and anchored in this sheet.</summary>
    public OdsChart AddChart(OdsChartType type, string categoriesAddress,
        IReadOnlyList<OdsChartSeries> series, long anchorRow, long anchorColumn,
        OdfRect bounds, string? title = null, string? name = null) {
        if (!Enum.IsDefined(typeof(OdsChartType), type)) throw new ArgumentOutOfRangeException(nameof(type));
        if (anchorRow < 0) throw new ArgumentOutOfRangeException(nameof(anchorRow));
        if (anchorColumn < 0) throw new ArgumentOutOfRangeException(nameof(anchorColumn));
        if (series == null || series.Count < 1 || series.Count > 16)
            throw new ArgumentException("A chart requires one to sixteen series.", nameof(series));
        int pointCount = ValidateChartRange(categoriesAddress, nameof(categoriesAddress), singleCell: false);
        if (pointCount > 4096) throw new ArgumentOutOfRangeException(nameof(categoriesAddress));
        foreach (OdsChartSeries item in series) {
            if (item == null) throw new ArgumentException("Chart series cannot be null.", nameof(series));
            if (ValidateChartRange(item.ValuesAddress, nameof(series), singleCell: false) != pointCount)
                throw new ArgumentException("Each chart series must match the category count.", nameof(series));
            if (item.LabelAddress != null) ValidateChartRange(item.LabelAddress, nameof(series), singleCell: true);
        }
        if (!bounds.X.TryToPoints(out double x) || !bounds.Y.TryToPoints(out double y) ||
            !bounds.Width.TryToPoints(out double width) || !bounds.Height.TryToPoints(out double height) ||
            x < 0 || y < 0 || width <= 0 || height <= 0 || width > 1500 || height > 1500)
            throw new ArgumentOutOfRangeException(nameof(bounds), "Chart bounds require positive absolute sizes up to 1500 points.");
        if (title?.Length > 32767) throw new ArgumentOutOfRangeException(nameof(title));
        OdsCell anchor = Cell(anchorRow, anchorColumn);
        if (anchor.IsCovered) throw new InvalidOperationException("A chart cannot be anchored in a covered cell.");

        int ordinal = 1;
        string directory;
        do { directory = "Object " + ordinal++.ToString(CultureInfo.InvariantCulture) + "/"; }
        while (_document.Package.ContainsEntry(directory) || _document.Package.ContainsEntry(directory + "content.xml"));
        string chartClass = type == OdsChartType.Line ? "chart:line" : "chart:bar";
        string chartName = string.IsNullOrWhiteSpace(name) ? "Chart " + (ordinal - 1).ToString(CultureInfo.InvariantCulture) : name!;
        XNamespace chart = OdfNamespaces.Chart;
        XElement plot = new XElement(chart + "plot-area",
            new XElement(chart + "axis", new XAttribute(chart + "dimension", "x"),
                new XAttribute(chart + "name", "primary-x"),
                new XElement(chart + "categories", new XAttribute(OdfNamespaces.Table + "cell-range-address", categoriesAddress))),
            new XElement(chart + "axis", new XAttribute(chart + "dimension", "y"),
                new XAttribute(chart + "name", "primary-y")));
        foreach (OdsChartSeries item in series) {
            var output = new XElement(chart + "series",
                new XAttribute(chart + "values-cell-range-address", item.ValuesAddress),
                new XAttribute(chart + "class", chartClass),
                new XAttribute(chart + "attached-axis", "primary-y"),
                new XElement(chart + "data-point", new XAttribute(chart + "repeated", pointCount)));
            if (item.LabelAddress != null)
                output.SetAttributeValue(chart + "label-cell-address", item.LabelAddress);
            plot.Add(output);
        }
        var chartElement = new XElement(chart + "chart",
            new XAttribute(chart + "class", chartClass),
            new XAttribute(chart + "style-name", "ChartStyle"),
            new XAttribute(OdfNamespaces.Svg + "width", bounds.Width.ToString()),
            new XAttribute(OdfNamespaces.Svg + "height", bounds.Height.ToString()));
        if (!string.IsNullOrEmpty(title)) {
            var paragraph = new XElement(OdfNamespaces.Text + "p");
            OdfTextCodec.Append(paragraph, title);
            chartElement.Add(new XElement(chart + "title", paragraph));
        }
        chartElement.Add(plot);
        var root = new XElement(OdfNamespaces.Office + "document-content");
        OdfXmlCodec.AddStandardNamespaces(root);
        root.SetAttributeValue(XNamespace.Xmlns + "chart", chart.NamespaceName);
        root.SetAttributeValue(OdfNamespaces.Office + "version", _document.Version.ToToken());
        var properties = new XElement(OdfNamespaces.Style + "chart-properties",
            new XAttribute(chart + "stacked", "false"),
            new XAttribute(chart + "percentage", "false"),
            new XAttribute(chart + "three-dimensional", "false"));
        if (type != OdsChartType.Line) properties.SetAttributeValue(chart + "vertical", type == OdsChartType.Bar ? "true" : "false");
        root.Add(new XElement(OdfNamespaces.Office + "automatic-styles",
            new XElement(OdfNamespaces.Style + "style",
                new XAttribute(OdfNamespaces.Style + "name", "ChartStyle"),
                new XAttribute(OdfNamespaces.Style + "family", "chart"), properties,
                new XElement(OdfNamespaces.Style + "graphic-properties",
                    new XAttribute(OdfNamespaces.Draw + "fill", "none"),
                    new XAttribute(OdfNamespaces.Draw + "stroke", "none")))),
            new XElement(OdfNamespaces.Office + "body",
                new XElement(OdfNamespaces.Office + "chart", chartElement)));
        var part = new XDocument(new XDeclaration("1.0", "UTF-8", null), root);
        var frame = new XElement(OdfNamespaces.Draw + "frame",
            new XAttribute(OdfNamespaces.Draw + "name", chartName),
            new XAttribute(OdfNamespaces.Svg + "x", bounds.X.ToString()),
            new XAttribute(OdfNamespaces.Svg + "y", bounds.Y.ToString()),
            new XAttribute(OdfNamespaces.Svg + "width", bounds.Width.ToString()),
            new XAttribute(OdfNamespaces.Svg + "height", bounds.Height.ToString()),
            new XElement(OdfNamespaces.Draw + "object",
                new XAttribute(OdfNamespaces.XLink + "href", directory),
                new XAttribute(OdfNamespaces.XLink + "type", "simple"),
                new XAttribute(OdfNamespaces.XLink + "show", "embed"),
                new XAttribute(OdfNamespaces.XLink + "actuate", "onLoad")));
        _document.Package.AddOrReplaceEntry(directory, Array.Empty<byte>(), "application/vnd.oasis.opendocument.chart");
        _document.Package.AddOrReplaceEntry(directory + "content.xml", OdfXmlCodec.Save(part), "text/xml");
        anchor.Element.Add(frame);
        _document.MarkPartDirty("content.xml");
        return OdsChart.TryRead(_document, frame, anchorRow, anchorColumn)
            ?? throw new InvalidDataException("The authored chart could not be read from its ODS package part.");
    }

    private int ValidateChartRange(string address, string parameter, bool singleCell) {
        if (!SpreadsheetRangeReference.TryParse(address, SpreadsheetAddressDialect.OpenDocument,
            out SpreadsheetRangeReference? range) || !range!.Start.IsCell)
            throw new ArgumentException("Chart data requires an ODF cell address or one-dimensional range.", parameter);
        SpreadsheetCellReference start = range.Start;
        SpreadsheetCellReference end = range.End ?? start;
        if (!end.IsCell || string.IsNullOrEmpty(start.SheetName) ||
            (range.End != null && string.IsNullOrEmpty(end.SheetName) &&
                !string.Equals(start.SheetName, Name, StringComparison.Ordinal)) ||
            (end.SheetName != null && !string.Equals(start.SheetName, end.SheetName, StringComparison.Ordinal)) ||
            _document.GetSheet(start.SheetName!) == null ||
            start.Row!.Value > end.Row!.Value || start.Column!.Value > end.Column!.Value ||
            start.Row.Value != end.Row.Value && start.Column.Value != end.Column.Value)
            throw new ArgumentException("Chart data requires one existing sheet and a one-dimensional cell range.", parameter);
        long count = start.Row.Value == end.Row.Value
            ? end.Column.Value - start.Column.Value + 1L : end.Row.Value - start.Row.Value + 1;
        if (count < 1 || count > 4096 || singleCell && count != 1)
            throw new ArgumentOutOfRangeException(parameter, "Chart ranges are limited to 4096 cells; labels require one cell.");
        return checked((int)count);
    }
}
