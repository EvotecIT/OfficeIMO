using OfficeIMO.Spreadsheet;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdsDocument {
    /// <summary>Native data pilot tables in document order.</summary>
    public IReadOnlyList<OdsDataPilotTable> DataPilotTables => SpreadsheetBody
        .Element(OdfNamespaces.Table + "data-pilot-tables")?
        .Elements(OdfNamespaces.Table + "data-pilot-table")
        .Select(element => new OdsDataPilotTable(this, element)).ToList()
        ?? (IReadOnlyList<OdsDataPilotTable>)Array.Empty<OdsDataPilotTable>();

    /// <summary>Adds a local-range ODF data pilot table; the caller supplies the complete output range.</summary>
    public OdsDataPilotTable AddDataPilotTable(string name, string sourceRangeAddress, string targetRangeAddress) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Data pilot name cannot be empty.", nameof(name));
        SpreadsheetRangeReference source = OdsDataPilotTable.ParseLocalRange(sourceRangeAddress, nameof(sourceRangeAddress));
        SpreadsheetRangeReference target = OdsDataPilotTable.ParseLocalRange(targetRangeAddress, nameof(targetRangeAddress));
        if (GetSheet(source.Start.SheetName!) == null) throw new ArgumentException("Source worksheet does not exist.", nameof(sourceRangeAddress));
        if (GetSheet(target.Start.SheetName!) == null) throw new ArgumentException("Target worksheet does not exist.", nameof(targetRangeAddress));
        if (DataPilotTables.Any(table => string.Equals(table.Name, name, StringComparison.Ordinal))) {
            throw new InvalidOperationException($"A data pilot table named '{name}' already exists.");
        }
        XElement? container = SpreadsheetBody.Element(OdfNamespaces.Table + "data-pilot-tables");
        if (container == null) {
            container = new XElement(OdfNamespaces.Table + "data-pilot-tables");
            InsertSpreadsheetMetadata(container, OdfNamespaces.Table + "consolidation",
                OdfNamespaces.Table + "dde-links");
        }
        var element = new XElement(OdfNamespaces.Table + "data-pilot-table",
            new XAttribute(OdfNamespaces.Table + "name", name),
            new XAttribute(OdfNamespaces.Table + "target-range-address", targetRangeAddress),
            new XElement(OdfNamespaces.Table + "source-cell-range",
                new XAttribute(OdfNamespaces.Table + "cell-range-address", sourceRangeAddress)));
        container.Add(element);
        MarkPartDirty("content.xml");
        return new OdsDataPilotTable(this, element);
    }
}
