using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

public partial class VisioPage {
    /// <summary>Adds a connector between free page points using the page default unit unless a unit is supplied.</summary>
    public VisioConnector AddConnector(string id, OfficePoint start, OfficePoint end, ConnectorKind kind = ConnectorKind.Straight, VisioMeasurementUnit? unit = null) {
        VisioMeasurementUnit measurement = unit ?? DefaultUnit;
        var connector = new VisioConnector(id, new OfficePoint(start.X.ToInches(measurement), start.Y.ToInches(measurement)),
            new OfficePoint(end.X.ToInches(measurement), end.Y.ToInches(measurement))) { Kind = kind };
        Connectors.Add(connector);
        return connector;
    }
}
