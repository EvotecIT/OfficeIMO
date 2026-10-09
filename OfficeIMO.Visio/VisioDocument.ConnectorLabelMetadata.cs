using System.Globalization;

namespace OfficeIMO.Visio;

public partial class VisioDocument {
    private const string ConnectorLabelPlacementPropName = "OfficeIMOConnectorLabelPlacement";

    // A reserved Shape Data row follows the existing OfficeIMOOriginalId convention.
    // Native text cells carry the current frame; this row retains only API anchor intent.
    private static bool RestoreConnectorLabelAnchor(VisioConnector connector) {
        if (connector.LabelPlacement == null || !connector.Data.TryGetValue(ConnectorLabelPlacementPropName, out string? value)) return false;
        string[] fields = value.Split(';');
        VisioConnectorLabelPlacement placement = connector.LabelPlacement;
        if (value == "v1;page") placement.AnchorKind = VisioConnectorLabelAnchorKind.Page;
        else if (value == "v1;native;page-angle") {
            placement.AnchorKind = VisioConnectorLabelAnchorKind.Native; placement.KeepPageAngle = true;
        }
        else if (fields.Length == 5 && fields[0] == "v1" && fields[1] == "path" &&
            TryLabelNumber(fields[2], out double position) && TryLabelNumber(fields[3], out double x) && TryLabelNumber(fields[4], out double y)) {
            placement.Position = position; placement.OffsetX = x; placement.OffsetY = y;
            placement.AnchorKind = VisioConnectorLabelAnchorKind.Path;
        } else return false; // Unknown producer data remains ordinary, preserved Shape Data.
        connector.Data.Remove(ConnectorLabelPlacementPropName);
        for (int i = connector.ShapeData.Count - 1; i >= 0; i--)
            if (connector.ShapeData[i].Name == ConnectorLabelPlacementPropName) connector.ShapeData.RemoveAt(i);
        for (int i = connector.PreservedDataRows.Count - 1; i >= 0; i--)
            if ((string?)connector.PreservedDataRows[i].Attribute("N") == ConnectorLabelPlacementPropName) connector.PreservedDataRows.RemoveAt(i);
        return true;
    }

    private static IDictionary<string, string> GetConnectorLabelData(VisioConnector connector) {
        VisioConnectorLabelPlacement? placement = connector.LabelPlacement;
        if (placement == null) return connector.Data;
        bool pageAngle = placement.KeepPageAngle || placement.NativeFrame != null &&
            (connector.TextStyle?.TextAngle ?? 0) != placement.NativeFrame.Angle;
        if (placement.AnchorKind == VisioConnectorLabelAnchorKind.Native && !pageAngle) return connector.Data;
        if (connector.Data.ContainsKey(ConnectorLabelPlacementPropName) || connector.ShapeData.Any(row => row.Name == ConnectorLabelPlacementPropName))
            throw new NotSupportedException("The reserved OfficeIMOConnectorLabelPlacement Shape Data name is already in use.");
        string value = placement.AnchorKind == VisioConnectorLabelAnchorKind.Native ? "v1;native;page-angle" :
            placement.AnchorKind == VisioConnectorLabelAnchorKind.Page ? "v1;page" :
            "v1;path;" + placement.Position.ToString("R", CultureInfo.InvariantCulture) + ";" +
            placement.OffsetX.ToString("R", CultureInfo.InvariantCulture) + ";" + placement.OffsetY.ToString("R", CultureInfo.InvariantCulture);
        return new Dictionary<string, string>(connector.Data) { [ConnectorLabelPlacementPropName] = value };
    }

    private static bool TryLabelNumber(string value, out double result) =>
        double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out result) && !double.IsNaN(result) && !double.IsInfinity(result);
}
