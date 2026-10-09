using System;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

public partial class VisioConnector {
    private OfficePoint _freeStart;
    private OfficePoint _freeEnd;
    internal VisioConnectorAttachment? StartAttachment { get; set; }
    internal VisioConnectorAttachment? EndAttachment { get; set; }

    /// <summary>Creates a connector between free page points expressed in inches.</summary>
    public VisioConnector(string id, OfficePoint start, OfficePoint end) {
        Id = id;
        StartPoint = start;
        EndPoint = end;
        LineColor = OfficeColor.Black;
        LineWeight = 0.0138889;
        LinePattern = 1;
        Kind = ConnectorKind.Straight;
    }

    /// <summary>Resolved start in page inches. Setting this value detaches the source end.</summary>
    public OfficePoint StartPoint {
        get => VisioConnectorEndpoints.Resolve(this, true);
        set {
            ValidatePoint(value);
            _freeStart = value; From = null; FromConnectionPoint = null; StartAttachment = null;
            ClearEndpointPreservation(true);
        }
    }

    /// <summary>Resolved end in page inches. Setting this value detaches the target end.</summary>
    public OfficePoint EndPoint {
        get => VisioConnectorEndpoints.Resolve(this, false);
        set {
            ValidatePoint(value);
            _freeEnd = value; To = null; ToConnectionPoint = null; EndAttachment = null;
            ClearEndpointPreservation(false);
        }
    }

    internal OfficePoint GetFreePoint(bool start) => start ? _freeStart : _freeEnd;

    internal void ClearEndpointPreservation(bool start) {
        string prefix = start ? "Begin" : "End";
        PreservedEndpointCellElements.Remove(prefix + "X");
        PreservedEndpointCellElements.Remove(prefix + "Y");
        if (start) {
            StartAttachment = null; PreservedFromConnectionCell = null;
            PreservedBeginConnectAttributes.Clear(); PreservedBeginConnectAttributeOrder.Clear();
        } else {
            EndAttachment = null; PreservedToConnectionCell = null;
            PreservedEndConnectAttributes.Clear(); PreservedEndConnectAttributeOrder.Clear();
        }
    }

    private static void ValidatePoint(OfficePoint point) {
        if (double.IsNaN(point.X) || double.IsInfinity(point.X) || double.IsNaN(point.Y) || double.IsInfinity(point.Y))
            throw new ArgumentOutOfRangeException(nameof(point), "Connector coordinates must be finite.");
    }
}
