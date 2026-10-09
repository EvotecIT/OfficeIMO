using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

internal static partial class VisioShapeResizing {
    private static void ValidateConnectors(VisioPage page, IReadOnlyDictionary<VisioShape, Node> nodes) {
        foreach (VisioConnector source in page.Connectors) {
            bool fromChanged = source.From != null && nodes.ContainsKey(source.From);
            bool toChanged = source.To != null && nodes.ContainsKey(source.To);
            if (!fromChanged && !toChanged) continue;
            var candidate = new VisioConnector(source.Id, source.GetFreePoint(true), source.GetFreePoint(false)) {
                From = fromChanged ? nodes[source.From!].Candidate : source.From,
                To = toChanged ? nodes[source.To!].Candidate : source.To,
                FromConnectionPoint = ScalePoint(source.FromConnectionPoint, fromChanged ? nodes[source.From!] : null),
                ToConnectionPoint = ScalePoint(source.ToConnectionPoint, toChanged ? nodes[source.To!] : null),
                StartAttachment = source.StartAttachment, EndAttachment = source.EndAttachment, Kind = source.Kind
            };
            foreach (var point in source.Waypoints) candidate.Waypoints.Add(point);
            foreach (XElement section in source.PreservedGeometrySections) candidate.PreservedGeometrySections.Add(new XElement(section));
            if (!Finite(candidate.StartPoint.X) || !Finite(candidate.StartPoint.Y) || !Finite(candidate.EndPoint.X) || !Finite(candidate.EndPoint.Y))
                throw new NotSupportedException("Resizing requires finite attached connector endpoints.");
            if (source.NativeGeometry is VisioConnectorNativeGeometry geometry && geometry.AppliesTo(candidate))
                _ = geometry.CreateShape(candidate, requireCompleteRows: true);
        }
    }

    private static VisioConnectionPoint? ScalePoint(VisioConnectionPoint? point, Node? node) => point == null || node == null ? point :
        new VisioConnectionPoint(point.X * node.X, point.Y * node.Y, point.DirX, point.DirY);
}
