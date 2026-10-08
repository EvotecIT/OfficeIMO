using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

/// <summary>Prepares detached resize frames and geometry before applying them to the existing object graph.</summary>
internal static partial class VisioShapeResizing {
    private sealed class Node {
        internal VisioShape Source = null!;
        internal VisioShape Candidate = null!;
        internal double X, Y;
    }

    internal static void Resize(VisioPage page, VisioShape shape, double width, double height) {
        if (!Finite(width) || width <= 0) throw new ArgumentOutOfRangeException(nameof(width), "Width must be finite and positive.");
        if (!Finite(height) || height <= 0) throw new ArgumentOutOfRangeException(nameof(height), "Height must be finite and positive.");
        if (!Finite(shape.Width) || !Finite(shape.Height) || shape.Width <= 0 || shape.Height <= 0)
            throw new NotSupportedException("Resizing requires finite positive original shape dimensions.");
        if (width == shape.Width && height == shape.Height) return;
        var nodes = new Dictionary<VisioShape, Node>();
        VisioShape candidate = Prepare(shape, width / shape.Width, height / shape.Height, shape.PinX, shape.PinY, nodes);
        candidate.Parent = shape.Parent;
        ValidateConnectors(page, nodes);
        foreach (Node node in nodes.Values) Apply(node);
    }

    private static VisioShape Prepare(VisioShape source, double x, double y, double pinX, double pinY,
        IDictionary<VisioShape, Node> nodes, IReadOnlyDictionary<VisioShape, VisioShape>? artwork = null,
        IReadOnlyDictionary<string, string>? geometryReferences = null) {
        var target = new VisioShape(source.Id) {
            Width = source.Width * x, Height = source.Height * y,
            PinX = pinX, PinY = pinY, LocPinX = source.LocPinX * x, LocPinY = source.LocPinY * y,
            Angle = source.Angle, Master = source.Master, MasterShape = source.MasterShape,
            TextStyle = source.TextStyle?.Clone(), NativeCellMetadata = source.NativeCellMetadata?.Clone()
        };
        foreach (double value in new[] { x, y, target.Width, target.Height, pinX, pinY, target.LocPinX, target.LocPinY, target.Angle })
            if (!Finite(value)) throw new NotSupportedException("Resizing requires finite shape frames and scale factors.");
        if (x <= 0 || y <= 0 || target.Width < 0 || target.Height < 0)
            throw new NotSupportedException("Resizing does not support negative shape dimensions.");
        nodes.Add(source, new Node { Source = source, Candidate = target, X = x, Y = y });
        foreach (XElement cell in source.PreservedCellElements) target.PreservedCellElements.Add(new XElement(cell));
        foreach (var point in source.ConnectionPoints) {
            if (!Finite(point.X * x) || !Finite(point.Y * y)) throw new NotSupportedException("Resizing requires finite connection points.");
            target.ConnectionPoints.Add(new VisioConnectionPoint(point.X * x, point.Y * y, point.DirX, point.DirY));
        }
        ScaleTextFrame(source, target, x, y);
        VisioShape geometrySource = source;
        double geometryX = x, geometryY = y;
        if (artwork != null) {
            geometrySource = artwork[source];
            if (geometrySource.PreservedGeometrySections.Count > 0 || VisioForeignImage.IsForeign(geometrySource)) {
                if (!Finite(geometrySource.Width) || !Finite(geometrySource.Height) || geometrySource.Width <= 0 || geometrySource.Height <= 0)
                    throw new NotSupportedException("Replacement artwork requires finite positive master dimensions.");
                geometryX = target.Width / geometrySource.Width; geometryY = target.Height / geometrySource.Height;
            }
            target.NativeCellMetadata ??= VisioNativeCellMetadata.Empty();
            target.NativeCellMetadata.ReplaceArtworkState(geometrySource.NativeCellMetadata, geometryReferences);
        } else if (source.PreservedGeometrySections.Count == 0 && (source.MasterShape ?? source.Master?.Shape) is VisioShape master && master.PreservedGeometrySections.Count > 0) {
            if (master.Width <= 0 || master.Height <= 0) throw new NotSupportedException("Inherited geometry requires positive master dimensions.");
            geometrySource = master; geometryX = target.Width / master.Width; geometryY = target.Height / master.Height;
        }
        foreach (XElement section in geometrySource.PreservedGeometrySections) target.PreservedGeometrySections.Add(new XElement(section));
        VisioGeometryScaling.Scale(target, geometryX, geometryY, geometrySource);
        if (artwork != null) {
            foreach (XElement cell in target.PreservedCellElements.Where(IsImagePlacementCell).ToArray()) target.PreservedCellElements.Remove(cell);
            VisioForeignImage.ScalePlacement(geometrySource, target, geometryX, geometryY);
            if (geometryReferences != null) {
                // Only new artwork uses the replacement master's Sheet IDs. Existing
                // instance data, text and formulas already refer to the page graph.
                foreach (XAttribute formula in target.PreservedGeometrySections.SelectMany(section => section.DescendantsAndSelf()).Attributes("F")
                    .Concat(target.PreservedCellElements.Where(IsImagePlacementCell).Attributes("F")))
                    formula.Value = VisioShapeFormulaReferences.Rewrite(formula.Value, geometryReferences)!;
            }
        } else VisioForeignImage.ScalePlacement(source, target, x, y);
        foreach (VisioShape child in source.Children) {
            (double childX, double childY) = FrameScale(child.Angle, x, y);
            target.Children.Add(Prepare(child, childX, childY, child.PinX * x, child.PinY * y, nodes, artwork, geometryReferences));
        }
        return target;
    }

    private static void Apply(Node node) {
        VisioShape shape = node.Source, target = node.Candidate;
        shape.PinX = target.PinX; shape.PinY = target.PinY;
        shape.Width = target.Width; shape.Height = target.Height;
        shape.LocPinX = target.LocPinX; shape.LocPinY = target.LocPinY;
        shape.HasExplicitWidth = true; shape.HasExplicitHeight = true;
        for (int i = 0; i < shape.ConnectionPoints.Count; i++) {
            shape.ConnectionPoints[i].X = target.ConnectionPoints[i].X;
            shape.ConnectionPoints[i].Y = target.ConnectionPoints[i].Y;
        }
        if (shape.TextStyle is VisioTextStyle style && target.TextStyle is VisioTextStyle resized) {
            style.TextPinX = resized.TextPinX; style.TextPinY = resized.TextPinY;
            style.TextWidth = resized.TextWidth; style.TextHeight = resized.TextHeight;
            style.TextLocPinX = resized.TextLocPinX; style.TextLocPinY = resized.TextLocPinY;
        }
        shape.PreservedGeometrySections.Clear();
        foreach (XElement section in target.PreservedGeometrySections) shape.PreservedGeometrySections.Add(section);
        shape.PreservedCellElements.Clear();
        foreach (XElement cell in target.PreservedCellElements) shape.PreservedCellElements.Add(cell);
        // Loaded shapes retain a separate ordered XML view of unmodeled cells.
        // Keep those entries on the same resized values as the primary collection.
        for (int i = 0; i < shape.PreservedShapeChildren.Count; i++) {
            XElement? original = shape.PreservedShapeChildren[i].RawElement;
            if (original?.Name.LocalName != "Cell") continue;
            XElement? resizedCell = target.PreservedCellElements.FirstOrDefault(c => (string?)c.Attribute("N") == (string?)original.Attribute("N"));
            if (resizedCell != null) shape.PreservedShapeChildren[i] = new VisioShape.PreservedShapeChildEntry(new XElement(resizedCell));
            else if (IsImagePlacementCell(original)) { shape.PreservedShapeChildren.RemoveAt(i); i--; }
        }
        shape.NativeCellMetadata = target.NativeCellMetadata;
    }

    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private static bool IsImagePlacementCell(XElement cell) => (string?)cell.Attribute("N") is "ImgOffsetX" or "ImgOffsetY" or "ImgWidth" or "ImgHeight";
}
