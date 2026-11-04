using System;
using System.Collections.Generic;

namespace OfficeIMO.Visio;

internal static partial class VisioShapeResizing {
    /// <summary>Prepares replacement artwork and all selected frames before mutating any live shape.</summary>
    internal static Action PrepareMasterReplacement(VisioPage page, IReadOnlyList<VisioShape> roots,
        IReadOnlyDictionary<VisioShape, VisioShape> artwork, double? width, double? height,
        IReadOnlyDictionary<VisioShape, IReadOnlyDictionary<string, string>> references,
        IReadOnlyDictionary<VisioShape, VisioShape[]> addedChildren) {
        var nodes = new Dictionary<VisioShape, Node>();
        bool resized = false;
        foreach (VisioShape shape in roots) {
            double w = width ?? shape.WidthInches, h = height ?? shape.HeightInches;
            if (!Finite(w) || !Finite(h) || w <= 0 || h <= 0 || !Finite(shape.WidthInches) || !Finite(shape.HeightInches) || shape.WidthInches <= 0 || shape.HeightInches <= 0)
                throw new NotSupportedException("Master replacement requires finite positive root dimensions.");
            resized |= w != shape.WidthInches || h != shape.HeightInches;
            VisioShape candidate = Prepare(shape, w / shape.WidthInches, h / shape.HeightInches, shape.PinXInches, shape.PinYInches, nodes, artwork, references[shape]);
            candidate.Parent = shape.Parent;
            if (addedChildren.TryGetValue(shape, out VisioShape[]? children)) {
                VisioShape master = artwork[shape];
                if (!Finite(master.WidthInches) || !Finite(master.HeightInches) || master.WidthInches <= 0 || master.HeightInches <= 0)
                    throw new NotSupportedException("Replacement child trees require finite positive master frames.");
                double x = w / master.WidthInches, y = h / master.HeightInches;
                foreach (VisioShape child in children) {
                    (double cx, double cy) = FrameScale(child.Angle, x, y);
                    candidate.Children.Add(Prepare(child, cx, cy, child.PinXInches * x, child.PinYInches * y, nodes, artwork, references[shape]));
                }
            }
        }
        if (resized) ValidateConnectors(page, nodes);
        return () => { foreach (Node node in nodes.Values) Apply(node); };
    }
}
