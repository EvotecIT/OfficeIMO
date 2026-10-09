using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdgShape {
    /// <summary>Native transform for a leaf shape. Use TransformChildren to transform a group.</summary>
    public override string? Transform {
        get => base.Transform;
        set {
            if (IsGroup) throw new NotSupportedException("ODF groups do not have a native transform attribute. Add children and call TransformChildren instead.");
            base.Transform = value;
        }
    }

    /// <summary>
    /// Applies an affine transform to existing group children in their parent coordinate space.
    /// Nested groups remain groups; free connector and line coordinates are baked. Attached connectors
    /// and unknown elements are rejected before any child changes. Connector labels permit translation only;
    /// baking free connector geometry preserves the declared stroke width rather than scaling it.
    /// </summary>
    public void TransformChildren(string transform) {
        if (!IsGroup) throw new InvalidOperationException("Only groups have child transforms.");
        OfficeTransform operation = OdfDrawingTransform.Parse(transform);
        var edits = new List<Action>();
        PrepareGroupTransform(this, OdfDrawingTransform.Parse(base.Transform).Then(operation), edits);
        foreach (Action edit in edits) edit();
        Dirty();
    }

    private static void PrepareGroupTransform(OdgShape group, OfficeTransform transform, List<Action> edits) {
        foreach (OdgShape child in group.Children) {
            OfficeTransform combined = OdfDrawingTransform.Parse(child.Transform).Then(transform);
            if (child.IsGroup) { PrepareGroupTransform(child, combined, edits); continue; }
            if (child.IsConnector) {
                XAttribute[] attributes = child.PrepareFreeConnectorTransform(combined);
                edits.Add(() => child.ApplyBakedConnectorRoute(attributes));
            } else if (child.ElementName == "line") {
                OfficePoint start = combined.TransformPoint(new OfficePoint(child.X1.ToPoints(), child.Y1.ToPoints()));
                OfficePoint end = combined.TransformPoint(new OfficePoint(child.X2.ToPoints(), child.Y2.ToPoints()));
                if (!Finite(start.X) || !Finite(start.Y) || !Finite(end.X) || !Finite(end.Y)) throw new InvalidDataException("Transformed endpoints must be finite.");
                edits.Add(() => {
                    child.Element.SetAttributeValue(OdfNamespaces.Svg + "x1", OdfLength.Points(start.X));
                    child.Element.SetAttributeValue(OdfNamespaces.Svg + "y1", OdfLength.Points(start.Y));
                    child.Element.SetAttributeValue(OdfNamespaces.Svg + "x2", OdfLength.Points(end.X));
                    child.Element.SetAttributeValue(OdfNamespaces.Svg + "y2", OdfLength.Points(end.Y));
                    child.Element.Attribute(OdfNamespaces.Draw + "transform")?.Remove();
                });
            } else {
                if (child.ElementName is not ("rect" or "ellipse" or "circle" or "path" or "polygon" or "polyline" or "frame" or "custom-shape"))
                    throw new NotSupportedException("TransformChildren cannot transform native " + child.ElementName + ".");
                string matrix = FormatNativeMatrix(combined);
                edits.Add(() => child.Element.SetAttributeValue(OdfNamespaces.Draw + "transform", matrix));
            }
        }
        edits.Add(() => group.Element.Attribute(OdfNamespaces.Draw + "transform")?.Remove());
    }

    private static string FormatNativeMatrix(OfficeTransform transform) {
        string F(double number) => number.ToString("R", CultureInfo.InvariantCulture);
        return $"matrix({F(transform.M11)} {F(transform.M12)} {F(transform.M21)} {F(transform.M22)} {F(transform.OffsetX)}pt {F(transform.OffsetY)}pt)";
    }
}
