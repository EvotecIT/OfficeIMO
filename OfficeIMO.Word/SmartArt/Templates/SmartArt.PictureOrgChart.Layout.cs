using DocumentFormat.OpenXml.Packaging;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;

namespace OfficeIMO.Word.SmartArt.Templates {
    /// Minimal strongly-typed Layout for Picture Organization Chart SmartArt.
    internal static class SmartArtPictureOrgChartLayout {
        internal static void PopulateLayout(DiagramLayoutDefinitionPart part) {
            var layout = new Dgm.LayoutDefinition { UniqueId = "urn:microsoft.com/office/officeart/2005/8/layout/pictureorgchart" };
            layout.AddNamespaceDeclaration("dgm", "http://schemas.openxmlformats.org/drawingml/2006/diagram");
            layout.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            layout.Append(new Dgm.Title { Val = "" });
            layout.Append(new Dgm.Description { Val = "" });

            var cats = new Dgm.CategoryList();
            cats.Append(new Dgm.Category { Type = "hierarchy", Priority = (UInt32Value)700U });
            layout.Append(cats);

            var layoutNode = new Dgm.LayoutNode { Name = "pictureOrgChart" };

            var shape = new Dgm.Shape { Blip = "" };
            shape.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            shape.Append(new Dgm.AdjustList());
            layoutNode.Append(shape);

            var forEach = new Dgm.ForEach {
                Name = "nodes",
                Axis = new ListValue<EnumValue<Dgm.AxisValues>> { InnerText = "ch" },
                PointType = new ListValue<EnumValue<Dgm.ElementValues>> { InnerText = "node" }
            };

            var node = new Dgm.LayoutNode { Name = "node" };
            var frame = new Dgm.Shape { Type = "roundRect", Blip = "" };
            frame.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            frame.Append(new Dgm.AdjustList());
            node.Append(frame);

            var picture = new Dgm.Shape { Type = "rect", Blip = "" };
            picture.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            picture.Append(new Dgm.AdjustList());
            node.Append(picture);

            forEach.Append(node);
            layoutNode.Append(forEach);

            layout.Append(layoutNode);
            part.LayoutDefinition = layout;
        }
    }
}
