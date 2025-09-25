using DocumentFormat.OpenXml.Packaging;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;

namespace OfficeIMO.Word.SmartArt.Templates {
    /// Minimal strongly-typed Layout for Cycle SmartArt (cycle2).
    internal static class SmartArtCycleLayout {
        internal static void PopulateLayout(DiagramLayoutDefinitionPart part) {
            var layout = new Dgm.LayoutDefinition { UniqueId = "urn:microsoft.com/office/officeart/2005/8/layout/cycle2" };
            layout.AddNamespaceDeclaration("dgm", "http://schemas.openxmlformats.org/drawingml/2006/diagram");
            layout.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            layout.Append(new Dgm.Title { Val = "" });
            layout.Append(new Dgm.Description { Val = "" });

            var cats = new Dgm.CategoryList();
            cats.Append(new Dgm.Category { Type = "cycle", Priority = (UInt32Value)1000U });
            layout.Append(cats);

            var layoutNode = new Dgm.LayoutNode { Name = "cycle" };

            // Base shape for the diagram
            var shape = new Dgm.Shape { Blip = "" };
            shape.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            shape.Append(new Dgm.AdjustList());
            layoutNode.Append(shape);

            // Iterate child nodes as ellipses (cycle)
            var forEach = new Dgm.ForEach {
                Name = "nodes",
                Axis = new ListValue<EnumValue<Dgm.AxisValues>> { InnerText = "ch" },
                PointType = new ListValue<EnumValue<Dgm.ElementValues>> { InnerText = "node" }
            };
            var node = new Dgm.LayoutNode { Name = "node" };
            var nodeShape = new Dgm.Shape { Type = "ellipse", Blip = "" };
            nodeShape.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            nodeShape.Append(new Dgm.AdjustList());
            node.Append(nodeShape);
            forEach.Append(node);
            layoutNode.Append(forEach);

            layout.Append(layoutNode);
            part.LayoutDefinition = layout;
        }
    }
}

