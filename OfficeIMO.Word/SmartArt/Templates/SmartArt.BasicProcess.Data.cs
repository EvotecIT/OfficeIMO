using DocumentFormat.OpenXml.Packaging;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;

namespace OfficeIMO.Word.SmartArt.Templates {
    /// Minimal strongly-typed Data for Basic Process SmartArt (one placeholder node).
    internal static class SmartArtBasicProcessData {
        internal static void PopulateData(DiagramDataPart part) {
            var root = new Dgm.DataModelRoot();
            root.AddNamespaceDeclaration("dgm", "http://schemas.openxmlformats.org/drawingml/2006/diagram");
            root.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            var model = new Dgm.DataModel();

            var pts = new Dgm.PointList();

            // Document node
            var docId = "{" + System.Guid.NewGuid().ToString().ToUpper() + "}";
            var docPt = new Dgm.Point { ModelId = docId, Type = Dgm.PointValues.Document };
            var docProps = new Dgm.PropertySet {
                LayoutTypeId = "urn:microsoft.com/office/officeart/2005/8/layout/default",
                LayoutCategoryId = "list",
                QuickStyleTypeId = "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1",
                QuickStyleCategoryId = "simple",
                ColorType = "urn:microsoft.com/office/officeart/2005/8/colors/accent1_2",
                ColorCategoryId = "accent1",
                Placeholder = false
            };
            docPt.Append(docProps);
            docPt.Append(new Dgm.ShapeProperties());
            var docText = new Dgm.TextBody();
            docText.Append(new A.BodyProperties());
            docText.Append(new A.ListStyle());
            var docPara = new A.Paragraph();
            docPara.Append(new A.EndParagraphRunProperties { Language = "en-US" });
            docText.Append(docPara);
            docPt.Append(docText);

            // One child placeholder node
            var childId = "{" + System.Guid.NewGuid().ToString().ToUpper() + "}";
            var childPt = new Dgm.Point { ModelId = childId };
            var childProps = new Dgm.PropertySet { Placeholder = true, PlaceholderText = "[Text]" };
            childPt.Append(childProps);
            childPt.Append(new Dgm.ShapeProperties());
            var childText = new Dgm.TextBody();
            childText.Append(new A.BodyProperties());
            childText.Append(new A.ListStyle());
            var childPara = new A.Paragraph();
            childPara.Append(new A.EndParagraphRunProperties { Language = "en-US" });
            childText.Append(childPara);
            childPt.Append(childText);

            pts.Append(docPt);
            pts.Append(childPt);

            // Connect doc -> child
            var cxns = new Dgm.ConnectionList();
            var cxnId = "{" + System.Guid.NewGuid().ToString().ToUpper() + "}";
            var cxn = new Dgm.Connection { ModelId = cxnId, SourceId = docId, DestinationId = childId, SourcePosition = 0U, DestinationPosition = 0U };
            cxns.Append(cxn);

            model.Append(pts);
            model.Append(cxns);
            model.Append(new Dgm.Background());
            model.Append(new Dgm.Whole());

            root.Append(model);
            part.DataModelRoot = root;
        }
    }
}

