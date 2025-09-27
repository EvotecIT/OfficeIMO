using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;

namespace OfficeIMO.Word.SmartArt.Templates {
    /// Minimal strongly-typed Data for Hierarchy SmartArt (root with two child placeholders).
    internal static class SmartArtHierarchyData {
        private const string LayoutId = "urn:microsoft.com/office/officeart/2005/8/layout/hierarchy1";

        internal static void PopulateData(DiagramDataPart part) {
            var model = new Dgm.DataModel();
            model.AddNamespaceDeclaration("dgm", "http://schemas.openxmlformats.org/drawingml/2006/diagram");
            model.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            string NewId() => "{" + System.Guid.NewGuid().ToString().ToUpper() + "}";

            var docId = NewId();
            var rootId = NewId();
            var child1Id = NewId();
            var child2Id = NewId();

            var pts = new Dgm.PointList();

            var docPt = new Dgm.Point { ModelId = docId, Type = Dgm.PointValues.Document };
            var docProps = new Dgm.PropertySet {
                LayoutTypeId = LayoutId,
                LayoutCategoryId = "hierarchy",
                QuickStyleTypeId = "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1",
                QuickStyleCategoryId = "simple",
                ColorType = "urn:microsoft.com/office/officeart/2005/8/colors/accent1_2",
                ColorCategoryId = "accent1",
                Placeholder = false
            };
            docPt.Append(docProps);
            docPt.Append(new Dgm.ShapeProperties());
            docPt.Append(CreateEmptyTextBody());

            var rootPt = new Dgm.Point { ModelId = rootId };
            rootPt.Append(new Dgm.PropertySet { Placeholder = true, PlaceholderText = "[Manager]" });
            rootPt.Append(new Dgm.ShapeProperties());
            rootPt.Append(CreateEmptyTextBody());

            var child1Pt = new Dgm.Point { ModelId = child1Id };
            child1Pt.Append(new Dgm.PropertySet { Placeholder = true, PlaceholderText = "[Employee 1]" });
            child1Pt.Append(new Dgm.ShapeProperties());
            child1Pt.Append(CreateEmptyTextBody());

            var child2Pt = new Dgm.Point { ModelId = child2Id };
            child2Pt.Append(new Dgm.PropertySet { Placeholder = true, PlaceholderText = "[Employee 2]" });
            child2Pt.Append(new Dgm.ShapeProperties());
            child2Pt.Append(CreateEmptyTextBody());

            pts.Append(docPt);
            pts.Append(rootPt);
            pts.Append(child1Pt);
            pts.Append(child2Pt);

            var cxns = new Dgm.ConnectionList();
            cxns.Append(new Dgm.Connection { ModelId = NewId(), SourceId = docId, DestinationId = rootId, SourcePosition = 0U, DestinationPosition = 0U });
            cxns.Append(new Dgm.Connection { ModelId = NewId(), SourceId = rootId, DestinationId = child1Id, SourcePosition = 0U, DestinationPosition = 0U });
            cxns.Append(new Dgm.Connection { ModelId = NewId(), SourceId = rootId, DestinationId = child2Id, SourcePosition = 1U, DestinationPosition = 0U });

            model.Append(pts);
            model.Append(cxns);
            model.Append(new Dgm.Background());
            model.Append(new Dgm.Whole());

            using var ms = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(model.OuterXml));
            part.FeedData(ms);
        }

        private static Dgm.TextBody CreateEmptyTextBody() {
            var text = new Dgm.TextBody();
            text.Append(new A.BodyProperties());
            text.Append(new A.ListStyle());
            var para = new A.Paragraph();
            para.Append(new A.EndParagraphRunProperties { Language = "en-US" });
            text.Append(para);
            return text;
        }
    }
}
