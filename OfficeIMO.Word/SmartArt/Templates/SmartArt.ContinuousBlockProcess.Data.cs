using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;

namespace OfficeIMO.Word.SmartArt.Templates {
    /// Minimal strongly-typed Data for Continuous Block Process SmartArt (two-step process).
    internal static class SmartArtContinuousBlockProcessData {
        private const string LayoutId = "urn:microsoft.com/office/officeart/2005/8/layout/process6";

        internal static void PopulateData(DiagramDataPart part) {
            var model = new Dgm.DataModel();
            model.AddNamespaceDeclaration("dgm", "http://schemas.openxmlformats.org/drawingml/2006/diagram");
            model.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            string NewId() => "{" + System.Guid.NewGuid().ToString().ToUpper() + "}";

            var docId = NewId();
            var step1Id = NewId();
            var step2Id = NewId();

            var pts = new Dgm.PointList();

            var docPt = new Dgm.Point { ModelId = docId, Type = Dgm.PointValues.Document };
            docPt.Append(new Dgm.PropertySet {
                LayoutTypeId = LayoutId,
                LayoutCategoryId = "process",
                QuickStyleTypeId = "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1",
                QuickStyleCategoryId = "simple",
                ColorType = "urn:microsoft.com/office/officeart/2005/8/colors/accent1_2",
                ColorCategoryId = "accent1",
                Placeholder = false
            });
            docPt.Append(new Dgm.ShapeProperties());
            docPt.Append(CreateEmptyTextBody());

            var step1Pt = new Dgm.Point { ModelId = step1Id };
            step1Pt.Append(new Dgm.PropertySet { Placeholder = true, PlaceholderText = "[Step 1]" });
            step1Pt.Append(new Dgm.ShapeProperties());
            step1Pt.Append(CreateEmptyTextBody());

            var step2Pt = new Dgm.Point { ModelId = step2Id };
            step2Pt.Append(new Dgm.PropertySet { Placeholder = true, PlaceholderText = "[Step 2]" });
            step2Pt.Append(new Dgm.ShapeProperties());
            step2Pt.Append(CreateEmptyTextBody());

            pts.Append(docPt);
            pts.Append(step1Pt);
            pts.Append(step2Pt);

            var cxns = new Dgm.ConnectionList();
            cxns.Append(new Dgm.Connection { ModelId = NewId(), SourceId = docId, DestinationId = step1Id, SourcePosition = 0U, DestinationPosition = 0U });
            cxns.Append(new Dgm.Connection { ModelId = NewId(), SourceId = step1Id, DestinationId = step2Id, SourcePosition = 0U, DestinationPosition = 0U });

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
