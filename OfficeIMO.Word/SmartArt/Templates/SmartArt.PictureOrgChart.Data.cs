using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;

namespace OfficeIMO.Word.SmartArt.Templates {
    /// Minimal strongly-typed Data for Picture Organization Chart SmartArt (manager with one direct report and picture placeholder).
    internal static class SmartArtPictureOrgChartData {
        private const string LayoutId = "urn:microsoft.com/office/officeart/2005/8/layout/pictureorgchart";

        internal static void PopulateData(DiagramDataPart part) {
            var model = new Dgm.DataModel();
            model.AddNamespaceDeclaration("dgm", "http://schemas.openxmlformats.org/drawingml/2006/diagram");
            model.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            string NewId() => "{" + System.Guid.NewGuid().ToString().ToUpper() + "}";

            var docId = NewId();
            var managerId = NewId();
            var reportId = NewId();
            var photoId = NewId();

            var pts = new Dgm.PointList();

            var docPt = new Dgm.Point { ModelId = docId, Type = Dgm.PointValues.Document };
            docPt.Append(new Dgm.PropertySet {
                LayoutTypeId = LayoutId,
                LayoutCategoryId = "hierarchy",
                QuickStyleTypeId = "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1",
                QuickStyleCategoryId = "simple",
                ColorType = "urn:microsoft.com/office/officeart/2005/8/colors/accent1_2",
                ColorCategoryId = "accent1",
                Placeholder = false
            });
            docPt.Append(new Dgm.ShapeProperties());
            docPt.Append(CreateEmptyTextBody());

            var managerPt = new Dgm.Point { ModelId = managerId };
            managerPt.Append(new Dgm.PropertySet { Placeholder = true, PlaceholderText = "[Name]" });
            managerPt.Append(new Dgm.ShapeProperties());
            managerPt.Append(CreateEmptyTextBody());

            var reportPt = new Dgm.Point { ModelId = reportId };
            reportPt.Append(new Dgm.PropertySet { Placeholder = true, PlaceholderText = "[Report]" });
            reportPt.Append(new Dgm.ShapeProperties());
            reportPt.Append(CreateEmptyTextBody());

            var photoPt = new Dgm.Point { ModelId = photoId };
            photoPt.Append(new Dgm.PropertySet { Placeholder = true, PlaceholderText = "[Picture]" });
            photoPt.Append(new Dgm.ShapeProperties());
            photoPt.Append(CreateEmptyTextBody());

            pts.Append(docPt);
            pts.Append(managerPt);
            pts.Append(reportPt);
            pts.Append(photoPt);

            var cxns = new Dgm.ConnectionList();
            cxns.Append(new Dgm.Connection { ModelId = NewId(), SourceId = docId, DestinationId = managerId, SourcePosition = 0U, DestinationPosition = 0U });
            cxns.Append(new Dgm.Connection { ModelId = NewId(), SourceId = managerId, DestinationId = reportId, SourcePosition = 0U, DestinationPosition = 0U });
            cxns.Append(new Dgm.Connection { ModelId = NewId(), Type = Dgm.ConnectionValues.PresentationOf, SourceId = reportId, DestinationId = photoId, SourcePosition = 0U, DestinationPosition = 0U });

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
