using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Designs a slide master and layouts, fills placeholders natively and links shapes with glued connectors.
    /// </summary>
    public static class MastersLayoutsConnectorsPowerPoint {
        public static void Example_MastersLayoutsConnectors(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Masters, layouts and connectors");
            string filePath = Path.Combine(folderPath, "MastersLayoutsConnectors.pptx");
            string logoPath = Path.Combine(Directory.GetCurrentDirectory(), "Images", "EvotecLogo.png");
            string photoPath = Path.Combine(Directory.GetCurrentDirectory(), "Images", "Kulek.jpg");
            OfficeColor brand = OfficeColor.Parse("1F4E79");

            using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
            OfficeLength slideWidth = presentation.SlideSize.Width;
            OfficeLength slideHeight = presentation.SlideSize.Height;

            // Everything on the master shows on every slide.
            PowerPointSlideMaster master = presentation.SlideMasters[0];
            master.Name = "Contoso";
            master.AddShape(OfficePresetShapeType.Rectangle, OfficeLength.Points(0), OfficeLength.Points(0),
                slideWidth, OfficeLength.Points(10), "Brand band").Fill(brand);
            master.AddImage(logoPath, slideWidth - OfficeLength.Points(110), slideHeight - OfficeLength.Points(40),
                OfficeLength.Points(90), OfficeLength.Points(28));

            // A custom layout with a title, a picture placeholder and a content placeholder.
            PowerPointSlideLayout layout = master.AddLayout("Picture and content");
            layout.AddPlaceholder(PowerPointPlaceholderType.Title, new PowerPointLayoutBox(
                OfficeLength.Points(40), OfficeLength.Points(30), slideWidth - OfficeLength.Points(80), OfficeLength.Points(50)));
            layout.AddPlaceholder(PowerPointPlaceholderType.Picture, new PowerPointLayoutBox(
                OfficeLength.Points(40), OfficeLength.Points(100), OfficeLength.Points(260), OfficeLength.Points(300)));
            layout.AddPlaceholder(PowerPointPlaceholderType.Object, new PowerPointLayoutBox(
                OfficeLength.Points(330), OfficeLength.Points(100), slideWidth - OfficeLength.Points(370), OfficeLength.Points(300)));
            layout.AddPlaceholder(PowerPointPlaceholderType.SlideNumber, new PowerPointLayoutBox(
                OfficeLength.Points(40), slideHeight - OfficeLength.Points(36), OfficeLength.Points(60), OfficeLength.Points(20)));

            // A darker copy of the layout; slides using it can be moved back when it is removed.
            PowerPointSlideLayout dark = layout.Duplicate("Picture and content (dark)");
            dark.BackgroundColor = OfficeColor.Parse("1B1B1B");

            PowerPointSlide slide = presentation.AddSlide(layout);
            slide.AddTextBox("Quarterly review").BindToPlaceholder(PowerPointPlaceholderType.Title);
            slide.AddImage(photoPath).BindToPlaceholder(PowerPointPlaceholderType.Picture);
            PowerPointTable table = slide.AddTable(3, 2);
            table.BindToPlaceholder(PowerPointPlaceholderType.Object);
            table.GetCell(0, 0).Text = "Region";
            table.GetCell(0, 1).Text = "Revenue";
            table.GetCell(1, 0).Text = "North";
            table.GetCell(1, 1).Text = "120";
            table.GetCell(2, 0).Text = "South";
            table.GetCell(2, 1).Text = "95";

            // Glued connectors stay attached when the shapes move.
            PowerPointSlide flow = presentation.AddSlide(PowerPointSlideLayoutType.TitleOnly);
            flow.AddTextBox("Approval flow").BindToPlaceholder(PowerPointPlaceholderType.Title);
            PowerPointTextBox draft = flow.AddTextShape(OfficePresetShapeType.RoundRectangle, "Draft",
                OfficeLength.Points(60), OfficeLength.Points(170), OfficeLength.Points(140), OfficeLength.Points(60));
            PowerPointTextBox review = flow.AddTextShape(OfficePresetShapeType.FlowChartDecision, "Review",
                OfficeLength.Points(300), OfficeLength.Points(160), OfficeLength.Points(140), OfficeLength.Points(80));
            PowerPointTextBox publish = flow.AddTextShape(OfficePresetShapeType.Ellipse, "Publish",
                OfficeLength.Points(540), OfficeLength.Points(300), OfficeLength.Points(140), OfficeLength.Points(70));

            foreach (PowerPointTextBox step in new[] { draft, review, publish }) {
                step.FillColor = OfficeColor.Parse("DDEBF7");
                step.OutlineColor = brand;
            }

            flow.AddConnector(draft, review, PowerPointConnectorKind.Straight).SetLineEnds(null, OfficeLineMarkerKind.Triangle);
            PowerPointConnectionShape approved = flow.AddConnector(review, 3, publish, 0, PowerPointConnectorKind.Elbow);
            approved.Stroke(brand, OfficeLength.Points(2));
            approved.SetLineEnds(null, OfficeLineMarkerKind.Triangle);
            flow.AddConnector(publish, 2, draft, 2, PowerPointConnectorKind.Curved).OutlineDash = PowerPointLineDashStyle.Dash;

            // Moving a glued shape reroutes its connectors.
            publish.MoveBy(OfficeLength.Points(0), OfficeLength.Points(20));

            dark.Remove(layout);
            presentation.Save();
            if (openPowerPoint) ExampleFileLauncher.Open(filePath);
        }
    }
}
