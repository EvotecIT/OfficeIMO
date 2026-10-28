using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;

namespace OfficeIMO.Examples.Showcase;

internal static partial class DocumentFeatureShowcase {
    private static void CreatePowerPointExamples(string output) {
        CreatePowerPointExample(output, "powerpoint-data-tables", "sales-table.pptx", CreateSlideTable);
        CreatePowerPointExample(output, "powerpoint-shapes", "workflow-shapes.pptx", CreateSlideShapes);
        CreatePowerPointExample(output, "powerpoint-text", "formatted-text.pptx", CreateSlideText);
        CreatePowerPointExample(output, "powerpoint-theme", "ThemeAndLayout.pptx",
            PowerPoint.ThemeAndLayoutPowerPoint.Example_PowerPointThemeAndLayout, previewSlide: 2);
    }

    private static void CreatePowerPointExample(string output, string id, string fileName,
        Action<string, bool> generate, int previewSlide = 0) {
        string folder = CreateExampleFolder(output, id);
        generate(folder, false);
        using PowerPointDocument presentation = PowerPointDocument.Load(Path.Combine(folder, fileName));
        var errors = presentation.ValidateDocument();
        if (errors.Count > 0) {
            throw new InvalidOperationException(id + ": " + errors[0].Description);
        }
        presentation.Slides[previewSlide].ExportImage(OfficeImageExportFormat.Png)
            .Save(Path.Combine(folder, "preview.png"), OfficeImageExportFileConflictPolicy.Replace);
        presentation.Slides[previewSlide].ExportImage(OfficeImageExportFormat.Svg)
            .Save(Path.Combine(folder, "preview.svg"), OfficeImageExportFileConflictPolicy.Replace);
        presentation.SaveAsPdf(Path.Combine(folder, "slides.pdf"));
    }

    private static PowerPointSlide AddFeatureSlide(PowerPointDocument presentation, string title, string subtitle) {
        PowerPointSlide slide = presentation.AddSlide();
        PowerPointTextBox heading = slide.AddTitle(title,global::OfficeIMO.OfficeLength.Centimeters(1.5),global::OfficeIMO.OfficeLength.Centimeters(1.2),global::OfficeIMO.OfficeLength.Centimeters(29),global::OfficeIMO.OfficeLength.Centimeters(1.6));
        heading.FontSize = 32;
        heading.Color = "17365D";
        PowerPointTextBox description = slide.AddTextBox(subtitle,global::OfficeIMO.OfficeLength.Centimeters(1.5),global::OfficeIMO.OfficeLength.Centimeters(3),global::OfficeIMO.OfficeLength.Centimeters(29),global::OfficeIMO.OfficeLength.Centimeters(1.1));
        description.FontSize = 16;
        description.Color = "526179";
        return slide;
    }

    private static void CreateSlideTable(string folder, bool open) {
        using PowerPointDocument presentation = PowerPointDocument.Create(Path.Combine(folder, "sales-table.pptx"));
        PowerPointSlide slide = AddFeatureSlide(presentation, "A table that stays editable", "Generate slide cells from application data, then style the rows and columns.");
        PowerPointTable table = slide.AddTable(5,4,global::OfficeIMO.OfficeLength.Centimeters(1.5),global::OfficeIMO.OfficeLength.Centimeters(5),global::OfficeIMO.OfficeLength.Centimeters(30),global::OfficeIMO.OfficeLength.Centimeters(9));
        string[,] values = {
            { "Product", "Q1", "Q2", "Change" },
            { "Platform", "120", "148", "+23%" },
            { "Support", "86", "99", "+15%" },
            { "Training", "42", "61", "+45%" },
            { "Services", "75", "84", "+12%" }
        };
        for (int row = 0; row < 5; row++) {
            for (int column = 0; column < 4; column++) {
                PowerPointTableCell cell = table.GetCell(row, column);
                cell.Text = values[row, column];
                cell.FontSize = 20;
                cell.Bold = row == 0;
                cell.Color = row == 0 ? "FFFFFF" : "17365D";
                cell.FillColor = row == 0 ? "17365D" : row % 2 == 0 ? "EAF1FB" : "F8FAFC";
                cell.VerticalAlignment = PowerPointTextVerticalAlignment.Center;
                cell.PaddingLeftPoints = 14;
                cell.HorizontalAlignment = column == 0 ? PowerPointTextAlignment.Left : PowerPointTextAlignment.Center;
            }
        }
        presentation.Save();
    }

    private static void CreateSlideShapes(string folder, bool open) {
        using PowerPointDocument presentation = PowerPointDocument.Create(Path.Combine(folder, "workflow-shapes.pptx"));
        PowerPointSlide slide = AddFeatureSlide(presentation, "Build the workflow from shapes", "Named rectangles, text boxes, and layers remain editable in the PPTX.");
        string[] titles = { "01 / Capture", "02 / Validate", "03 / Deliver" };
        string[] bodies = { "Collect the source\nand define the inputs.", "Check the document\nand inspect its output.", "Share an editable file\nand a review copy." };
        string[] fills = { "EAF1FB", "F1EAFC", "E7F6ED" };
        for (int index = 0; index < titles.Length; index++) {
            double left = 1.5 + index * 10.3;
            slide.AddRectangleCm(left, 5.5, 9.3, 7.5, "Stage " + (index + 1)).Fill(fills[index]).Stroke("CBD5E1", 1);
            PowerPointTextBox label = slide.AddTextBox(titles[index],global::OfficeIMO.OfficeLength.Centimeters(left + 0.5),global::OfficeIMO.OfficeLength.Centimeters(6.2),global::OfficeIMO.OfficeLength.Centimeters(8.3),global::OfficeIMO.OfficeLength.Centimeters(1.5));
            label.FontSize = 24;
            label.Color = "17365D";
            PowerPointTextBox body = slide.AddTextBox(bodies[index],global::OfficeIMO.OfficeLength.Centimeters(left + 0.5),global::OfficeIMO.OfficeLength.Centimeters(8.4),global::OfficeIMO.OfficeLength.Centimeters(8.3),global::OfficeIMO.OfficeLength.Centimeters(3.5));
            body.FontSize = 20;
            body.Color = "526179";
        }
        presentation.Save();
    }

    private static void CreateSlideText(string folder, bool open) {
        using PowerPointDocument presentation = PowerPointDocument.Create(Path.Combine(folder, "formatted-text.pptx"));
        PowerPointSlide slide = AddFeatureSlide(presentation, "Rich text without a screenshot", "Mix emphasis, color, and list paragraphs in native PowerPoint text boxes.");
        PowerPointTextBox text = slide.AddTextBox(string.Empty,global::OfficeIMO.OfficeLength.Centimeters(1.5),global::OfficeIMO.OfficeLength.Centimeters(5),global::OfficeIMO.OfficeLength.Centimeters(19),global::OfficeIMO.OfficeLength.Centimeters(10));
        text.FontSize = 22;
        PowerPointParagraph heading = text.AddParagraph("Make the decision clear.");
        PowerPointTextStyle.Subtitle.WithColor("17365D").Apply(heading);
        PowerPointParagraph line = text.AddParagraph();
        line.AddText("Use ");
        line.AddFormattedText("bold", bold: true).SetColor("2563EB");
        line.AddText(" for the outcome and ");
        line.AddFormattedText("italics", italic: true).SetColor("7C3AED");
        line.AddText(" for context.");
        text.AddBullet("Keep the main point visible.");
        text.AddBullet("Group the supporting details.");
        text.AddBullet("Leave enough room to read.");
        text.ApplyAutoSpacing(lineSpacingMultiplier: 1.25, spaceAfterPoints: 12);
        PowerPointTextBox note = slide.AddTextBox("Editable text\n\nSearch, select, and revise the content in the slide.",global::OfficeIMO.OfficeLength.Centimeters(22),global::OfficeIMO.OfficeLength.Centimeters(5),global::OfficeIMO.OfficeLength.Centimeters(9.5),global::OfficeIMO.OfficeLength.Centimeters(10));
        note.FillColor = "EAF1FB";
        note.Color = "17365D";
        note.FontSize = 22;
        note.SetTextMarginsCm(0.6, 0.6, 0.6, 0.6);
        presentation.Save();
    }
}
