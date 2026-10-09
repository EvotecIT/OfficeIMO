using global::ChartForgeX.Themes;
using global::ChartForgeX.VisualArtifacts;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.ChartForgeX.Examples;

internal static class DocumentDelivery {
    public static void Write(string output, IReadOnlyList<VisualSpecimen> specimens) {
        foreach (var specimen in specimens) {
            File.WriteAllText(Path.Combine(output, "source-" + specimen.Name + ".svg"), specimen.Artifact.ToSvg());
            File.WriteAllBytes(Path.Combine(output, "source-" + specimen.Name + ".png"), specimen.Artifact.ToPng());
        }
        WriteWord(Path.Combine(output, "service-review.docx"), specimens);
        WriteExcel(Path.Combine(output, "service-review.xlsx"), specimens);
        WritePowerPoint(Path.Combine(output, "service-review.pptx"), specimens);
        WritePdf(Path.Combine(output, "service-review.pdf"), specimens);
        WriteVisio(output, specimens);
        DeliveryEvidence.SaveJson(Path.Combine(output, "visual-conversions.json"), specimens.Select(item => {
            var conversion = item.Convert(450D);
            return new { item.Name, item.ModeName, conversion.Id, conversion.WidthPoints, conversion.HeightPoints,
                conversion.AlternativeText, conversion.PlacementMediaType, conversion.Report };
        }));
    }

    private static void WriteWord(string path, IReadOnlyList<VisualSpecimen> specimens) {
        using var document = WordDocument.Create(path);
        document.Sections[0].PageSettings.PageSize = WordPageSize.A4;
        document.Margins.Type = WordMargin.Normal;
        double usableWidth = ((document.Sections[0].PageSettings.Width ?? throw new InvalidOperationException("Missing page width."))
            - document.Margins.Left - document.Margins.Right) / 20D;
        for (int index = 0; index < specimens.Count; index++) {
            if (index > 0) document.AddPageBreak();
            var specimen = specimens[index];
            var heading = document.AddParagraph().AddText(Heading(specimen));
            heading.FontFamily = "Arial"; heading.FontSizePoints = 18D;
            document.AddParagraph().AddVisualArtifact(specimen.Convert(usableWidth));
            var caption = document.AddParagraph().AddText(specimen.Caption);
            caption.FontFamily = "Arial"; caption.FontSizePoints = 11D;
        }
        document.Save();
    }

    private static void WriteExcel(string path, IReadOnlyList<VisualSpecimen> specimens) {
        using var document = ExcelDocument.Create(path);
        foreach (var specimen in specimens) {
            var sheet = document.AddWorksheet(specimen.Name);
            for (int column = 1; column <= 9; column++) sheet.SetColumnWidth(column, 12D);
            for (int row = 1; row <= 26; row++) sheet.SetRowHeight(row, 18D);
            sheet.MergeRange("B1:I2");
            sheet.CellValue(1, 2, Heading(specimen));
            sheet.CellFontName(1, 2, "Arial"); sheet.CellFontSize(1, 2, 18D);
            // The picture remains within the authored Letter print width after the half-inch margins.
            sheet.SetPageSetup(fitToWidth: 1, fitToHeight: 0, paperSize: ExcelPaperSize.Letter);
            sheet.SetMarginsPreset(ExcelMarginPreset.Narrow);
            sheet.AddVisualArtifact(4, 2, specimen.Convert(612D - 72D));
            sheet.MergeRange("B21:I23");
            sheet.CellValue(21, 2, specimen.Caption);
            sheet.CellFontName(21, 2, "Arial"); sheet.CellFontSize(21, 2, 11D);
            sheet.CellWrapText(21, 2);
        }
        document.Save();
    }

    private static void WritePowerPoint(string path, IReadOnlyList<VisualSpecimen> specimens) {
        using var presentation = PowerPointPresentation.Create(path);
        presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);
        double slideWidth = presentation.SlideSize.WidthPoints, slideHeight = presentation.SlideSize.HeightPoints;
        foreach (var specimen in specimens) {
            var slide = presentation.AddSlide();
            var colors = VisualTheme.Graphite().Resolve(specimen.Mode);
            slide.BackgroundColor = colors.Background.ToHex().TrimStart('#');
            Text(slide, Heading(specimen), 36D, 16D, slideWidth - 72D, 30D, 18, colors.Foreground.ToHex());
            var conversion = specimen.Convert(slideWidth - 72D);
            slide.AddVisualArtifact(conversion, (slideWidth - conversion.WidthPoints) / 2D, 60D);
            Text(slide, specimen.Caption, 36D, Math.Min(slideHeight - 52D, 60D + conversion.HeightPoints + 18D),
                slideWidth - 72D, 36D, 11, colors.Foreground.ToHex());
        }
        presentation.Save();
    }

    private static void WritePdf(string path, IReadOnlyList<VisualSpecimen> specimens) {
        var options = new PdfOptions { PageSize = PageSizes.A4, DefaultFontSize = 11D };
        options.EnableTaggedPdfCatalogMarkers();
        double usableWidth = options.PageWidth - options.MarginLeft - options.MarginRight;
        PdfDocument.Create(document => document.Content(content => {
            for (int index = 0; index < specimens.Count; index++) {
                if (index > 0) content.PageBreak();
                var specimen = specimens[index];
                content.H1(Heading(specimen));
                content.AddVisualArtifact(specimen.Convert(usableWidth), spacingBefore: 12D, spacingAfter: 18D);
                content.Text(specimen.Caption);
            }
        }), options).Save(path);
    }

    private static void WriteVisio(string output, IReadOnlyList<VisualSpecimen> specimens) {
        var diagrams = specimens.Where(item => item.EditableTopology).ToArray();
        var book = diagrams.Select(item => item.Artifact.ToInterchangeEnvelope()).ToOfficeVisioBook(
            new OfficeVisioVisualOptions { LayoutMode = OfficeVisioVisualLayoutMode.Preserve, PixelsPerInch = 96D });
        book.Document.Save(Path.Combine(output, "service-delivery.vsdx"));
        DeliveryEvidence.SaveJson(Path.Combine(output, "native-visio-fidelity.json"), book.Pages.Select((page, index) => new {
            diagrams[index].Name, page.Report, page.Page.Width, page.Page.Height,
            Shapes = page.Page.Shapes.Select(shape => new { shape.Id, shape.Text, shape.PinX, shape.PinY, shape.Width, shape.Height }),
            Connectors = page.Page.Connectors.Select(edge => new { edge.Id, edge.Label, edge.Waypoints.Count })
        }));
    }

    private static void Text(PowerPointSlide slide, string text, double left, double top, double width, double height,
        int size, string color) {
        var box = slide.AddTextBox(text, PowerPointUnits.FromPoints(left), PowerPointUnits.FromPoints(top),
            PowerPointUnits.FromPoints(width), PowerPointUnits.FromPoints(height));
        box.FontName = "Arial"; box.FontSize = size; box.Color = color.TrimStart('#');
    }

    private static string Heading(VisualSpecimen specimen) => specimen.Artifact.Title + " · " + specimen.ModeName;
}
