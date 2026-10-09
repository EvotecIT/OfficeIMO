using System;
using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.OpenDocument;
using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.OpenDocument;

internal static class DrawDocument {
    internal static void Example(string folderPath) {
        Directory.CreateDirectory(folderPath);
        OdgDocument drawing = OdgDocument.Create();
        OdgPage page = drawing.AddPage("Workflow", OdfLength.Centimeters(24), OdfLength.Centimeters(16));
        OdgShape start = page.Shapes.AddRoundedRectangle(OdfRect.FromCentimeters(1, 3, 6, 3), OdfLength.Centimeters(0.4));
        start.Text = "Receive order";
        start.FillColor = OdfColor.Parse("#D1E9FF");
        start.FontSize = OdfLength.Points(18);
        OdgShape end = page.Shapes.AddEllipse(OdfRect.FromCentimeters(12, 3, 6, 3)); end.Text = "Check stock";
        page.Shapes.AddConnector(start.AddGluePoint(OdgGluePointAlignment.Right), end.AddGluePoint(OdgGluePointAlignment.Left));
        drawing.Layers.Add("Notes", OdgLayerDisplay.Screen);
        page.Shapes.AddTextBox(OdfRect.FromCentimeters(1, 8, 10, 2), "Internal review notes").Layer = "Notes";
        OdgShape badge = page.Shapes.AddPath(OdfRect.FromCentimeters(12, 8, 6, 3),
            new OdfViewBox(0, 0, 200, 100), "M0 50 Q0 0 50 0 H150 Q200 0 200 50 T150 100 H50 Q0 100 0 50 Z");
        badge.FillColor = OdfColor.Parse("#DCFAE6");
        drawing.Save(Path.Combine(folderPath, "workflow.odg"));
        drawing.SaveFlatXml(Path.Combine(folderPath, "workflow.fodg"));

        OdgDocument edited = OdgDocument.LoadFlatXml(Path.Combine(folderPath, "workflow.fodg"));
        edited.Pages[0].Shapes[0].Text = "Order received";
        edited.Save(Path.Combine(folderPath, "workflow-edited.odg"));

        // This profile approximates text and styling; inspect the report before using the preview.
        OdfConversionResult<OfficeDrawing> projection = edited.Pages[0].ToDrawing();
        foreach (var mapping in projection.Report.Mappings)
            Console.WriteLine($"{mapping.Feature}: {mapping.Status}");
        OfficeDrawing scene = projection.Value;
        File.WriteAllText(Path.Combine(folderPath, "workflow.svg"),
            OfficeDrawingSvgExporter.ToSvg(scene, 1, OfficeSvgSizeUnit.Point));
        PdfDocument.Create().Compose(document => document.Page(pdfPage => pdfPage
            .Size(scene.Width, scene.Height).Margin(0).Content(content => content.Drawing(scene))))
            .Save(Path.Combine(folderPath, "workflow.pdf"));
    }
}
