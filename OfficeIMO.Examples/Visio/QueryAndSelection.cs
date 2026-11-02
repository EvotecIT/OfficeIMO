using System;
using System.IO;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Examples.Visio {
    public static class QueryAndSelection {
        public static void Example_QueryAndSelection(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Query and selection editing");
            string filePath = Path.Combine(folderPath, "Query And Selection.vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Ownership", OfficeLength.Inches(11), OfficeLength.Inches(8.5));

            VisioShape intake = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "intake", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(6, page.DefaultUnit)), "Receive request");
            VisioShape review = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "review", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(6, page.DefaultUnit)), "Review request");
            VisioShape decision = page.AddStencilShape(VisioStencils.Flowchart.Get("decision"), "approved", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(6, page.DefaultUnit)), "Approved?");
            VisioShape archive = page.AddStencilShape(VisioStencils.Flowchart.Get("data"), "archive", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3.5, page.DefaultUnit)), "Archive");

            page.AddConnector(intake, review, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left).EndArrow = EndArrow.Triangle;
            page.AddConnector(review, decision, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left).EndArrow = EndArrow.Triangle;
            page.AddConnector(decision, archive, ConnectorKind.Dynamic, VisioSide.Bottom, VisioSide.Top).EndArrow = EndArrow.Triangle;

            intake.Data["Owner"] = "Operations";
            review.Data["Owner"] = "Operations";
            decision.Data["Owner"] = "Compliance";
            archive.Data["Owner"] = "Records";

            page.SelectWithData("Owner", "Operations")
                .Fill(Color.FromRgb(224, 244, 255))
                .Stroke(Color.FromRgb(0, 120, 212), 0.02)
                .Text(shape => shape.Text + "\n(Operations)");

            page.SelectByMaster("Decision")
                .Fill(Color.FromRgb(255, 242, 204))
                .Stroke(Color.FromRgb(191, 144, 0), 0.02);

            page.SelectOutgoingConnectors(review)
                .Stroke(Color.FromRgb(0, 120, 212), 0.025)
                .EndArrow(EndArrow.Triangle)
                .Label("handoff");

            page.SelectConnectors(connector => connector.To is { } target && target.Data.TryGetValue("Owner", out string? owner) && owner == "Records")
                .LinePattern(2)
                .LineColor(Color.FromRgb(91, 155, 213));

            document.Save();

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
