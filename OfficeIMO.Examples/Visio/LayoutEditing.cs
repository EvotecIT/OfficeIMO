using System;
using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Examples.Visio {
    public static class LayoutEditing {
        public static void Example_LayoutEditing(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Layout editing helpers");
            string filePath = Path.Combine(folderPath, "Layout Editing.vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Cleanup", OfficeLength.Inches(11), OfficeLength.Inches(8.5));

            VisioShape intake = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "intake", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(6, page.DefaultUnit)), "Receive request from customer");
            VisioShape review = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "review", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(6, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5.2, page.DefaultUnit)), "Review documentation and assign owner");
            VisioShape approve = page.AddStencilShape(VisioStencils.Flowchart.Get("decision"), "approve", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(9, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4.1, page.DefaultUnit)), "Approved?");
            VisioShape archive = page.AddStencilShape(VisioStencils.Flowchart.Get("data"), "archive", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1.8, page.DefaultUnit)), "Archive case data");

            foreach (VisioShape shape in new[] { intake, review, approve, archive }) {
                shape.Data["CleanUp"] = "Yes";
            }

            VisioConnector intakeReview = page.AddConnector(intake, review, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left)
                .PlaceLabelAt(6, 5.2, width: 0.2, height: 0.1);
            intakeReview.EndArrow = EndArrow.Triangle;
            intakeReview.Label = "review";
            VisioConnector reviewApprove = page.AddConnector(review, approve, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left)
                .PlaceLabel(0.55, offsetY: 0.15, width: 0.2, height: 0.1);
            reviewApprove.EndArrow = EndArrow.Triangle;
            reviewApprove.Label = "decision";
            VisioConnector approveArchive = page.AddConnector(approve, archive, ConnectorKind.Dynamic, VisioSide.Bottom, VisioSide.Top)
                .PlaceLabel(0.55, offsetX: 0.1, width: 0.2, height: 0.1);
            approveArchive.EndArrow = EndArrow.Triangle;
            approveArchive.Label = "store";

            page.SelectWithData("CleanUp", "Yes")
                .ResizeToText(new OfficeFontInfo("Calibri", 11), horizontalPadding: 0.3, verticalPadding: 0.16)
                .Stroke(Color.FromRgb(0, 120, 212), 0.02);

            page.SelectShapes(shape => shape.Id == "intake" || shape.Id == "review" || shape.Id == "approve")
                .Align(VisioVerticalAlignment.Middle)
                .DistributeHorizontally();

            page.SelectByMaster("Decision")
                .Fill(Color.FromRgb(255, 242, 204))
                .Stroke(Color.FromRgb(191, 144, 0), 0.02);

            page.PolishDiagram(new VisioDiagramPolishOptions {
                MaximumConnectorLabelWidth = 1.4,
                FitHorizontalMargin = 0.6,
                FitVerticalMargin = 0.45
            });
            document.Save();

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
