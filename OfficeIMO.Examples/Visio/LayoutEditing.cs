using System;
using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Examples.Visio {
    public static class LayoutEditing {
        public static void Example_LayoutEditing(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Layout editing helpers");
            string filePath = Path.Combine(folderPath, "Layout Editing.vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Cleanup", 11, 8.5);

            VisioShape intake = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "intake", 2, 6, "Receive request from customer");
            VisioShape review = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "review", 6, 5.2, "Review documentation and assign owner");
            VisioShape approve = page.AddStencilShape(VisioStencils.Flowchart.Get("decision"), "approve", 9, 4.1, "Approved?");
            VisioShape archive = page.AddStencilShape(VisioStencils.Flowchart.Get("data"), "archive", 5, 1.8, "Archive case data");

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
