using System;
using System.IO;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;

namespace OfficeIMO.Examples.Visio {
    public static class ConnectorRouting {
        public static void Example_ConnectorRouting(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Connector routing");
            string filePath = Path.Combine(folderPath, "Visio Connector Routing.vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioStyleTheme theme = VisioStyleTheme.Technical();
            VisioPage page = document.AddPage("Routing", OfficeLength.Inches(11), OfficeLength.Inches(8.5));
            page.PlacementStyle = VisioPlacementStyle.HierarchyLeftToRightMiddle;
            page.PlacementDepth = VisioPlacementDepth.Medium;
            page.PlacementFlip = VisioPlacementFlip.Horizontal | VisioPlacementFlip.Rotate90;
            page.MoveShapesAwayOnDrop = true;
            page.ResizePageToFitLayout = true;
            page.EnableLayoutGrid = true;
            page.SetLayoutGridSizing(1.2, 0.45);
            page.ConnectorRouteStyle = VisioPageRouteStyle.FlowchartTopToBottom;
            page.ConnectorRouteAppearance = VisioLineRouteExtension.Straight;
            page.LineJumpStyle = VisioLineJumpStyle.Gap;
            page.LineJumpCode = VisioLineJumpCode.DisplayOrder;
            page.HorizontalLineJumpDirection = VisioHorizontalLineJumpDirection.Up;
            page.VerticalLineJumpDirection = VisioVerticalLineJumpDirection.Right;
            page.SetConnectorSpacing(0.25, 0.45);

            VisioShape intake = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "intake", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(6.2, page.DefaultUnit)), "Intake");
            VisioShape review = page.AddStencilShape(VisioStencils.Flowchart.Get("decision"), "review", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5.5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(6.2, page.DefaultUnit)), "Needs review?");
            VisioShape approve = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "approve", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(9, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(6.2, page.DefaultUnit)), "Approve");
            VisioShape rework = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "rework", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5.5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3.1, page.DefaultUnit)), "Rework");
            VisioShape archive = page.AddStencilShape(VisioStencils.Flowchart.Get("data"), "archive", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(9, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3.1, page.DefaultUnit)), "Archive");
            review.PlacementStyle = VisioPlacementStyle.HierarchyLeftToRightMiddle;
            review.PlacementFlip = VisioPlacementFlip.Horizontal | VisioPlacementFlip.Rotate90;
            review.PlowCode = VisioShapePlowCode.Always;
            archive.AllowHorizontalConnectorRoutingThrough = false;
            archive.AllowVerticalConnectorRoutingThrough = false;

            page.SelectShapes(shape => shape.MasterNameU != null).Style(theme.Primary);
            review.ApplyStyle(theme.Decision);
            archive.ApplyStyle(theme.Success);

            page.AddConnector(intake, review, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left)
                .RouteOrthogonal()
                .ApplyStyle(theme.DataConnector);

            VisioConnector yes = page.AddConnector(review, approve, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left)
                .RouteOrthogonal()
                .ApplyStyle(theme.Connector)
                .PlaceLabel(0.58, offsetY: 0.2);
            yes.Label = "yes";
            yes.RouteStyle = VisioPageRouteStyle.FlowchartLeftToRight;
            yes.RouteAppearance = VisioLineRouteExtension.Curved;
            yes.LineJumpStyle = VisioLineJumpStyle.Square;
            yes.LineJumpCode = VisioConnectorLineJumpCode.Always;
            yes.HorizontalJumpDirection = VisioHorizontalLineJumpDirection.Up;
            yes.VerticalJumpDirection = VisioVerticalLineJumpDirection.Right;
            yes.RerouteBehavior = VisioConnectorRerouteBehavior.OnCrossover;

            VisioConnector no = page.AddConnector(review, rework, ConnectorKind.Dynamic, VisioSide.Bottom, VisioSide.Top)
                .RouteOrthogonal(VisioConnectorRouteStyle.VerticalThenHorizontal, -0.25)
                .ApplyStyle(theme.ControlConnector)
                .PlaceLabel(0.55, offsetX: -0.35);
            no.Label = "no";

            page.AddConnector(rework, archive, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left)
                .RouteThrough(VisioConnectorWaypoint.At(7.2, 3.1), VisioConnectorWaypoint.At(7.2, 3.9))
                .ApplyStyle(theme.ControlConnector);

            page.FitToContent(0.7, 0.55);
            document.Save();
            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
