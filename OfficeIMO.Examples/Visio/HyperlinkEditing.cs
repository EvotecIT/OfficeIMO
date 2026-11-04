using System;
using System.IO;
using OfficeIMO.Visio;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Examples.Visio {
    public static class HyperlinkEditing {
        public static void Example_HyperlinkEditing(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Hyperlink editing");
            string filePath = Path.Combine(folderPath, "Hyperlink Editing.vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Linked Architecture", OfficeLength.Inches(11), OfficeLength.Inches(8.5));

            VisioShape portal = page.AddRectangle(OfficeLength.Inches(2), OfficeLength.Inches(5.5), OfficeLength.Inches(2.2), OfficeLength.Inches(1), "Portal");
            VisioShape api = page.AddRectangle(OfficeLength.Inches(5.5), OfficeLength.Inches(5.5), OfficeLength.Inches(2.2), OfficeLength.Inches(1), "API");
            VisioShape runbook = page.AddRectangle(OfficeLength.Inches(9), OfficeLength.Inches(5.5), OfficeLength.Inches(2.2), OfficeLength.Inches(1), "Runbook");

            portal.FillColor = Color.LightYellow;
            portal.AddHyperlink("https://github.com/EvotecIT/OfficeIMO", "OfficeIMO repository");

            VisioConnector portalToApi = page.AddConnector(portal, api, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);
            portalToApi.Label = "OpenAPI";
            portalToApi.PlaceLabel(0.5, 0, 0.25);
            portalToApi.AddHyperlink("https://example.org/openapi.json", "API contract");

            VisioConnector apiToRunbook = page.AddConnector(api, runbook, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);
            apiToRunbook.Label = "Escalation";
            apiToRunbook.PlaceLabel(0.5, 0, 0.25);
            apiToRunbook.AddHyperlink("https://example.org/runbook", "Runbook");

            page.SelectWithHyperlinks().Stroke(Color.DodgerBlue, 0.02);
            page.SelectConnectorsWithHyperlinks().EndArrow(EndArrow.Triangle);

            document.Save();

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
