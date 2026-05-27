using System;
using System.IO;
using OfficeIMO.Visio;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Examples.Visio {
    public static class HyperlinkEditing {
        public static void Example_HyperlinkEditing(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Hyperlink editing");
            string filePath = Path.Combine(folderPath, "Hyperlink Editing.vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Linked Architecture", 11, 8.5);

            VisioShape portal = page.AddRectangle(2, 5.5, 2.2, 1, "Portal");
            VisioShape api = page.AddRectangle(5.5, 5.5, 2.2, 1, "API");
            VisioShape runbook = page.AddRectangle(9, 5.5, 2.2, 1, "Runbook");

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
