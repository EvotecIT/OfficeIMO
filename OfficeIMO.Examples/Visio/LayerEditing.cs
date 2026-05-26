using System;
using System.IO;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Examples.Visio {
    public static class LayerEditing {
        public static void Example_LayerEditing(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Layer editing");
            string filePath = Path.Combine(folderPath, "Layer Editing.vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Layered Architecture", 11, 8.5);
            page.AddLayer("Infrastructure");
            page.AddLayer("Annotations").Print = false;

            VisioShape server = page.AddStencilShape(VisioStencils.Network.Get("server"), "server", 2, 5, "Server");
            VisioShape database = page.AddStencilShape(VisioStencils.Network.Get("database"), "database", 5, 5, "Database");
            VisioShape note = page.AddStencilShape(VisioStencils.BasicShapes.Get("rectangle"), "note", 8, 5, "Internal note");
            VisioConnector connector = page.AddConnector(server, database, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);
            connector.Label = "SQL";

            page.AddToLayer("Infrastructure", server)
                .AddToLayer("Infrastructure", database)
                .AddToLayer("Infrastructure", connector)
                .AddToLayer("Annotations", note);

            page.SelectLayer("Infrastructure").Stroke(Color.DodgerBlue, 0.02);
            page.SelectLayer("Annotations").Fill(Color.LightYellow);

            document.Save();

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
