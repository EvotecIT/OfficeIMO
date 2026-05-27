using System;
using System.IO;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

namespace OfficeIMO.Examples.Visio {
    public static class SequenceDiagramBuilder {
        public static void Example_SequenceDiagramBuilder(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Sequence diagram builder");
            string filePath = Path.Combine(folderPath, "Sequence Diagram Builder.vsdx");

            VisioDocument.Create(filePath)
                .SequenceDiagram("Checkout Sequence", sequence => sequence
                    .Title()
                    .Theme(VisioStyleTheme.Fluent())
                    .PageSize(8, 5)
                    .Actor("customer", "Customer")
                    .Participant("web", "Web App")
                    .Control("api", "Orders API")
                    .Database("db", "Orders DB")
                    .Call("customer", "web", "Checkout", "checkout")
                    .Call("web", "api", "POST /orders", "post-order")
                    .Async("api", "db", "Persist order", "persist")
                    .Return("api", "web", "201 Created", "created")
                    .SelfMessage("web", "Render receipt", id: "render"))
                .Save();

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
