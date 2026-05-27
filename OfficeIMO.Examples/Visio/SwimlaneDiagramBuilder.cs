using System;
using System.IO;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

namespace OfficeIMO.Examples.Visio {
    public static class SwimlaneDiagramBuilder {
        public static void Example_SwimlaneDiagramBuilder(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Swimlane diagram builder");
            string filePath = Path.Combine(folderPath, "Swimlane Diagram Builder.vsdx");

            VisioDocument.Create(filePath)
                .SwimlaneDiagram("Order Fulfillment", swim => swim
                    .Theme(VisioStyleTheme.Modern())
                    .Lane("customer", "Customer")
                    .Lane("sales", "Sales")
                    .Lane("ops", "Operations")
                    .Phase("request", "Request")
                    .Phase("review", "Review")
                    .Phase("approval", "Approval")
                    .Phase("fulfill", "Fulfill")
                    .Start("start", "Submit order", "customer", "request")
                    .Step("qualify", "Qualify order", "sales", "review")
                    .Decision("approved", "Approved?", "sales", "approval")
                    .Step("revise", "Revise request", "customer", "approval")
                    .Step("pick", "Pick items", "ops", "approval")
                    .Data("invoice", "Create invoice", "sales", "fulfill")
                    .End("ship", "Ship order", "ops", "fulfill")
                    .Flow("start", "qualify", "handoff")
                    .Flow("qualify", "approved")
                    .Exception("approved", "revise", "no")
                    .Handoff("approved", "pick", "yes")
                    .Flow("pick", "invoice")
                    .Flow("invoice", "ship"))
                .Save();

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
