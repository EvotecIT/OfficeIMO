using System;
using System.IO;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

namespace OfficeIMO.Examples.Visio {
    public static class StyleThemes {
        public static void Example_StyleThemes(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Reusable style themes");
            string flowchartPath = Path.Combine(folderPath, "Styled Flowchart.vsdx");
            string blockPath = Path.Combine(folderPath, "Styled Block Diagram.vsdx");
            string darkPath = Path.Combine(folderPath, "Styled Dark Diagram.vsdx");

            VisioStyleTheme minimal = VisioStyleTheme.Minimal();
            VisioDocument flowchart = VisioDocument.Create(flowchartPath)
                .Flowchart("Styled Approval Flow", flow => flow
                    .Theme(minimal)
                    .Start("start", "Request received")
                    .Step("review", "Review request")
                    .Decision("approved", "Approved?")
                    .Step("publish", "Publish decision")
                    .End("done", "Done")
                    .Branch("approved", "No", "review"));

            VisioPage flowPage = flowchart.Pages[0];
            flowPage.SelectByMaster("Decision").Style(minimal.Decision);
            flowPage.SelectConnectedConnectors(flowPage.FindShapeById("approved")!).Style(minimal.ControlConnector);
            flowPage.FitToContent(0.6, 0.45);
            flowchart.Save();

            VisioStyleTheme technical = VisioStyleTheme.Technical();
            VisioDocument.Create(blockPath)
                .BlockDiagram("Styled System Blocks", diagram => diagram
                    .Theme(technical)
                    .Region("zone", "Processing Zone", 0, 0, 3, 1)
                    .Block("input", "Input", 0, 0)
                    .EmphasisBlock("processor", "Processor", 1, 0)
                    .Block("output", "Output", 2, 0)
                    .DataFlow("input", "processor")
                    .ControlFlow("processor", "output", "control"))
                .Save();

            VisioStyleTheme dark = VisioStyleTheme.Dark();
            VisioDocument.Create(darkPath)
                .Flowchart("Dark Theme Approval", flow => flow
                    .Theme(dark)
                    .Start("start", "New request")
                    .Step("triage", "Triage")
                    .Decision("ready", "Ready?")
                    .Step("ship", "Ship")
                    .End("done", "Done")
                    .Branch("ready", "No", "triage"))
                .Save();

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(flowchartPath) { UseShellExecute = true });
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(blockPath) { UseShellExecute = true });
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(darkPath) { UseShellExecute = true });
            }
        }
    }
}
