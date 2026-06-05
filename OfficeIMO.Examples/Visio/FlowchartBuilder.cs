using System;
using System.IO;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

namespace OfficeIMO.Examples.Visio {
    public static class FlowchartBuilder {
        public static void Example_FlowchartBuilder(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Flowchart builder");
            string filePath = Path.Combine(folderPath, "Flowchart Builder.vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .Flowchart("Property buying Flowchart", flow => flow
                    .PageSize(9.5, 15.5)
                    .Layout(VisioFlowchartLayout.TwoColumnContinuation)
                    .RouteBranches(laneSpacing: 0.5)
                    .Start("start", "Start with an agent\nyou trust")
                    .Step("consult", "Consult with agent to\ndetermine your property\nwants and needs")
                    .Step("paperwork", "Review and complete\npaperwork")
                    .Step("loan", "Go to preferred lender,\nget pre-qualified and\npre-approval for loan\namount")
                    .Step("market", "With agent, analyze\nmarket to choose\nproperties of interest")
                    .Step("view", "View properties\nwith agent")
                    .OffPage("jump", "A")
                    .Continue("resume", "A")
                    .Step("offer", "Select ideal property\nand write offer to\npurchase")
                    .Decision("agreement", "Negotiate\n& Counteroffer:\nAgreement?")
                    .Step("contract", "Accept the contract")
                    .Step("underwriting", "Secure underwriting,\nobtain loan approval")
                    .Step("closing", "Select/Contact closing\nattorney for title exam\nand title insurance")
                    .Step("inspection", "Schedule inspection\nand survey")
                    .End("close", "Close on the\nproperty")
                    .Branch("agreement", "No", "market"));
            document.Pages[0].FitToContent(0.7, 0.55);
            document.Save();

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
