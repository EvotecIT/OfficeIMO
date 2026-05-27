using System;
using System.IO;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

namespace OfficeIMO.Examples.Visio {
    public static class BlockDiagramBuilder {
        public static void Example_BlockDiagramBuilder(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Block diagram builder");
            string filePath = Path.Combine(folderPath, "Block Diagram Builder.vsdx");

            VisioDocument.Create(filePath)
                .BlockDiagram("Block Diagram", diagram => diagram
                    .Region("processor", "Processor", 1, 2, 2, 2)
                    .Block("input", "Input Device", 0, 2)
                    .EmphasisBlock("memory", "Memory Unit", 1, 2)
                    .Block("storage", "Secondary\nStorage", 1, 0, VisioBlockShapeKind.Data)
                    .Block("control", "Control Unit", 1, 3)
                    .Block("alu", "Arithmetic &\nLogic Unit", 1, 4)
                    .Block("output", "Output Device", 3, 2)
                    .DataFlow("input", "memory")
                    .DataFlow("memory", "output")
                    .DataFlow("storage", "memory")
                    .DataFlow("control", "alu")
                    .ControlFlow("control", "input")
                    .ControlFlow("control", "memory")
                    .ControlFlow("control", "output", "Control Flow"))
                .Save();

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
