using System;
using System.IO;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

namespace OfficeIMO.Examples.Visio {
    public static class ArchitectureDiagramBuilder {
        public static void Example_ArchitectureDiagramBuilder(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Architecture diagram builder");
            string filePath = Path.Combine(folderPath, "Architecture Diagram Builder.vsdx");

            VisioDocument.Create(filePath)
                .ArchitectureDiagram("Jenkins on Azure", diagram => diagram
                    .Theme(VisioStyleTheme.Technical())
                    .PageSize(14, 8.5)
                    .Region("vnet", "Virtual Network", 1, 0, 4, 3)
                    .Region("subnet", "Build Subnet", 1, 1, 4, 2)
                    .Actor("users", "Users", 0, 1)
                    .Gateway("public-ip", "Public IP", 1, 1)
                    .Service("jenkins", "Jenkins Server", 2, 1)
                    .Compute("agent", "Build Agent", 3, 1)
                    .Database("data", "Data", 2, 2)
                    .Storage("artifacts", "Artifacts", 4, 2)
                    .Security("vault", "Key Vault", 2, 0)
                    .DataFlow("users", "public-ip", "HTTPS")
                    .DataFlow("public-ip", "jenkins", "route")
                    .ControlFlow("jenkins", "agent", "scale")
                    .Dependency("jenkins", "data", "state")
                    .Dependency("jenkins", "vault", "secrets")
                    .DataFlow("agent", "artifacts", "publish"))
                .Save();

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
