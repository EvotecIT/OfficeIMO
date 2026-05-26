using System;
using System.IO;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

namespace OfficeIMO.Examples.Visio {
    public static class OrgChartDiagramBuilder {
        public static void Example_OrgChartDiagramBuilder(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Org chart diagram builder");
            string filePath = Path.Combine(folderPath, "Org Chart Diagram Builder.vsdx");

            VisioDocument.Create(filePath)
                .OrgChartDiagram("Leadership", org => org
                    .Theme(VisioStyleTheme.Modern())
                    .Root("ceo", "Marta Nowak", "Chief Executive Officer")
                    .Assistant("ea", "Eli Green", "Executive Assistant", "ceo")
                    .Manager("cto", "Alex Chen", "Chief Technology Officer", "ceo")
                    .Manager("coo", "Sam Rivera", "Chief Operating Officer", "ceo")
                    .Manager("cfo", "Priya Shah", "Chief Financial Officer", "ceo")
                    .TeamBand("engineering", "Engineering", "cto")
                    .TeamBand("operations", "Operations", "coo")
                    .Position("platform", "Nina Patel", "Platform Lead", "cto", "engineering")
                    .Position("security", "Owen Brooks", "Security Lead", "cto", "engineering")
                    .Vacancy("sre", "Open SRE Role", "coo", "operations")
                    .External("advisor", "Taylor Reed", "Advisor", "cfo"))
                .Save();

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
