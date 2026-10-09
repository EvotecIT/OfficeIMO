using System;
using System.IO;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio;

public static class LegacyXml {
    public static void Example_LegacyXml(string folderPath) {
        Directory.CreateDirectory(folderPath);
        VisioDocument document = VisioDocument.Create();
        VisioPage page = document.AddPage("Workflow", 10, 7);
        var start = new VisioShape("1", 2, 2, 2, 1, "Start");
        var end = new VisioShape("2", 6, 2, 2, 1, "Finish");
        page.Shapes.Add(start);
        page.Shapes.Add(end);
        page.Connectors.Add(new VisioConnector(start, end));

        var output = document.ToLegacyXmlResult();
        foreach (var diagnostic in output.Report.FidelityDiagnostics)
            Console.WriteLine(diagnostic.Message);
        // This example accepts the reported modern theme/resize omissions for its simple drawing.
        string path = Path.Combine(folderPath, "workflow.vdx");
        File.WriteAllBytes(path, output.Value);

        var imported = VisioDocument.LoadLegacyXml(path);
        foreach (var diagnostic in imported.Report.FidelityDiagnostics)
            Console.WriteLine(diagnostic.Message);
        imported.Value.Pages[0].Shapes[0].Text = "Reviewed";
        imported.Value.Save(Path.Combine(folderPath, "workflow.vsdx"));
        imported.Value.SaveLegacyXml(Path.Combine(folderPath, "workflow-edited.vdx"), allowOmissions: true);
    }
}
