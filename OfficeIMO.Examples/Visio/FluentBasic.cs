using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    public static class FluentBasicVisio {
        public static void Example_FluentBasicVisio(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Fluent basic diagram");
            string filePath = Path.Combine(folderPath, "Fluent Basic Visio.vsdx");

            var doc = VisioDocument.Create(filePath);
            doc
               .ConfigureProperties(i => { i.SetTitle("Fluent Visio"); i.SetAuthor("OfficeIMO"); })
               .AddPage("Page-1", p => {
                   p.AddRectangle("S1", OfficeLength.Inches(1), OfficeLength.Inches(1), OfficeLength.Inches(2), OfficeLength.Inches(1), "Start");
                   p.AddDiamond("D1", OfficeLength.Inches(4), OfficeLength.Inches(1.5), OfficeLength.Inches(2), OfficeLength.Inches(2), "Decision");
                   p.AddEllipse("E1", OfficeLength.Inches(7), OfficeLength.Inches(1.5), OfficeLength.Inches(2), OfficeLength.Inches(1), "End");
                   p.AddConnector("S1", "D1", configure: c => {
                       c.SetKind(ConnectorKind.RightAngle);
                       c.SetArrowEnd(EndArrow.Triangle);
                   });
                   p.AddConnector("D1", "E1", configure: c => {
                       c.SetKind(ConnectorKind.RightAngle);
                       c.SetArrowEnd(EndArrow.Triangle);
                       c.SetLabel("Yes");
                   });
               }
               );
            doc.Save();

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}

