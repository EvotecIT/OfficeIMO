using System;
using System.IO;
using OfficeIMO.Visio;
using OfficeIMO.Drawing;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Examples.Visio {
    /// <summary>
    /// Demonstrates basic <see cref="VisioDocument"/> usage with colors.
    /// </summary>
    public static class BasicVisioDocument {
        public static void Example_BasicVisio(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Creating basic document");
            string filePath = Path.Combine(folderPath, "Basic Visio.vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            document
                .ConfigureProperties(info => { info.SetTitle("Basic Visio"); info.SetAuthor("OfficeIMO"); })
                .AddPage("Page-1", OfficeLength.Inches(8.5), OfficeLength.Inches(11), VisioMeasurementUnit.Inches, configure: p => {
                    p.AddRectangle("start", OfficeLength.Inches(2), OfficeLength.Inches(2), OfficeLength.Inches(2), OfficeLength.Inches(1), "Rectangle");
                    p.GetShape("start", s => {
                        s.SetFill(Color.LightBlue);
                        s.SetStroke(Color.DarkBlue, 0.02);
                        s.SetFont("Aptos");
                        s.SetFontSize(14);
                        s.SetTextColor(Color.DarkBlue);
                        s.SetBold();
                        s.SetItalic();
                        s.SetUnderline(OfficeTextDecorationStyle.Double);
                        s.SetStrikethrough(OfficeTextDecorationStyle.Single);
                        s.SetBaseline(OfficeTextBaseline.Superscript);
                        s.SetCapitalization(VisioTextCapitalization.AllCaps);
                        s.TransformTextCase(OfficeTextCase.TitleCase);
                    });
                });
            document.Save();
            File.WriteAllText(Path.Combine(folderPath, "Basic Visio.svg"), document.ToSvg());
            document.SaveAsImage(Path.Combine(folderPath, "Basic Visio.png"), new VisioImageExportOptions { TargetDpi = 144 });

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}

