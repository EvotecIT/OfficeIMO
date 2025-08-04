using System;
using System.IO;
using OfficeIMO.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlImages(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlImage.docx");
            byte[] imageBytes = File.ReadAllBytes(Path.Combine("Assets", "OfficeIMO.png"));
            string base64 = Convert.ToBase64String(imageBytes);
            string html = $"<p><img src=\"data:image/png;base64,{base64}\" /></p>";

            using (MemoryStream ms = new MemoryStream()) {
                HtmlToWordConverter.Convert(html, ms, new HtmlToWordOptions());
                File.WriteAllBytes(filePath, ms.ToArray());

                ms.Position = 0;
                string roundTrip = WordToHtmlConverter.Convert(ms, new WordToHtmlOptions());
                Console.WriteLine(roundTrip);
            }

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
