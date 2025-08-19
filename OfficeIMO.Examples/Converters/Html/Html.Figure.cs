using System;
using System.IO;
using OfficeIMO.Word.Html;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlFigureWithCaption(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlFigureWithCaption.docx");
            var imagePath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images", "EvotecLogo.png");
            byte[] imageBytes = File.ReadAllBytes(imagePath);
            string base64 = Convert.ToBase64String(imageBytes);
            string html = $"<figure><img src=\"data:image/png;base64,{base64}\" alt=\"Logo\"/><figcaption>OfficeIMO Logo</figcaption></figure>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            doc.Save(filePath);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
