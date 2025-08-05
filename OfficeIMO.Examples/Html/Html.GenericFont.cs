using System;
using System.IO;
using OfficeIMO.Converters;
using OfficeIMO.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlGenericFont(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlGenericFont.docx");
            string html = "<p>Generic font sample.</p>";

            ConverterRegistry.Register("html->word", () => new HtmlToWordConverter());
            using MemoryStream input = new MemoryStream(Encoding.UTF8.GetBytes(html));
            using MemoryStream output = new MemoryStream();
            IWordConverter converter = ConverterRegistry.Resolve("html->word");
            converter.Convert(input, output, new HtmlToWordOptions { FontFamily = "monospace" });
            File.WriteAllBytes(filePath, output.ToArray());

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}

