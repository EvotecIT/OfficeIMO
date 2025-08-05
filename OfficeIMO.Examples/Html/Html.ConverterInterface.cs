using System;
using System.IO;
using System.Text;
using OfficeIMO.Converters;
using OfficeIMO.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlInterface(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlInterface.docx");
            string html = "<p>Hello world</p>";
            using MemoryStream input = new MemoryStream(Encoding.UTF8.GetBytes(html));
            using MemoryStream output = new MemoryStream();
            ConverterRegistry.Register("html->word", () => new HtmlToWordConverter());
            IWordConverter converter = ConverterRegistry.Resolve("html->word");
            converter.Convert(input, output, new HtmlToWordOptions());
            File.WriteAllBytes(filePath, output.ToArray());
            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
