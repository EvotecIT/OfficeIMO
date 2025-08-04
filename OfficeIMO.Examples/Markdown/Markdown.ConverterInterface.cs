using System;
using System.IO;
using System.Text;
using OfficeIMO.Converters;
using OfficeIMO.Markdown;

namespace OfficeIMO.Examples.Markdown {
    internal static partial class Markdown {
        public static void Example_MarkdownInterface(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "MarkdownInterface.docx");
            string markdown = "# Title\nContent";
            using MemoryStream input = new MemoryStream(Encoding.UTF8.GetBytes(markdown));
            using MemoryStream output = new MemoryStream();
            IWordConverter converter = new MarkdownToWordConverter();
            converter.Convert(input, output, new MarkdownToWordOptions());
            File.WriteAllBytes(filePath, output.ToArray());
            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
