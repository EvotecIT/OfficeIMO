using System;
using System.IO;
using OfficeIMO.Markdown;
using OfficeIMO.Word;
using System.Text;

namespace OfficeIMO.Examples.Markdown {
    internal static partial class Markdown {
        public static void Example_MarkdownImages(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "MarkdownImage.docx");
            string assetPath = Path.Combine("Assets", "OfficeIMO.png");
            byte[] bytes = File.ReadAllBytes(assetPath);
            string base64 = Convert.ToBase64String(bytes);
            string markdown = $"![OfficeIMO logo](data:image/png;base64,{base64})";

            ConverterRegistry.Register("markdown->word", () => new MarkdownToWordConverter());
            using MemoryStream mdStream = new MemoryStream(Encoding.UTF8.GetBytes(markdown));
            using MemoryStream wordStream = new MemoryStream();
            IWordConverter converter = ConverterRegistry.Resolve("markdown->word");
            converter.Convert(mdStream, wordStream, new MarkdownToWordOptions());
            File.WriteAllBytes(filePath, wordStream.ToArray());

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
