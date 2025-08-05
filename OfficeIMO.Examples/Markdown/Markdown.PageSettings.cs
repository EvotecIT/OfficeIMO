using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Markdown;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Markdown {
    internal static partial class Markdown {
        public static void Example_MarkdownPageSettings(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "MarkdownPageSettings.docx");
            string markdown = "Hello World";

            using MemoryStream wordStream = new MemoryStream();
            MarkdownToWordConverter.Convert(markdown, wordStream, new MarkdownToWordOptions {
                DefaultOrientation = PageOrientationValues.Landscape,
                DefaultPageSize = WordPageSize.A5
            });

            File.WriteAllBytes(filePath, wordStream.ToArray());
            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
