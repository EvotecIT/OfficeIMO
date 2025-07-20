using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    /// <summary>
    /// Demonstrates saving back to the original stream when no file path is provided.
    /// </summary>
    internal static partial class SaveToStream {
        public static void Example_SaveToOriginalStream(string folderPath, bool openWord) {
            Console.WriteLine("[*] Saving document back to the original stream");
            using var stream = new MemoryStream();
            using (var document = WordDocument.Create(stream)) {
                document.AddParagraph("Stream paragraph");
                document.Save();
            }

            string filePath = Path.Combine(folderPath, "SaveToOriginalStream.docx");
            using (var file = new FileStream(filePath, FileMode.Create, FileAccess.Write)) {
                stream.Position = 0;
                stream.CopyTo(file);
            }
            Helpers.Open(filePath, openWord);
        }
    }
}
