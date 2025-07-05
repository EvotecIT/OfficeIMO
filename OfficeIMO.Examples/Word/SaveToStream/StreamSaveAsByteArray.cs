using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class SaveToStream {
        public static void Example_SaveAsByteArray(string folderPath, bool openWord) {
            Console.WriteLine("[*] Saving document as a byte array");
            byte[] bytes;
            using (var document = WordDocument.Create()) {
                document.AddParagraph("Saved to byte array");
                bytes = document.SaveAsByteArray();
            }

            string filePath = Path.Combine(folderPath, "SaveAsByteArray.docx");
            File.WriteAllBytes(filePath, bytes);
            Helpers.Open(filePath, openWord);
        }

        public static void Example_SaveAsMemoryStream(string folderPath, bool openWord) {
            Console.WriteLine("[*] Saving document to a MemoryStream");
            using var document = WordDocument.Create();
            document.AddParagraph("Saved to memory stream");

            using MemoryStream stream = document.SaveAsMemoryStream();

            string filePath = Path.Combine(folderPath, "SaveAsMemoryStream.docx");
            using (var file = new FileStream(filePath, FileMode.Create, FileAccess.Write)) {
                stream.CopyTo(file);
            }
            Helpers.Open(filePath, openWord);
        }

        public static void Example_SaveAsStream(string folderPath, bool openWord) {
            Console.WriteLine("[*] Cloning document into a provided stream");
            using var document = WordDocument.Create();
            document.AddParagraph("Cloned into stream");

            using var stream = new MemoryStream();
            using var cloned = document.SaveAs(stream);

            string filePath = Path.Combine(folderPath, "SaveAsStream.docx");
            using (var file = new FileStream(filePath, FileMode.Create, FileAccess.Write)) {
                stream.Position = 0;
                stream.CopyTo(file);
            }
            Helpers.Open(filePath, openWord);
        }
    }
}
