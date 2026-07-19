using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    /// <summary>
    /// Demonstrates saving documents directly to various stream types.
    /// </summary>
    internal static partial class SaveToStream {
        /// <summary>
        /// Saves a document to a byte array and writes it to disk.
        /// </summary>
        /// <param name="folderPath">Directory to store the file.</param>
        /// <param name="openWord">Opens Word when <c>true</c>.</param>
        public static void Example_ToBytes(string folderPath, bool openWord) {
            Console.WriteLine("[*] Saving document as a byte array");
            byte[] bytes;
            using (var document = WordDocument.Create()) {
                document.AddParagraph("Saved to byte array");
                bytes = document.ToBytes();
            }

            string filePath = Path.Combine(folderPath, "ToDocx.docx");
            File.WriteAllBytes(filePath, bytes);
            if (openWord) ExampleFileLauncher.Open(filePath);
        }

        /// <summary>
        /// Saves a document to a <see cref="MemoryStream"/> and writes it to disk.
        /// </summary>
        /// <param name="folderPath">Directory to store the file.</param>
        /// <param name="openWord">Opens Word when <c>true</c>.</param>
        public static void Example_ToStream(string folderPath, bool openWord) {
            Console.WriteLine("[*] Saving document to a MemoryStream");
            using var document = WordDocument.Create();
            document.AddParagraph("Saved to memory stream");

            using MemoryStream stream = document.ToStream();

            string filePath = Path.Combine(folderPath, "ToDocxStream.docx");
            using (var file = new FileStream(filePath, FileMode.Create, FileAccess.Write)) {
                stream.CopyTo(file);
            }
            if (openWord) ExampleFileLauncher.Open(filePath);
        }

        /// <summary>
        /// Saves a document into a provided <see cref="Stream"/> instance.
        /// </summary>
        /// <param name="folderPath">Directory to store the file.</param>
        /// <param name="openWord">Opens Word when <c>true</c>.</param>
        public static void Example_SaveAsStream(string folderPath, bool openWord) {
            Console.WriteLine("[*] Saving document into a provided stream");
            using var document = WordDocument.Create();
            document.AddParagraph("Cloned into stream");

            using var stream = new MemoryStream();
            document.Save(stream);

            string filePath = Path.Combine(folderPath, "SaveAsStream.docx");
            using (var file = new FileStream(filePath, FileMode.Create, FileAccess.Write)) {
                stream.Position = 0;
                stream.CopyTo(file);
            }
            if (openWord) ExampleFileLauncher.Open(filePath);
        }
    }
}
