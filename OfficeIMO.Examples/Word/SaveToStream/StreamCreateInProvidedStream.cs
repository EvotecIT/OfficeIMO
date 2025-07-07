using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    /// <summary>
    /// Examples for creating documents directly in provided streams.
    /// </summary>
    internal static partial class SaveToStream {
        /// <summary>
        /// Creates a document in a <see cref="MemoryStream"/> and saves it to disk.
        /// </summary>
        /// <param name="folderPath">Directory to store the file.</param>
        /// <param name="openWord">Opens Word when <c>true</c>.</param>
        public static void Example_CreateInProvidedStream(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document directly in a memory stream");
            using var stream = new MemoryStream();
            using (var document = WordDocument.Create(stream)) {
                document.AddParagraph("Stream paragraph");
                document.Save(stream);
            }

            string filePath = Path.Combine(folderPath, "CreateInStream.docx");
            using (var file = new FileStream(filePath, FileMode.Create, FileAccess.Write)) {
                stream.Position = 0;
                stream.CopyTo(file);
            }
            Helpers.Open(filePath, openWord);
        }

        /// <summary>
        /// Creates a document using a provided <see cref="FileStream"/>.
        /// </summary>
        /// <param name="folderPath">Directory to store the file.</param>
        /// <param name="openWord">Opens Word when <c>true</c>.</param>
        public static void Example_CreateInProvidedStreamAdvanced(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document using a FileStream");
            string filePath = Path.Combine(folderPath, "CreateInFileStream.docx");
            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite)) {
                using (var document = WordDocument.Create(fs)) {
                    document.AddParagraph("Created via FileStream");
                    document.Save(fs);
                }
            }
            Helpers.Open(filePath, openWord);
        }
    }
}
