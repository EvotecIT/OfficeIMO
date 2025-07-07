using System;
using System.IO;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Examples.Word {
    /// <summary>
    /// Examples for creating documents in a stream and setting properties.
    /// </summary>
    internal static partial class SaveToStream {
        /// <summary>
        /// Creates a document in memory, populates properties and writes it to disk.
        /// </summary>
        /// <param name="folderPath">Directory to store the file.</param>
        /// <param name="openWord">Opens Word when <c>true</c>.</param>
        public static void Example_StreamDocumentProperties(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document and saving to stream");
            string filePath = System.IO.Path.Combine(folderPath, "StreamDocumentProperties.docx");

            var document = WordDocument.Create();

            document.BuiltinDocumentProperties.Title = "Cover Page Templates";
            document.BuiltinDocumentProperties.Subject = "How to use Cover Pages with TOC";
            document.BuiltinDocumentProperties.Creator = "foo";
            document.BuiltinDocumentProperties.Description = "foo";
            document.BuiltinDocumentProperties.Title = "foo";
            document.BuiltinDocumentProperties.Creator = "foo";
            document.BuiltinDocumentProperties.Keywords = "foo";

            var stream = new MemoryStream();
            document.Save(stream);

            using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write)) {
                stream.CopyTo(fileStream);
            }

            Helpers.Open(filePath, openWord);
        }

    }
}
