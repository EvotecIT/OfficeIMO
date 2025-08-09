using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class LoadDocuments {
        /// <summary>
        /// Demonstrates loading a document in read-only mode with style overrides enabled.
        /// </summary>
        /// <param name="openWord">Whether to open the document after loading.</param>
        public static void LoadWordDocument_Sample4(bool openWord) {
            Console.WriteLine("[*] Load external Word Document - Sample 4");

            string documentPaths = Path.Combine(Directory.GetCurrentDirectory(), "Templates");
            string fullPath = Path.Combine(documentPaths, "sample1.docx");

            using (WordDocument document = WordDocument.Load(fullPath, readOnly: true, overrideStyles: true)) {
                Console.WriteLine("Document loaded in read-only mode. Style overrides were ignored.");
            }
        }
    }
}

