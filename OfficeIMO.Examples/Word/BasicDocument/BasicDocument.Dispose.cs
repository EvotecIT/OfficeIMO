using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class BasicDocument {
        /// <summary>
        /// Creates a document and disposes it multiple times.
        /// </summary>
        public static void Example_DisposeMultipleTimes(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document and disposing it multiple times");
            string filePath = System.IO.Path.Combine(folderPath, "DisposeMultipleTimes.docx");
            WordDocument document = WordDocument.Create(filePath);
            document.AddParagraph("This is my test");
            document.Save();
            document.Dispose();
            document.Dispose();
            Helpers.Open(filePath, openWord);
        }
    }
}
