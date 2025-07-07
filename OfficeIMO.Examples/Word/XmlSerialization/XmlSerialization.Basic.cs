using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    /// <summary>
    /// Provides simple XML serialization examples.
    /// </summary>
    internal static partial class XmlSerialization {
        /// <summary>
        /// Saves a paragraph to XML and re-inserts it into the document.
        /// </summary>
        /// <param name="folderPath">Destination folder for the file.</param>
        /// <param name="openWord">Opens Word when <c>true</c>.</param>
        public static void Example_XmlSerializationBasic(string folderPath, bool openWord) {
            Console.WriteLine("[*] Demonstrating basic XML serialization");
            string filePath = Path.Combine(folderPath, "XmlSerializationBasic.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Serialized paragraph");
                string xml = paragraph.ToXml();

                document.AddParagraphFromXml(xml);
                document.Save(openWord);
            }
        }
    }
}
