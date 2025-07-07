using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    /// <summary>
    /// Example showcasing advanced XML serialization and editing.
    /// </summary>
    internal static partial class XmlSerialization {
        /// <summary>
        /// Creates a document, exports a paragraph to XML, edits it and imports it back.
        /// </summary>
        /// <param name="folderPath">Destination folder for the file.</param>
        /// <param name="openWord">Opens Word when <c>true</c>.</param>
        public static void Example_XmlSerializationAdvanced(string folderPath, bool openWord) {
            Console.WriteLine("[*] Demonstrating advanced XML manipulation");
            string filePath = Path.Combine(folderPath, "XmlSerializationAdvanced.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Original text");
                string xml = paragraph.ToXml();

                // modify XML to change the displayed text
                xml = xml.Replace("Original text", "Text updated via XML");

                document.AddParagraphFromXml(xml);
                document.Save(openWord);
            }
        }
    }
}
