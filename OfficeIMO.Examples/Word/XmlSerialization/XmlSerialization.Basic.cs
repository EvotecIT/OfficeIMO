using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class XmlSerialization {
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
