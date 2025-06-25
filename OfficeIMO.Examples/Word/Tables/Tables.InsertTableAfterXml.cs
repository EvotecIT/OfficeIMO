using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Tables {
        internal static void Example_InsertTableAfterWithXml(string folderPath, bool openWord) {
            Console.WriteLine("[*] Inserting table after paragraph and using XML roundtrip");
            string filePath = Path.Combine(folderPath, "Example-InsertTableAfterXml.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var anchor = document.AddParagraph("Before table");
                var toClone = document.AddParagraph("This paragraph will be cloned using XML");

                // create table but do not insert
                var table = document.CreateTable(2, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Cell";

                // insert table after the first paragraph
                document.InsertTableAfter(anchor, table);

                // export paragraph to xml and re-import
                string xml = toClone.ToXml();
                document.AddParagraphFromXml(xml);

                document.Save(openWord);
            }
        }
    }
}
