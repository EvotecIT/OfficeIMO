using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class LoadDocuments {
        public static void LoadWordDocument_Sample1(bool openWord) {
            Console.WriteLine("[*] Load external Word Document - Sample 1");

            string documentPaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Templates");
            string fullPath = System.IO.Path.Combine(documentPaths, "sample1.docx");
            using (WordDocument document = WordDocument.Load(System.IO.Path.Combine(documentPaths, "sample1.docx"), false)) {
                Console.WriteLine(fullPath);
                Console.WriteLine("Sections count: " + document.Sections.Count);
                Console.WriteLine("Tables count: " + document.Tables.Count);
                Console.WriteLine("Paragraphs count: " + document.Paragraphs.Count);

                foreach (var paragraph in document.Paragraphs) {
                    if (paragraph.Text.StartsWith("You can do")) {
                        paragraph.Text = "Maybe you can't!";
                    }
                }

                // changing books from 1 to 5
                document.Tables[0].Rows[1].Cells[1].Paragraphs[0].Text = "5";

                document.Save(openWord);
            }
        }
    }
}
