using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class BasicDocument {
        public static void Example_BasicLoadHamlet(string templatesPath, string folderPath, bool openWord) {
            Console.WriteLine("[*] Loading Hamlet Document");
            string filePath = System.IO.Path.Combine(templatesPath, "Hamlet.docx");

            using (WordDocument document = WordDocument.Load(filePath)) {
                Console.WriteLine($"Footnotes position: {document.FootnoteSettings.Position}");
                Console.WriteLine($"Endnotes position: {document.EndnoteSettings.Position}");
                Console.WriteLine($"Footnotes start: {document.FootnoteSettings.StartNumber}");
                Console.WriteLine($"Endnotes restart: {document.EndnoteSettings.NumberingRestart}");

                document.AddFootnoteProperties(position: WordFootnotePosition.PageBottom,
                                            restartNumbering: WordNoteNumberRestart.EachSection,
                                            startNumber: 1);
                document.AddEndnoteProperties(position: WordEndnotePosition.SectionEnd,
                                            restartNumbering: WordNoteNumberRestart.EachSection,
                                            startNumber: 1);

                Console.WriteLine("----");
                Console.WriteLine(document.Sections.Count);
                Console.WriteLine("----");
                Console.WriteLine(document.Sections[0].Paragraphs.Count);
                Console.WriteLine(document.Sections[0].Paragraphs.Count);
                Console.WriteLine(document.Sections[0].Paragraphs.Count);

                Console.WriteLine(document.Sections[0].HyperLinks.Count);
                Console.WriteLine(document.HyperLinks.Count);
                Console.WriteLine(document.Fields.Count);
                document.Save();
                if (openWord) document.OpenInApplication();
            }
        }
    }
}
