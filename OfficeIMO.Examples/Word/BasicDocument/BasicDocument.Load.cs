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
                Console.WriteLine($"Footnotes position: {document.FootnoteProperties?.FootnotePosition?.Val}");
                Console.WriteLine($"Endnotes position: {document.EndnoteProperties?.EndnotePosition?.Val}");
                Console.WriteLine($"Footnotes start: {document.FootnoteProperties?.NumberingStart?.Val}");
                Console.WriteLine($"Endnotes restart: {document.EndnoteProperties?.NumberingRestart?.Val}");

                document.AddFootnoteProperties(position: FootnotePositionValues.PageBottom,
                                            restartNumbering: RestartNumberValues.EachSection,
                                            startNumber: 1);
                document.AddEndnoteProperties(position: EndnotePositionValues.SectionEnd,
                                            restartNumbering: RestartNumberValues.EachSection,
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
                document.Save(openWord);
            }
        }
    }
}
