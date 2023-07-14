using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class FootNotes {
        internal static void Example_DocumentWithFootNotes(string folderPath, bool openWord) {
            Console.WriteLine("[*] Opening Document with foot notes");
            var filePath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Templates", "DocumentWithFootNotes.docx");

            using (WordDocument document = WordDocument.Load(filePath)) {
                //foreach (var field in document.Fields) {
                //    Console.WriteLine("...Type: " + field.FieldType);
                //    Console.WriteLine("...Format switch: " + field.FieldFormat);
                //    Console.WriteLine("...Instruction: " + String.Join(" ", field.FieldInstructions));
                //    Console.WriteLine("...Switches: " + String.Join(" ", field.FieldSwitches));
                //}

                var fileTarget = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Documents", "Document with FootNotes01.docx");
                document.Save(fileTarget, openWord);
            }
        }

        internal static void Example_DocumentWithFootNotesEmpty(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with footnotes/end notes");
            string filePath = System.IO.Path.Combine(folderPath, "Document with FootNotes02.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Basic paragraph");
                paragraph.ParagraphAlignment = JustificationValues.Center;

                document.AddParagraph("This is my text").AddFootNote("This is a footnote to my text")
                    .AddText(" continuing").AddFootNote("2nd footnote!");

                Console.WriteLine("EndNotes count " + document.EndNotes.Count);
                Console.WriteLine("EndNotes Section count " + document.Sections[0].EndNotes.Count);

                Console.WriteLine("FootNotes count " + document.FootNotes.Count);
                Console.WriteLine("FootNotes Section count " + document.Sections[0].FootNotes.Count);


                var lastFootNoteParagraph = document.AddParagraph("Another paragraph").AddFootNote("more footnotes!")
                    .AddText(" more within paragraph").AddFootNote("4th footnote!");

                Console.WriteLine("Is paragraph foot note: " + lastFootNoteParagraph.IsFootNote);

                var footNoteParagraphs = lastFootNoteParagraph.FootNote.Paragraphs;

                Console.WriteLine("Paragraphs within footnote: " + footNoteParagraphs.Count);
                Console.WriteLine("What's the text: " + footNoteParagraphs[1].Text);
                footNoteParagraphs[1].Bold = true;

                document.AddParagraph("Testing endnote - 1").AddEndNote("Test end note 1");

                document.AddParagraph("Test 1");

                document.AddSection();

                document.AddParagraph("Testing endnote - 2").AddEndNote("Test end note 2");

                Console.WriteLine("EndNotes count " + document.EndNotes.Count);
                Console.WriteLine("EndNotes Section count " + document.Sections[0].EndNotes.Count);

                Console.WriteLine("FootNotes count " + document.FootNotes.Count);
                Console.WriteLine("FootNotes Section count " + document.Sections[0].FootNotes.Count);

                document.Save(openWord);
            }
        }
    }
}