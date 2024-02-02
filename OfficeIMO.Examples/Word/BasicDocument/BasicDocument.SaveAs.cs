using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class BasicDocument {
        public static void Example_BasicDocument(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with Save");
            string filePath = System.IO.Path.Combine(folderPath, "EmptyDocumentFirst.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.BuiltinDocumentProperties.Title = "This is my title";
                document.BuiltinDocumentProperties.Creator = "Przemysław Kłys";
                document.BuiltinDocumentProperties.Keywords = "word, docx, test";

                document.AddParagraph("This is my test");

                document.Save(openWord);
            }
        }

        public static void Example_BasicDocumentSaveAs1(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with SaveAs");
            string filePath = System.IO.Path.Combine(folderPath, "EmptyDocument1.docx");
            string filePathOutput = System.IO.Path.Combine(folderPath, "EmptyDocumentSaveAs1.docx");
            string filePathOutput2 = System.IO.Path.Combine(folderPath, "EmptyDocumentSaveAs2.docx");

            WordDocument document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "This is my title";
            document.BuiltinDocumentProperties.Creator = "Przemysław Kłys";
            document.BuiltinDocumentProperties.Keywords = "word, docx, test";
            document.AddParagraph("This is my test in document");

            // We're checking if the file is locked (it should be)
            Console.WriteLine("File: " + filePath + " is locked: " + filePath.IsFileLocked());

            // We're checking if the file is locked (it shouldn't be - yet)
            Console.WriteLine("File: " + filePathOutput + " is locked: " + filePathOutput.IsFileLocked());

            document.Save(filePathOutput, false);

            // both files should not be locked
            Console.WriteLine("File: " + filePath + " is locked: " + filePath.IsFileLocked());
            Console.WriteLine("File: " + filePathOutput + " is locked: " + filePathOutput.IsFileLocked());

            WordDocument document1 = WordDocument.Load(filePath);

            document1.AddParagraph("This is my test in document 2");
            document1.Save(filePathOutput2, openWord);
        }


        public static void Example_BasicDocumentSaveAs2(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with SaveAs");
            string filePath = System.IO.Path.Combine(folderPath, "FirstDocument1.docx");

            WordDocument document = WordDocument.Create();
            document.BuiltinDocumentProperties.Title = "This is my title";
            document.BuiltinDocumentProperties.Creator = "Przemysław Kłys";
            document.BuiltinDocumentProperties.Keywords = "word, docx, test";

            document.AddParagraph("This is my test in document");

            document.Save(filePath, openWord);
        }

        public static void Example_BasicDocumentSaveAs3(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with SaveAs");
            string filePath1 = System.IO.Path.Combine(folderPath, "FirstDocument11.docx");
            string filePath2 = System.IO.Path.Combine(folderPath, "FirstDocument12.docx");
            string filePath3 = System.IO.Path.Combine(folderPath, "FirstDocument13.docx");

            WordDocument document = WordDocument.Create();
            document.BuiltinDocumentProperties.Title = "This is my title";
            document.BuiltinDocumentProperties.Creator = "Przemysław Kłys";
            document.BuiltinDocumentProperties.Keywords = "word, docx, test";

            document.AddParagraph("This is my test in document 1");

            document.Save(filePath1);

            document.AddParagraph("This is my test in document 2");

            document.Save(filePath2);

            document.AddParagraph("This is my test in document 3");

            document.Save(filePath3, openWord);
        }
    }
}
