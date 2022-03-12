using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal partial class CustomAndBuiltinProperties {
        public static void Example_ReadWord(bool openWord) {
            Console.WriteLine("[*] Read Basic Word");

            string documentPaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Templates");

            WordDocument document = WordDocument.Load(System.IO.Path.Combine(documentPaths, "BasicDocument.docx"), true);

            Console.WriteLine("This document has " + document.Paragraphs.Count + " paragraphs. Cool right?");
            Console.WriteLine("+ Document Title: " + document.BuiltinDocumentProperties.Title);
            Console.WriteLine("+ Document Author: " + document.BuiltinDocumentProperties.Creator);
            Console.WriteLine("+ FileOpen: " + document.FileOpenAccess);

            document.Dispose();
        }
    }
}
