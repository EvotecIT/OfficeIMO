using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class BasicDocument {
        public static void Example_BasicWordWithMarginsInCentimeters(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with margins");
            string filePath = System.IO.Path.Combine(folderPath, "EmptyDocumentWithMarginsCentimeters.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.Sections[0].Margins.BottomCentimeters = 2.30;
                document.Sections[0].Margins.TopCentimeters = 5.50;
                document.Sections[0].Margins.LeftCentimeters = 3.01;
                document.Sections[0].Margins.RightCentimeters = 3.05;

                Console.WriteLine(document.Sections[0].Margins.BottomCentimeters);
                Console.WriteLine(document.Sections[0].Margins.TopCentimeters);
                Console.WriteLine(document.Sections[0].Margins.LeftCentimeters);
                Console.WriteLine(document.Sections[0].Margins.RightCentimeters);

                Console.WriteLine(document.Sections[0].Margins.Bottom);
                Console.WriteLine(document.Sections[0].Margins.Top);
                Console.WriteLine(document.Sections[0].Margins.Left.Value);
                Console.WriteLine(document.Sections[0].Margins.Right.Value);

                //document.Sections[0].Margins.Bottom = 10;
                //document.Sections[0].Margins.Top = 10;
                //document.Sections[0].Margins.Left = 600;
                //document.Sections[0].Margins.Right = 600;

                document.Settings.FontFamily = "Arial";
                document.Settings.FontSize = 9;

                document.AddParagraph("This should be Arial 9");

                var par = document.AddParagraph("This should be Tahoma 20");
                par.FontFamily = "Tahoma";
                par.FontSize = 20;

                document.Save(openWord);
            }
        }
    }
}
