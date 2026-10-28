using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class BasicDocument {
        public static void Example_BasicWordWithMarginsInCentimeters(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with margins");
            string filePath = System.IO.Path.Combine(folderPath, "EmptyDocumentWithMarginsCentimeters.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.Sections[0].Margins.Bottom = global::OfficeIMO.OfficeLength.Centimeters(2.3D);
                document.Sections[0].Margins.Top = global::OfficeIMO.OfficeLength.Centimeters(5.5D);
                document.Sections[0].Margins.Left = global::OfficeIMO.OfficeLength.Centimeters(3.01D);
                document.Sections[0].Margins.Right = global::OfficeIMO.OfficeLength.Centimeters(3.05D);

                Console.WriteLine(document.Sections[0].Margins.Bottom?.ToCentimeters());
                Console.WriteLine(document.Sections[0].Margins.Top?.ToCentimeters());
                Console.WriteLine(document.Sections[0].Margins.Left?.ToCentimeters());
                Console.WriteLine(document.Sections[0].Margins.Right?.ToCentimeters());

                Console.WriteLine(checked((int?)(document.Sections[0].Margins.Bottom?.ToTwips())));
                Console.WriteLine(checked((int?)(document.Sections[0].Margins.Top?.ToTwips())));
                Console.WriteLine(checked((uint)(document.Sections[0].Margins.Left.GetValueOrDefault().ToTwips())));
                Console.WriteLine(checked((uint)(document.Sections[0].Margins.Right.GetValueOrDefault().ToTwips())));


                document.Settings.FontFamily = "Arial";
                document.Settings.FontSize = 9;

                document.AddParagraph("This should be Arial 9");

                var par = document.AddParagraph("This should be Tahoma 20");
                par.FontFamily = "Tahoma";
                par.FontSize = 20;

                document.Save();
                if (openWord) document.OpenInApplication();
            }
        }
    }
}
