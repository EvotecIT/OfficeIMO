using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class BasicDocument {
        public static void Example_BasicWordWithMarginsInCentimeters(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with margins");
            string filePath = System.IO.Path.Combine(folderPath, "EmptyDocumentWithMarginsCentimeters.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.Sections[0].PageSetup.Margins.Bottom = global::OfficeIMO.OfficeLength.Centimeters(2.3D);
                document.Sections[0].PageSetup.Margins.Top = global::OfficeIMO.OfficeLength.Centimeters(5.5D);
                document.Sections[0].PageSetup.Margins.Left = global::OfficeIMO.OfficeLength.Centimeters(3.01D);
                document.Sections[0].PageSetup.Margins.Right = global::OfficeIMO.OfficeLength.Centimeters(3.05D);

                Console.WriteLine(document.Sections[0].PageSetup.Margins.Bottom?.ToCentimeters());
                Console.WriteLine(document.Sections[0].PageSetup.Margins.Top?.ToCentimeters());
                Console.WriteLine(document.Sections[0].PageSetup.Margins.Left?.ToCentimeters());
                Console.WriteLine(document.Sections[0].PageSetup.Margins.Right?.ToCentimeters());

                Console.WriteLine(checked((int?)(document.Sections[0].PageSetup.Margins.Bottom?.ToTwips())));
                Console.WriteLine(checked((int?)(document.Sections[0].PageSetup.Margins.Top?.ToTwips())));
                Console.WriteLine(checked((uint)(document.Sections[0].PageSetup.Margins.Left.GetValueOrDefault().ToTwips())));
                Console.WriteLine(checked((uint)(document.Sections[0].PageSetup.Margins.Right.GetValueOrDefault().ToTwips())));


                document.Settings.FontFamily = "Arial";
                document.Settings.FontSize = global::OfficeIMO.OfficeLength.Points(9);

                document.AddParagraph("This should be Arial 9");

                var par = document.AddParagraph("This should be Tahoma 20");
                par.FontFamily = "Tahoma";
                par.FontSize = global::OfficeIMO.OfficeLength.Points(20);

                document.Save();
                if (openWord) document.OpenInApplication();
            }
        }
    }
}
