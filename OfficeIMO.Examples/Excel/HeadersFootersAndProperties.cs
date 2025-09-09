using System;
using System.IO;
using OfficeIMO.Excel;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    internal static class HeadersFootersAndProperties {
        public static void Example(string folderPath, bool openExcel = false) {
            string filePath = Path.Combine(folderPath, "HeadersFootersAndProperties.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            using var doc = ExcelDocument.Create(filePath);

            // Built-in document properties
            doc.BuiltinDocumentProperties.Title = "Scan Summary";
            doc.BuiltinDocumentProperties.Creator = "OfficeIMO";
            doc.ApplicationProperties.Company = "Evotec";

            var sheet = doc.AddWorkSheet("Summary");

            // Header/Footer text (tokens: &P page, &N pages, &A sheet name)
            sheet.SetHeaderFooter(
                headerLeft: "&A",
                headerCenter: "Scan Summary",
                headerRight: "Page &P of &N",
                footerLeft: "Generated: &D &T",
                footerCenter: null,
                footerRight: "© Evotec");

            // Add a small logo to the center header
            var logoPath = Path.Combine(AppContext.BaseDirectory, "Assets", "OfficeIMO.png");
            if (File.Exists(logoPath)) {
                var bytes = File.ReadAllBytes(logoPath);
                sheet.SetHeaderImage(HeaderFooterPosition.Center, bytes, contentType: "image/png", widthPoints: 96, heightPoints: 32);
            } else {
                Console.WriteLine($"Logo file not found: {logoPath}");
            }

            sheet.Cell(1, 1, "Hello");
            sheet.Cell(2, 1, "World");

            doc.Save(openExcel);
        }
    }
}

