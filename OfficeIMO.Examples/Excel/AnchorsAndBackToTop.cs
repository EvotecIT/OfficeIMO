using System;
using System.IO;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent;

namespace OfficeIMO.Examples.Excel {
    internal static class AnchorsAndBackToTop {
        public static void Example(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel: Section anchors + Back-to-Top");
            string filePath = Path.Combine(folderPath, "Excel.AnchorsAndBackToTop.xlsx");

            using (var doc = ExcelDocument.Create(filePath)) {
                var s = new SheetComposer(doc, "Report");
                s.Title("Anchors & Back-to-Top Demo", "Demonstrates SectionWithAnchor + top links.");

                // A few sections with anchors and a back-to-top link after each header
                s.SectionWithAnchor("Overview")
                 .Paragraph("Some overview text...")
                 .SectionWithAnchor("Findings")
                 .Paragraph("Key findings listed here...")
                 .SectionWithAnchor("Details")
                 .Paragraph("Supporting details go here...")
                 .Finish(autoFitColumns: true);

                // Add TOC for convenience
                doc.AddTableOfContents(placeFirst: true, withHyperlinks: true);
                doc.Save(false);
                if (openExcel) doc.Open(filePath, true);
            }
        }
    }
}

