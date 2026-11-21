using CodeGlyphX;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Barcodes {
    /// <summary>QR codes and barcodes placed with the optional CodeGlyphX adapters, one call per format.</summary>
    internal static class BarcodeExamples {
        public static void Example(string folderPath) {
            Directory.CreateDirectory(folderPath);
            WordCodes(Path.Combine(folderPath, "Barcodes.docx"));
            ExcelCodes(Path.Combine(folderPath, "Barcodes.xlsx"));
            PowerPointCodes(Path.Combine(folderPath, "Barcodes.pptx"));
            PdfCodes(Path.Combine(folderPath, "Barcodes.pdf"));
        }

        private static void WordCodes(string path) {
            Console.WriteLine("[*] Word: QR code and Code 128 barcode");
            using var document = WordDocument.Create(path);
            document.AddParagraph("Scan to open the event agenda:");
            document.AddQrCode("https://evotec.xyz/events");
            document.AddParagraph("Order number:");
            document.AddBarcode(SymbolFormat.Code128, "ORD-2026-0042", options => options.Width = OfficeLength.Millimeters(60));
            document.Save();
        }

        private static void ExcelCodes(string path) {
            Console.WriteLine("[*] Excel: product QR codes and an EAN-13 barcode");
            using var workbook = ExcelDocument.Create(path);
            ExcelSheet sheet = workbook.AddSheet("Products");
            sheet.Cell(1, 1).SetValue("Product");
            sheet.Cell(1, 2).SetValue("Details");
            sheet.Cell(2, 1).SetValue("Contoso Router X2");
            sheet.AddQrCode("B2", "https://contoso.example/x2", options => options.Width = OfficeLength.Centimeters(2));
            sheet.Cell(8, 1).SetValue("Retail code");
            sheet.AddBarcode("B8", SymbolFormat.Ean, "590123412345");
            workbook.Save();
        }

        private static void PowerPointCodes(string path) {
            Console.WriteLine("[*] PowerPoint: QR code fitted to a layout box");
            using var presentation = PowerPointDocument.Create(path);
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddTextBox("Slides and recording", OfficeLength.Inches(0.75), OfficeLength.Inches(0.6), OfficeLength.Inches(8), OfficeLength.Inches(0.8));
            var box = new PowerPointLayoutBox(OfficeLength.Inches(3.5), OfficeLength.Inches(1.8), OfficeLength.Inches(3), OfficeLength.Inches(3));
            slide.AddQrCode("https://evotec.xyz/talks/officeimo", box);
            presentation.Save();
        }

        private static void PdfCodes(string path) {
            Console.WriteLine("[*] PDF: vector QR code, Data Matrix and GS1-128 barcode");
            PdfDocument.Create(pdf => pdf.Content(content => content
                    .H1("Shipping label")
                    .AddParagraph("Track this parcel:")
                    .AddQrCode("https://track.example.com/0042")
                    .AddParagraph("Serial number:")
                    .AddBarcode(SymbolFormat.DataMatrix, "SN-0042-PL", options => options.Width = OfficeLength.Millimeters(15))
                    .AddParagraph("GS1-128:")
                    .AddBarcode(SymbolFormat.Gs1Code128, "(01)09501101530003(17)261231")))
                .Save(path);
        }
    }
}
