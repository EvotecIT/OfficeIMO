using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace OfficeIMO.Examples.Word {
    internal static partial class Pdf {
        public static void Example_PdfCustomFonts(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating PDF with custom fonts");
            string docPath = Path.Combine(folderPath, "PdfCustomFonts.docx");
            string pdfPath = Path.Combine(folderPath, "PdfCustomFonts.pdf");
            string fontPath = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
                ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "arial.ttf")
                : RuntimeInformation.IsOSPlatform(OSPlatform.OSX)
                    ? "/System/Library/Fonts/Supplemental/Arial.ttf"
                    : "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf";

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("File font paragraph").FontFamily = "FileFont";
                document.AddParagraph("Stream font paragraph").FontFamily = "StreamFont";
                document.Save();
                using var fontStream = File.OpenRead(fontPath);
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    FontFilePaths = new Dictionary<string, string> { { "FileFont", fontPath } },
                    FontStreams = new Dictionary<string, Stream> { { "StreamFont", fontStream } }
                });
            }
        }
    }
}
