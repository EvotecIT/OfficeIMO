using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class BasicDocument {
        public static async Task Example_BasicWordAsync(string folderPath) {
            Console.WriteLine("[*] Async example for WordDocument");
            string filePath = Path.Combine(folderPath, "AsyncWord.docx");
            if (File.Exists(filePath)) File.Delete(filePath);

            using (var document = await WordDocument.CreateAsync(filePath, cancellationToken: CancellationToken.None)) {
                document.AddParagraph("Async paragraph");
                await document.SaveAsync();
            }

            using (var document = await WordDocument.LoadAsync(filePath, cancellationToken: CancellationToken.None)) {
                Console.WriteLine($"Paragraph count: {document.Paragraphs.Count}");
            }

            File.Delete(filePath);
        }
    }
}
