using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Markdown;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Markdown {
    internal static partial class Markdown {
        public static async Task Example_MarkdownInterfaceAsync(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "MarkdownInterfaceAsync.docx");
            string markdown = "# Title\nContent";
            using MemoryStream input = new MemoryStream(Encoding.UTF8.GetBytes(markdown));
            using MemoryStream output = new MemoryStream();
            ConverterRegistry.Register("markdown->word", () => new MarkdownToWordConverter());
            IWordConverter converter = ConverterRegistry.Resolve("markdown->word");
            using CancellationTokenSource cts = new CancellationTokenSource();
            await converter.ConvertAsync(input, output, new MarkdownToWordOptions(), cts.Token);
            await File.WriteAllBytesAsync(filePath, output.ToArray());
            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
