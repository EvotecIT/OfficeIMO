using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Html;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static async Task Example_HtmlInterfaceAsync(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlInterfaceAsync.docx");
            string html = "<p>Hello world</p>";
            using MemoryStream input = new MemoryStream(Encoding.UTF8.GetBytes(html));
            using MemoryStream output = new MemoryStream();
            ConverterRegistry.Register("html->word", () => new HtmlToWordConverter());
            IWordConverter converter = ConverterRegistry.Resolve("html->word");
            using CancellationTokenSource cts = new CancellationTokenSource();
            await converter.ConvertAsync(input, output, new HtmlToWordOptions(), cts.Token);
            await File.WriteAllBytesAsync(filePath, output.ToArray());
            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
