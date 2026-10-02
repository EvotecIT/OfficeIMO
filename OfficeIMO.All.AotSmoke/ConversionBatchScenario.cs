using System;
using System.IO;
using System.Threading.Tasks;
using OfficeIMO.Pdf;
using OfficeIMO.Workflows;

internal static class ConversionBatchScenario {
    internal static async Task RunAsync(string root) {
        string input = Path.Combine(root, "batch-input");
        string output = Path.Combine(root, "batch-output");
        string state = Path.Combine(root, "batch-state");
        Directory.CreateDirectory(input);
        File.WriteAllText(Path.Combine(input, "literal.txt"), "Native checkpoint text");
        File.WriteAllText(Path.Combine(input, "page.html"), "<p>Native checkpoint HTML</p>");
        File.WriteAllText(Path.Combine(input, "page.md"), "# Native checkpoint Markdown");
        File.WriteAllText(Path.Combine(input, "page.rtf"), "{\\rtf1\\ansi Native checkpoint RTF}");
        using (var word = OfficeIMO.Word.WordDocument.Create(Path.Combine(input, "page.docx"))) {
            word.AddParagraph("Native checkpoint Word"); word.Save();
        }
        using (var excel = OfficeIMO.Excel.ExcelDocument.Create(Path.Combine(input, "page.xlsx"))) {
            excel.AddWorksheet("Data").Cell(1, 1, "Native checkpoint Excel"); excel.Save();
        }
        using (var slides = OfficeIMO.PowerPoint.PowerPointPresentation.Create(Path.Combine(input, "page.pptx"))) {
            slides.AddSlide().AddTextBoxPoints("Native checkpoint slide", 40, 40, 500, 60); slides.Save();
        }
        const string password = "Native smoke output password";
        var request = new OfficeConversionBatchRequest {
            InputDirectory = input, OutputDirectory = output, CheckpointDirectory = state,
            ConversionOptions = new() {
                PlainText = new() { PdfOptions = new PdfOptions().SetEncryption(password, "Native smoke owner password") },
                Word = new() { Title = "Native Word" }, Excel = new() { IncludeSheetHeadings = true },
                PowerPoint = new() { IncludeHiddenSlides = true }, Html = new() { InteractiveFormControls = false },
                Markdown = new() { Title = "Native Markdown" }, Rtf = new() { IncludeHiddenText = true }
            }
        };
        var runner = new OfficeWorkflowRunner();
        var items = new BatchItems();
        OfficeConversionBatchResult first = await runner.RunBatchAsync(request, items);
        if (first.Completed != 7 || first.Failed != 0) throw new InvalidOperationException("Native checkpoint conversion failed: " + items.Summary);
        foreach (string file in Directory.GetFiles(output, "*.pdf")) {
            var pdf = PdfDocument.Load(file, new PdfLoadOptions { Password = password });
            if (pdf.Read().Pages.Count == 0) throw new InvalidOperationException("Native batch produced an empty PDF.");
        }
        OfficeConversionBatchResult resumed = await runner.RunBatchAsync(request);
        if (resumed.Reused != 7 || resumed.Failed != 0) throw new InvalidOperationException("Native checkpoint reuse failed.");
        Console.WriteLine("PASS | Native conversion batch: seven format option bags, encrypted TXT, verified checkpoint reuse.");
    }

    private sealed class BatchItems : IProgress<OfficeConversionBatchItemResult> {
        internal string Summary = string.Empty;
        public void Report(OfficeConversionBatchItemResult item) {
            if (item.Status != OfficeWorkflowStatus.Completed) lock (this) Summary += item.Summary + "\n";
        }
    }
}
