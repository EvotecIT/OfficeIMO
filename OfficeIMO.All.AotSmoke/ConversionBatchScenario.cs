using OfficeIMO;
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
            excel.AddSheet("Data").Cell(1, 1, "Native checkpoint Excel"); excel.Save();
        }
        using (var slides = OfficeIMO.PowerPoint.PowerPointDocument.Create(Path.Combine(input, "page.pptx"))) {
            slides.AddSlide().AddTextBox("Native checkpoint slide",global::OfficeIMO.OfficeLength.Points(40),global::OfficeIMO.OfficeLength.Points(40),global::OfficeIMO.OfficeLength.Points(500),global::OfficeIMO.OfficeLength.Points(60)); slides.Save();
        }
        const string password = "Native smoke output password";
        var request = new OfficeBatchOptions {
            InputDirectory = input, OutputDirectory = output, ResumeFrom = Path.Combine(state, "batch.json"),
            ConversionOptions = new() {
                PlainText = new() { PdfOptions = new PdfCreateOptions().SetEncryption(password, "Native smoke owner password") },
                Word = new() { Title = "Native Word" },
                Excel = new() { IncludeSheetHeadings = true },
                PowerPoint = new() { IncludeHiddenSlides = true },
                Html = new() { InteractiveFormControls = false },
                Markdown = new() { Title = "Native Markdown" },
                RegisteredSettings = new OfficeWorkflowRouteSettings { Rtf = new() { IncludeHiddenText = true, PdfOptions = new PdfCreateOptions { DefaultFontSize = OfficeLength.Points(14) } } }
            }
        };
        OfficeWorkflow.RegisterConversions();
        var items = new BatchItems();
        OfficeBatchReport first = await OfficeConverter.ConvertAsync(request, items);
        if (first.Summary.Completed != 7 || first.Summary.Failed != 0) throw new InvalidOperationException("Native checkpoint conversion failed: " + items.Summary);
        foreach (string file in Directory.GetFiles(output, "*.pdf")) {
            using var pdf = PdfDocument.Load(file, new PdfLoadOptions { Password = password });
            if (pdf.Read().Pages.Count == 0) throw new InvalidOperationException("Native batch produced an empty PDF.");
        }
        OfficeBatchReport resumed = await OfficeConverter.ConvertAsync(request);
        if (resumed.Summary.Reused != 7 || resumed.Summary.Failed != 0) throw new InvalidOperationException("Native checkpoint reuse failed.");
        OfficeConversionOptions changedText = request.ConversionOptions.Clone();
        changedText.PlainText!.PdfOptions!.DefaultFontSize += OfficeLength.Emus(1);
        OfficeBatchReport textRejected = await OfficeConverter.ConvertAsync(request with { ConversionOptions = changedText });
        if (textRejected.Summary.Failed != 1 || textRejected.Summary.Reused != 6)
            throw new InvalidOperationException("Native checkpoint reused changed text geometry.");
        OfficeConversionOptions changedRtf = request.ConversionOptions.Clone();
        var registered = (OfficeWorkflowRouteSettings)changedRtf.RegisteredSettings!;
        registered.Rtf!.PdfOptions!.DefaultFontSize += OfficeLength.Emus(1);
        OfficeBatchReport rtfRejected = await OfficeConverter.ConvertAsync(request with { ConversionOptions = changedRtf });
        if (rtfRejected.Summary.Failed != 1 || rtfRejected.Summary.Reused != 6)
            throw new InvalidOperationException("Native checkpoint reused changed registered-route geometry.");
        Console.WriteLine("PASS | Native conversion batch: seven format option bags, encrypted TXT, verified reuse and exact physical settings rejection.");
    }

    private sealed class BatchItems : IProgress<OfficeBatchItemReport> {
        internal string Summary = string.Empty;
        public void Report(OfficeBatchItemReport item) {
            if (item.Status != OfficeBatchItemStatus.Completed) lock (this) Summary += item.Summary + "\n";
        }
    }
}
