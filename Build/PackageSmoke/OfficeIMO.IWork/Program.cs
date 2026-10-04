using System.Security.Cryptography;
using OfficeIMO.IWork;
using OfficeIMO.Reader;
using OfficeIMO.Reader.IWork;
using OfficeIMO.Workflows;
using OfficeIMO.Workflows.IWork;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build();
OfficeWorkflowRunner runner = IWorkWorkflow.CreateRunner(conversionOptions: new IWorkConversionOptions {
    Mode = IWorkConversionMode.EditableOnly,
    NormalizeWorksheetNames = true
});
string outputRoot = Path.Combine(Path.GetTempPath(), "officeimo-iwork-package-smoke-" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(outputRoot);
try {
    foreach ((string extension, string route, string target, string text) in new[] {
        ("pages", "pages-docx", "docx", "hello pages"),
        ("numbers", "numbers-xlsx", "xlsx", "Z"),
        ("key", "keynote-pptx", "pptx", "hello keynote")
    }) {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "simple." + extension);
        byte[] payload = File.ReadAllBytes(path);
        string hash = Convert.ToHexString(SHA256.HashData(payload)).ToLowerInvariant();
        OfficeDocumentReadResult read = await reader.ReadDocumentAsync(path);
        if (read.Kind != ReaderInputKind.IWork || read.Source.SourceHash != hash
            || read.Source.LengthBytes != payload.LongLength
            || read.Chunks.Any(chunk => chunk.SourceHash != hash || chunk.SourceLengthBytes != payload.LongLength)
            || !read.Chunks.Any(chunk => chunk.Text.Contains(text, StringComparison.Ordinal))) {
            throw new InvalidOperationException("Packed Reader ingestion failed for " + extension + ".");
        }
        if (extension == "key") {
            IWorkKeynoteProjection source = IWorkSourceDocument.Open(path).ReadKeynote();
            IWorkTextBox body = source.Slides[0].TextBoxes.Single(box =>
                box.Content.PlainText.Contains("first bullet", StringComparison.Ordinal));
            IWorkTextBoxLayout? frame = body.Layout;
            IWorkListLayout? list = body.Content.Paragraphs[0].ListLayout;
            if (body.Geometry is not { WidthPoints: > 0, HeightPoints: > 0 }
                || frame is not { ShrinkToFit: true, LeftInsetPoints: 4, TopInsetPoints: 4,
                    RightInsetPoints: 4, BottomInsetPoints: 4 }
                || list is null || list.TextIndentEm != 1 || list.MarkerIndentPoints != 0
                || Math.Abs(list.MarkerScale - 1.23) > 0.000001) {
                throw new InvalidOperationException("Packed Keynote source APIs lost frame or list layout.");
            }
        }
        string output = Path.Combine(outputRoot, "converted." + target);
        OfficeWorkflowResult result = await runner.RunAsync(new OfficeWorkflowRequest {
            InputPath = path, OutputPath = output, Operation = OfficeWorkflowOperation.Convert,
            ConversionRouteId = route
        });
        if (!result.Succeeded || !File.Exists(output)
            || !result.Diagnostics.Any(diagnostic => diagnostic.Code == "OutputReopened")
            || result.ConversionEvidence?.Facts["projectionKind"] != "EditableReconstruction"
            || result.ConversionEvidence.Facts["partialEditableReconstruction"] != bool.FalseString
            || result.ConversionEvidence.Facts["omittedSourceUnitCount"] != "0"
            || result.ConversionEvidence.Facts["unassessedSourceUnitCount"] != "0") {
            throw new InvalidOperationException("Packed conversion failed for " + route + ": " + result.Summary);
        }
        switch (extension) {
            case "pages":
                using (var document = OfficeIMO.Word.WordDocument.Load(output)) {
                    if (!document.Paragraphs.Any(paragraph => paragraph.Text == "hello pages"))
                        throw new InvalidOperationException("Packed Pages conversion lost its body text.");
                }
                break;
            case "numbers":
                using (var workbook = OfficeIMO.Excel.ExcelDocument.Load(output)) {
                    if (workbook.Sheets.Count != 1 || workbook.Sheets[0].CellAt(1, 1).GetValue<string>() != "a"
                        || workbook.Sheets[0].CellAt(2, 2).GetValue<double>() != 2d
                        || workbook.Sheets[0].CellAt(3, 3).GetValue<string>() != "Z")
                        throw new InvalidOperationException("Packed Numbers conversion lost its typed grid values.");
                }
                break;
            case "key":
                using (var presentation = OfficeIMO.PowerPoint.PowerPointPresentation.Load(output)) {
                    if (presentation.Slides.Count != 2
                        || !presentation.Slides[0].TextBoxes.Any(box => box.Text.Contains("hello keynote", StringComparison.Ordinal))
                        || !presentation.Slides[1].TextBoxes.Any(box => box.Text.Contains("second slide", StringComparison.Ordinal)))
                        throw new InvalidOperationException("Packed Keynote conversion lost slide content or order.");
                }
                break;
        }
    }
} finally {
    Directory.Delete(outputRoot, recursive: true);
}
Console.WriteLine("OfficeIMO iWork Reader and conversion package smoke passed on "
    + System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription + ".");
