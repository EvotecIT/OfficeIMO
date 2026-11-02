using System.Text;
using OfficeIMO.Word;

namespace OfficeIMO.Tests;

public sealed class OfficeConverterSourceFormatTests {
    [Fact]
    public async Task SourceOverrideUsesTheRequestedCodecWithoutChangingFileIdentityOrSavedJobMeaning() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-source-format-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string input = Path.Combine(root, "review.txt"), output = Path.Combine(root, "review.docx");
            byte[] bytes = Encoding.UTF8.GetBytes("# Forced heading\n\nPreserved source identity.");
            File.WriteAllBytes(input, bytes);
            var options = new OfficeConversionOptions { SourceFormat = OfficeDocumentFormat.Markdown };
            OfficeConversionReport analyzed = await OfficeConverter.AnalyzeAsync(input, output, options);
            Assert.Equal(input, analyzed.Source);
            Assert.False(File.Exists(output));
            OfficeConversionReport converted = await OfficeConverter.ConvertAsync(input, output, options);
            Assert.Equal(input, converted.Source);
            using (var document = WordDocument.Load(output)) Assert.Contains(document.Paragraphs, paragraph => paragraph.Text == "Forced heading");
            using var html = new MemoryStream();
            var streamReport = await OfficeConverter.ConvertAsync(input, html, OfficeDocumentFormat.Html, options);
            Assert.Equal(input, streamReport.Source);
            Assert.Contains("<h1", Encoding.UTF8.GetString(html.ToArray()));
            Assert.DoesNotContain("# Forced heading", Encoding.UTF8.GetString(html.ToArray()));
            string jobPath = Path.Combine(root, "review.officejob.json");
            new OfficeBatchOptions { InputPaths = [input], OutputDirectory = Path.Combine(root, "word"), TargetExtension = ".docx", ConversionOptions = options }.Save(jobPath);
            var job = OfficeBatchOptions.Load(jobPath);
            Assert.Equal(OfficeDocumentFormat.Markdown, job.ConversionOptions.SourceFormat);
            var plan = await OfficeConverter.PlanAsync(job);
            Assert.Equal("markdown-docx", Assert.Single(plan.Items).RouteId);
            var batch = await OfficeConverter.ConvertAsync(job);
            Assert.Equal(1, batch.Summary.Completed);
            Assert.Equal(0, batch.Summary.Failed);
            using (var document = WordDocument.Load(Assert.Single(batch.Items).OutputPath!))
                Assert.Contains(document.Paragraphs, paragraph => paragraph.Text == "Forced heading");
            Assert.Equal(bytes, File.ReadAllBytes(input));
            Assert.Null(options.SourceExtension);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task WordFamilyOverrideRetainsTheLegacyCodecChosenByAnExplicitRoute() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-source-family-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string source = Path.Combine(root, "review.doc"), output = Path.Combine(root, "review.pdf");
            using (var document = WordDocument.Create()) { document.AddParagraph("Native legacy family"); document.Save(source); }
            var report = await OfficeConverter.ConvertAsync(source, output, new OfficeConversionOptions {
                SourceFormat = OfficeDocumentFormat.Word, RouteId = "doc-pdf", WordLoad = new WordLoadOptions { FailOnLoss = false }
            });
            Assert.Equal(output, report.OutputPath);
            using var pdf = OfficeIMO.Pdf.PdfDocument.Load(output);
            Assert.Contains("Native legacy family", pdf.Read().Text);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void ContradictorySourceFamilyAndCodecAreRejectedBeforeConversion() {
        var options = new OfficeConversionOptions { SourceFormat = OfficeDocumentFormat.Markdown, SourceExtension = ".html" };
        Assert.Throws<ArgumentException>(() => options.Clone());
    }
}
