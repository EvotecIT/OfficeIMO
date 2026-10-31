using System.Threading.Tasks;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfNativeSaveEncryptionTests {
    [Theory]
    [InlineData("bytes", false)]
    [InlineData("file", false)]
    [InlineData("file-async", false)]
    [InlineData("stream", false)]
    [InlineData("stream-async", false)]
    [InlineData("bytes", true)]
    [InlineData("file", true)]
    [InlineData("file-async", true)]
    [InlineData("stream", true)]
    [InlineData("stream-async", true)]
    public async Task OutputEncryptionHonorsNativePermissionsAndKeepsTheLiveModelUnchanged(string sink, bool loaded) {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-save-security-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            using PdfDocument generated = PdfDocument.Create(builder => builder.Content(content => content.Paragraph(paragraph => paragraph.Text("Native encrypted output"))));
            using PdfDocument? loadedDocument = loaded ? PdfDocument.Load(generated.ToBytes(new PdfSaveOptions { AllowLoss = true })) : null;
            PdfDocument document = loadedDocument ?? generated;
            var options = new PdfSaveOptions { AllowLoss = true,
                Encryption = PdfStandardEncryptionOptions.ForPermissions("fixture-owner", PdfStandardPermissions.Print | PdfStandardPermissions.Accessibility) };
            string path = Path.Combine(directory, "Saved.pdf");
            using var output = new MemoryStream();
            byte[] bytes;
            switch (sink) {
                case "bytes": bytes = document.ToBytes(options); break;
                case "file": document.Save(path, options); bytes = File.ReadAllBytes(path); break;
                case "file-async": await document.SaveAsync(path, options); bytes = File.ReadAllBytes(path); break;
                case "stream": document.Save(output, options); bytes = output.ToArray(); break;
                default: await document.SaveAsync(output, options); bytes = output.ToArray(); break;
            }
            using (PdfDocument reader = PdfDocument.Load(bytes)) Assert.Throws<PdfPermissionDeniedException>(() => reader.Read());
            using (PdfDocument owner = PdfDocument.Load(bytes, new PdfLoadOptions { Password = "fixture-owner" }))
                Assert.Contains("Native encrypted output", owner.Read().Text);
            Assert.Equal(sink.StartsWith("file", StringComparison.Ordinal) ? path : null, document.FilePath);
            using PdfDocument unencrypted = PdfDocument.Load(document.ToBytes(new PdfSaveOptions { AllowLoss = true }));
            Assert.Contains("Native encrypted output", unencrypted.Read().Text);
        } finally { Directory.Delete(directory, recursive: true); }
    }
    [Theory]
    [InlineData("pdfa")]
    [InlineData("pdfx")]
    [InlineData("invoice")]
    public void OutputEncryptionCannotRetainConformanceIdentification(string kind) {
        var createOptions = new PdfCreateOptions();
        if (kind == "pdfa") createOptions.SetPdfAIdentification(3, "B");
        else if (kind == "pdfx") createOptions.SetPdfXIdentification(PdfXIdentification.PdfX4());
        else createOptions.SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX());
        using PdfDocument document = PdfDocument.Create(builder => builder.Content(content =>
            content.Paragraph(paragraph => paragraph.Text("Conformance identification"))), createOptions);
        var save = new PdfSaveOptions { AllowLoss = true,
            Encryption = PdfStandardEncryptionOptions.ForPermissions("fixture-owner", PdfStandardPermissions.All) };
        byte[] source = document.ToBytes(new PdfSaveOptions { AllowLoss = true });
        using PdfDocument loaded = PdfDocument.Load(source);
        Assert.Throws<ArgumentException>(() => document.ToBytes(save));
        Assert.Throws<ArgumentException>(() => loaded.ToBytes(save));
        string root = Path.Combine(Path.GetTempPath(), "officeimo-pdf-conformance-security-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string path = Path.Combine(root, "Preserved.pdf");
            File.WriteAllBytes(path, source);
            Assert.Throws<ArgumentException>(() => loaded.Save(path, save));
            Assert.Equal(source, File.ReadAllBytes(path));
            Assert.Single(Directory.GetFiles(root));
        } finally { Directory.Delete(root, recursive: true); }
    }

}
