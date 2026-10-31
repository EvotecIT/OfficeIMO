using System.Threading.Tasks;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfNativeSaveLossTests {
    [Theory]
    [InlineData(false, false, false)]
    [InlineData(false, false, true)]
    [InlineData(false, true, false)]
    [InlineData(false, true, true)]
    [InlineData(true, false, false)]
    [InlineData(true, false, true)]
    [InlineData(true, true, false)]
    [InlineData(true, true, true)]
    public async Task WriterLossBlocksEveryAttemptBeforeDestinationWritesAndAllowLossWritesReadableOutput(bool asynchronous, bool fileOutput, bool encrypt) {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-native-loss-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "document.pdf");
        byte[] original = [1, 2, 3]; File.WriteAllBytes(path, original);
        using var destination = new MemoryStream(); destination.Write(original, 0, original.Length);
        using PdfDocument document = PdfDocument.Create(builder => builder.Content(content => content.Canvas(canvas =>
            canvas.TextBox("First\nSecond\nThird", global::OfficeIMO.OfficeLength.Points(24D), global::OfficeIMO.OfficeLength.Points(24D), global::OfficeIMO.OfficeLength.Points(120D), global::OfficeIMO.OfficeLength.Points(18D),
                new PdfCanvasTextBoxStyle { FontSize = global::OfficeIMO.OfficeLength.Points(9D), PaddingX = global::OfficeIMO.OfficeLength.Points(0D), PaddingY = global::OfficeIMO.OfficeLength.Points(0D) }))));
        var options = new PdfSaveOptions { Encryption = encrypt ? PdfStandardEncryptionOptions.ForPermissions("owner-secret", PdfStandardPermissions.Print) : null };
        try {
            for (int attempt = 0; attempt < 2; attempt++) {
                OfficeConversionException error = await Assert.ThrowsAsync<OfficeConversionException>(async () => {
                    if (fileOutput) {
                        if (asynchronous) await document.SaveAsync(path, options);
                        else document.Save(path, options);
                    } else {
                        if (asynchronous) await document.SaveAsync(destination, options);
                        else document.Save(destination, options);
                    }
                });
                Assert.True(Assert.IsType<OfficeConversionReport>(error.Report).HasLoss);
                Assert.Equal(original, File.ReadAllBytes(path)); Assert.Equal(original, destination.ToArray());
            }
            var accepted = new PdfSaveOptions { AllowLoss = true, Encryption = options.Encryption };
            if (fileOutput) {
                if (asynchronous) await document.SaveAsync(path, accepted);
                else document.Save(path, accepted);
            } else {
                if (asynchronous) await document.SaveAsync(destination, accepted);
                else document.Save(destination, accepted);
                Assert.Equal(0, destination.Position); Assert.True(destination.CanWrite);
            }
            var readOptions = encrypt ? new PdfLoadOptions { Password = "owner-secret" } : null;
            using PdfDocument reopened = fileOutput ? PdfDocument.Load(path, readOptions) : PdfDocument.Load(new MemoryStream(destination.ToArray()), readOptions);
            Assert.Single(reopened.Pages); Assert.Contains("First", reopened.ExtractText());
        } finally { Directory.Delete(directory, recursive: true); }
    }
}
