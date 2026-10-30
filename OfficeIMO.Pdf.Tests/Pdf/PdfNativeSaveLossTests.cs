using System.Threading.Tasks;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfNativeSaveLossTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task WriterLossBlocksEveryAttemptBeforeDestinationWritesAndAllowLossWritesReadableOutput(bool asynchronous, bool fileOutput) {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-native-loss-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "document.pdf");
        byte[] original = [1, 2, 3]; File.WriteAllBytes(path, original);
        using var destination = new MemoryStream(); destination.Write(original, 0, original.Length);
        using PdfDocument document = PdfDocument.Create(builder => builder.Content(content => content.Canvas(canvas =>
            canvas.TextBox("First\nSecond\nThird", global::OfficeIMO.OfficeLength.Points(24D), global::OfficeIMO.OfficeLength.Points(24D), global::OfficeIMO.OfficeLength.Points(120D), global::OfficeIMO.OfficeLength.Points(18D),
                new PdfCanvasTextBoxStyle { FontSize = 9D, PaddingX = 0D, PaddingY = 0D }))));
        try {
            for (int attempt = 0; attempt < 2; attempt++) {
                OfficeConversionException error = await Assert.ThrowsAsync<OfficeConversionException>(async () => {
                    if (fileOutput) {
                        if (asynchronous) await document.SaveAsync(path);
                        else document.Save(path);
                    } else {
                        if (asynchronous) await document.SaveAsync(destination);
                        else document.Save(destination);
                    }
                });
                Assert.True(Assert.IsType<OfficeConversionReport>(error.Report).HasLoss);
                Assert.Equal(original, File.ReadAllBytes(path)); Assert.Equal(original, destination.ToArray());
            }
            var accepted = new PdfSaveOptions { AllowLoss = true };
            if (fileOutput) {
                if (asynchronous) await document.SaveAsync(path, accepted);
                else document.Save(path, accepted);
            } else {
                if (asynchronous) await document.SaveAsync(destination, accepted);
                else document.Save(destination, accepted);
                Assert.Equal(0, destination.Position); Assert.True(destination.CanWrite);
            }
            using PdfDocument reopened = fileOutput ? PdfDocument.Load(path) : PdfDocument.Load(new MemoryStream(destination.ToArray()));
            Assert.Single(reopened.Pages); Assert.Contains("First", reopened.ExtractText());
        } finally { Directory.Delete(directory, recursive: true); }
    }
}
