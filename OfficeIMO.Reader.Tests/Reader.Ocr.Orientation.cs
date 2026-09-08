using OfficeIMO.Ocr;
using OfficeIMO.Ocr.Tesseract;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderOcrOrientationTests {
    [Theory]
    [InlineData(OcrOperation.DetectOrientation, true)]
    [InlineData(OcrOperation.DetectOrientation, false)]
    [InlineData(OcrOperation.RecognizeText, true)]
    [InlineData(OcrOperation.RecognizeText, false)]
    public async Task FailedProviderRequestsHonorTemporaryFileRetention(OcrOperation operation, bool retain) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-osd-retention-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            var engine = new TesseractOcrEngine(new TesseractOcrEngineOptions {
                ExecutablePath = Path.Combine(root, "missing-tesseract.exe"),
                TemporaryDirectory = root, KeepTemporaryFiles = retain
            });
            byte[] payload = { 1, 2, 3 };
            OcrResult? result = null;
            Exception? failure = await Record.ExceptionAsync(async () => result = await engine.RecognizeAsync(new OcrRequest {
                Operation = operation, Payload = payload, MediaType = "image/png", FileName = "source.png"
            }));
            // Unix process-group launch can start setsid successfully and then report the missing
            // executable as an exit status. Retention must hold for both launch-failure paths.
            if (failure is not null) {
                if (failure is System.ComponentModel.Win32Exception launchFailure) Assert.Equal(2, launchFailure.NativeErrorCode);
                else {
                    Assert.IsType<InvalidOperationException>(failure);
                    Assert.Contains("missing-tesseract", failure.Message, StringComparison.Ordinal);
                }
            } else {
                Assert.Equal(OcrOperation.DetectOrientation, operation);
                Assert.NotNull(result);
                Assert.Null(result!.Orientation);
                Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "tesseract-orientation-unavailable");
            }
            string[] retained = Directory.GetFiles(root, "input.png", SearchOption.AllDirectories);
            if (retain) Assert.Equal(payload, File.ReadAllBytes(Assert.Single(retained)));
            else Assert.Empty(Directory.GetDirectories(root));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Theory]
    [InlineData(0, "15.0", 1D)]
    [InlineData(90, "7.5", 0.5D)]
    [InlineData(180, "28.28", 1D)]
    [InlineData(270, "9.30", 0.62D)]
    public void OrientationParserReadsCorrectiveRotationAndCalibratedEvidence(int rotation, string confidence, double expected) {
        OcrOrientationResult? result = TesseractOcrEngine.ParseOrientation(
            "Page number: 0\r\nOrientation in degrees: 90\r\nRotate: " + rotation +
            "\r\nOrientation confidence: " + confidence + "\r\nScript: Latin\r\nScript confidence: 8.06\r\n");
        Assert.NotNull(result);
        Assert.Equal(rotation, result!.ClockwiseRotationDegrees);
        Assert.Equal(expected, result.Confidence, 6);
        Assert.Equal("Latin", result.Script);
    }

    [Theory]
    [InlineData("Rotate: 45\nOrientation confidence: 15")]
    [InlineData("Rotate: 90\nOrientation confidence: NaN")]
    [InlineData("Rotate: 90\nOrientation confidence: -1")]
    [InlineData("Rotate: 90\nOrientation confidence: Infinity")]
    [InlineData("Rotate: 90")]
    [InlineData("Rotate: 90\nRotate: 270\nOrientation confidence: 15")]
    [InlineData("Rotate: 90\nOrientation confidence: 15\nOrientation confidence: invalid")]
    public void OrientationParserRejectsIncompleteInvalidOrConflictingEvidence(string output) {
        Assert.Null(TesseractOcrEngine.ParseOrientation(output));
    }
}
