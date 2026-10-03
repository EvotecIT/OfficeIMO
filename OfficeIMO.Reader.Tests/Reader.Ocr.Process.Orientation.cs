using OfficeIMO.Ocr;
using OfficeIMO.Ocr.Process;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderOcrProcessOrientationTests {
    [Theory]
    [InlineData(OcrOperation.RecognizeText)]
    [InlineData(OcrOperation.DetectOrientation)]
    public async Task ProcessRequestPreservesOperationAndIndependentCapabilities(OcrOperation operation) {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-ocr-operation-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string response = Path.Combine(directory, "response.json"), captured = Path.Combine(directory, "captured.json");
            File.WriteAllText(response, ProcessOcrProtocol.SerializeResult(new OcrResult {
                Text = operation == OcrOperation.RecognizeText ? "Recognized" : string.Empty,
                Orientation = operation == OcrOperation.DetectOrientation ? new OcrOrientationResult { ClockwiseRotationDegrees = 90, Confidence = 0.9 } : null
            }));
            bool windows = RuntimeInformation.IsOSPlatform(OSPlatform.Windows);
            string script = Path.Combine(directory, windows ? "bridge.cmd" : "bridge.sh");
            File.WriteAllText(script, windows
                ? "@copy /Y \"%~1\" \"%~2\" >nul\r\n@copy /Y \"%~3\" \"%~4\" >nul\r\n"
                : "cp \"$1\" \"$2\"\ncp \"$3\" \"$4\"\n");
            var capabilities = new OcrEngineCapabilities { SupportsOrientationDetection = true, SupportedMediaTypes = new[] { "image/png" } };
            var engine = new ProcessOcrEngine(new ProcessOcrEngineOptions {
                FileName = windows ? Environment.GetEnvironmentVariable("ComSpec") ?? "cmd.exe" : "/bin/sh",
                Arguments = windows ? new[] { "/d", "/c", script, "{request}", captured, response, "{output}" }
                    : new[] { script, "{request}", captured, response, "{output}" },
                Capabilities = capabilities, TemporaryDirectory = directory
            });
            capabilities.SupportsOrientationDetection = false;
            engine.Capabilities.SupportsOrientationDetection = false;
            Assert.True(engine.Capabilities.SupportsOrientationDetection);

            OcrResult result = await OcrEngineRunner.RecognizeAsync(engine, new OcrRequest {
                Operation = operation, Payload = new byte[] { 1 }, MediaType = "image/png"
            }, TimeSpan.FromSeconds(10));
            using JsonDocument json = JsonDocument.Parse(File.ReadAllText(captured));
            Assert.Equal(2, json.RootElement.GetProperty("schemaVersion").GetInt32());
            if (operation == OcrOperation.DetectOrientation) {
                Assert.Equal("DetectOrientation", json.RootElement.GetProperty("operation").GetString());
                Assert.Equal(90, result.Orientation!.ClockwiseRotationDegrees);
            } else {
                // Existing strict version-2 recognition bridges still receive their original shape.
                Assert.False(json.RootElement.TryGetProperty("operation", out _));
                Assert.Equal("Recognized", result.Text);
            }
            Assert.Empty(Directory.EnumerateDirectories(directory, "officeimo-ocr-*"));
        } finally { Directory.Delete(directory, recursive: true); }
    }

    [Theory]
    [InlineData(OcrOperation.DetectOrientation, typeof(NotSupportedException))]
    [InlineData((OcrOperation)99, typeof(ArgumentOutOfRangeException))]
    public async Task UnsupportedOperationIsRejectedBeforeLaunchingProcess(OcrOperation operation, Type error) {
        var engine = new ProcessOcrEngine(new ProcessOcrEngineOptions { FileName = "must-not-launch" });
        await Assert.ThrowsAsync(error, () => engine.RecognizeAsync(new OcrRequest {
            Operation = operation, Payload = new byte[] { 1 }, MediaType = "image/png"
        }));
    }
}
