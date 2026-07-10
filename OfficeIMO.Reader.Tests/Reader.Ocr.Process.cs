using OfficeIMO.Reader;
using OfficeIMO.Reader.Ocr.Process;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderOcrProcessTests {
    [Fact]
    public void ProcessOfficeOcrProtocol_RejectsIncompatibleResponseVersion() {
        const string json = "{\"schemaId\":\"officeimo.reader.ocr.process-response\",\"schemaVersion\":2,\"result\":{\"text\":\"late\"}}";

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => ProcessOfficeOcrProtocol.DeserializeResult(json));

        Assert.Contains("version", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task OfficeOcrProcessRunner_DrainsAndBoundsStandardOutput() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-ocr-runner-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            bool windows = RuntimeInformation.IsOSPlatform(OSPlatform.Windows);
            string scriptPath = Path.Combine(directory, windows ? "output.cmd" : "output.sh");
            File.WriteAllText(scriptPath, windows ? "@echo 1234567890\r\n" : "printf 1234567890\n");
            OfficeOcrProcessResult result = await OfficeOcrProcessRunner.RunAsync(new OfficeOcrProcessCommand {
                FileName = windows ? Environment.GetEnvironmentVariable("ComSpec") ?? "cmd.exe" : "/bin/sh",
                Arguments = windows ? new[] { "/d", "/c", scriptPath } : new[] { scriptPath },
                MaxStandardOutputCharacters = 5,
                MaxStandardErrorCharacters = 5
            });

            Assert.Equal(0, result.ExitCode);
            Assert.Equal("12345", result.StandardOutput);
            Assert.True(result.StandardOutputTruncated);
        } finally {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task ProcessOfficeOcrEngine_RoundTripsVersionedJsonProtocolWithoutShellExpansion() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-ocr-process-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string responsePath = Path.Combine(directory, "fixture result.json");
            File.WriteAllText(responsePath, ProcessOfficeOcrProtocol.SerializeResult(new OfficeOcrEngineResult {
                Text = "Invoice 1042",
                Confidence = 0.97,
                Language = "eng",
                Provider = "fixture-process",
                Spans = new[] {
                    new OfficeOcrTextSpan { Sequence = 0, Level = OfficeOcrTextSpanLevel.Character, Text = "I", Confidence = 0.99 }
                }
            }));
            bool windows = RuntimeInformation.IsOSPlatform(OSPlatform.Windows);
            string scriptPath = Path.Combine(directory, windows ? "copy response.cmd" : "copy response.sh");
            File.WriteAllText(scriptPath, windows ? "@copy /Y \"%~1\" \"%~2\" >nul\r\n" : "cp \"$1\" \"$2\"\n");
            var arguments = windows
                ? new[] { "/d", "/c", scriptPath, responsePath, "{output}" }
                : new[] { scriptPath, responsePath, "{output}" };
            var engine = new ProcessOfficeOcrEngine(new ProcessOfficeOcrEngineOptions {
                FileName = windows ? Environment.GetEnvironmentVariable("ComSpec") ?? "cmd.exe" : "/bin/sh",
                Arguments = arguments,
                Id = "fixture-process",
                TemporaryDirectory = directory
            });
            byte[] payload = new byte[] { 1, 2, 3 };

            OfficeOcrEngineResult result = await engine.RecognizeAsync(new OfficeOcrEngineRequest {
                Candidate = new OfficeDocumentOcrCandidate { Id = "ocr-1", Kind = "image", AssetId = "asset-1" },
                Asset = new OfficeDocumentAsset { Id = "asset-1", Kind = "image", MediaType = "image/png", Extension = ".png" },
                Payload = payload,
                Language = "eng",
                Source = new OfficeDocumentSource { Path = "scan.pdf" }
            });

            Assert.Equal("Invoice 1042", result.Text);
            Assert.Equal("fixture-process", result.Provider);
            Assert.Equal(OfficeOcrTextSpanLevel.Character, Assert.Single(result.Spans).Level);
            Assert.Empty(Directory.EnumerateDirectories(directory, "officeimo-ocr-*"));
        } finally {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }
}
