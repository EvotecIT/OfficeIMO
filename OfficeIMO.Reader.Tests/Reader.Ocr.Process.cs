using OfficeIMO.Reader;
using OfficeIMO.Reader.Ocr.Process;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderOcrProcessTests {
    [Fact]
    public void OfficeOcrTemporaryStorage_CreatesOwnerOnlyUnixDirectoryAndFile() {
#if NET8_0_OR_GREATER
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;
        string root = Path.Combine(Path.GetTempPath(), "officeimo-private-ocr-test-" + Guid.NewGuid().ToString("N"));
        try {
            string requestDirectory = OfficeOcrTemporaryStorage.CreateRequestDirectory(root, "request-");
            string payloadPath = Path.Combine(requestDirectory, "payload.bin");
            string outputPath = Path.Combine(requestDirectory, "result.json");
            OfficeOcrTemporaryStorage.WriteAllBytes(payloadPath, new byte[] { 1, 2, 3 });
            File.WriteAllText(outputPath, "{}");
            OfficeOcrTemporaryStorage.EnsurePrivateFile(outputPath);

            const UnixFileMode allPermissions = UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute
                | UnixFileMode.GroupRead | UnixFileMode.GroupWrite | UnixFileMode.GroupExecute
                | UnixFileMode.OtherRead | UnixFileMode.OtherWrite | UnixFileMode.OtherExecute;
            Assert.Equal(
                UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute,
                File.GetUnixFileMode(requestDirectory) & allPermissions);
            Assert.Equal(
                UnixFileMode.UserRead | UnixFileMode.UserWrite,
                File.GetUnixFileMode(payloadPath) & allPermissions);
            Assert.Equal(
                UnixFileMode.UserRead | UnixFileMode.UserWrite,
                File.GetUnixFileMode(outputPath) & allPermissions);
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
#endif
    }

    [Fact]
    public void ProcessOfficeOcrProtocol_RejectsIncompatibleResponseVersion() {
        const string json = "{\"schemaId\":\"officeimo.reader.ocr.process-response\",\"schemaVersion\":2,\"result\":{\"text\":\"late\"}}";

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => ProcessOfficeOcrProtocol.DeserializeResult(json));

        Assert.Contains("version", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("{\"schemaVersion\":1,\"result\":{\"text\":\"unversioned\"}}", "schemaId")]
    [InlineData("{\"schemaId\":\"officeimo.reader.ocr.process-response\",\"result\":{\"text\":\"unversioned\"}}", "schemaVersion")]
    public void ProcessOfficeOcrProtocol_RejectsMissingResponseSchemaFields(string json, string field) {
        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => ProcessOfficeOcrProtocol.DeserializeResult(json));

        Assert.Contains(field, exception.Message, StringComparison.OrdinalIgnoreCase);
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
    public async Task OfficeOcrProcessRunner_TerminatesWrapperDescendantsAfterTimeout() {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-ocr-runner-pipe-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        int? childProcessId = null;
        try {
            string scriptPath = Path.Combine(directory, "inherited-pipe.sh");
            string childProcessPath = Path.Combine(directory, "child.pid");
            File.WriteAllText(scriptPath, "(trap '' HUP; sleep 30) &\necho $! > \"$1\"\nexit 0\n");
            var stopwatch = Stopwatch.StartNew();

            await Assert.ThrowsAsync<TimeoutException>(() => OfficeOcrProcessRunner.RunAsync(new OfficeOcrProcessCommand {
                FileName = "/bin/sh",
                Arguments = new[] { scriptPath, childProcessPath },
                Timeout = TimeSpan.FromMilliseconds(100)
            }));

            Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(1.5), "The process runner waited for inherited pipe handles after its timeout.");
            Assert.True(File.Exists(childProcessPath), "The wrapper did not record its child process id.");
            childProcessId = int.Parse(File.ReadAllText(childProcessPath), CultureInfo.InvariantCulture);
            Assert.True(
                await WaitForProcessExitAsync(childProcessId.Value, TimeSpan.FromSeconds(2)),
                "The process runner left a wrapper child alive after its timeout.");
        } finally {
            if (childProcessId.HasValue) TryKillProcess(childProcessId.Value);
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
            string scriptPath = Path.Combine(directory, windows ? "copy-response.cmd" : "copy-response.sh");
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

    private static async Task<bool> WaitForProcessExitAsync(int processId, TimeSpan timeout) {
        var stopwatch = Stopwatch.StartNew();
        while (stopwatch.Elapsed < timeout) {
            try {
                using System.Diagnostics.Process process = System.Diagnostics.Process.GetProcessById(processId);
                if (process.HasExited) return true;
            } catch (ArgumentException) {
                return true;
            }
            await Task.Delay(20);
        }
        return false;
    }

    private static void TryKillProcess(int processId) {
        try {
            using System.Diagnostics.Process process = System.Diagnostics.Process.GetProcessById(processId);
            if (!process.HasExited) process.Kill();
        } catch (ArgumentException) {
        } catch (InvalidOperationException) {
        }
    }
}
