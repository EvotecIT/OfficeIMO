using OfficeIMO.Workflows;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class PdfArchiveCommandTests {
    [Fact]
    public async Task ArchiveCommandExecutesAndResumesTheSharedContract() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-archive-cli-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Path.Combine(root, "input"));
        try {
            var request = new OfficePdfArchiveRequest {
                InputDirectory = Path.Combine(root, "input"), OutputDirectory = Path.Combine(root, "output"),
                CheckpointDirectory = Path.Combine(root, "state")
            };
            string requestPath = Path.Combine(root, "request.json");
            File.WriteAllText(requestPath, System.Text.Json.JsonSerializer.Serialize(new {
                request.InputDirectory, request.OutputDirectory, request.CheckpointDirectory
            }));
            File.WriteAllText(Path.Combine(request.InputDirectory, "source.txt"), "Literal <b>content</b>");
            for (int run = 0; run < 2; run++) {
                using var output = new MemoryStream();
                using var error = new StringWriter();
                int code = await OfficeImoToolApp.RunAsync(["workflow", "archive", "--request", requestPath], Stream.Null, output, error);
                Assert.Equal((int)OfficeImoToolExitCode.Success, code);
                string json = Encoding.UTF8.GetString(output.ToArray());
                Assert.Contains("\"Completed\":1", json);
                Assert.Contains("\"Reused\":" + run, json);
                Assert.Equal(string.Empty, error.ToString());
            }
            Assert.True(File.Exists(Path.Combine(request.OutputDirectory, "source.txt.pdf")));
        } finally { Directory.Delete(root, true); }
    }
}
