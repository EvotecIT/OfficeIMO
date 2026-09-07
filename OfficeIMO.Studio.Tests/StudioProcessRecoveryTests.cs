using System.Diagnostics;
using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioProcessRecoveryTests {
    [Theory]
    [InlineData("restore")]
    [InlineData("recover-missing")]
    [InlineData("private-verify")]
    public async Task AbruptTerminationPreservesSourceAndSupportsRecoveryOrPrivateRestart(string scenario) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-process-recovery-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string source = Path.Combine(root, "source.pdf");
            PdfDocument.Create(builder => builder.Page(page => page.Size(600, 800))).Save(source);
            byte[] original = File.ReadAllBytes(source);
            using (var writer = Start(scenario == "private-verify" ? "private-write" : "write", root)) {
                Task<string> errors = writer.StandardError.ReadToEndAsync();
                try {
                    string? ready = await writer.StandardOutput.ReadLineAsync().WaitAsync(TimeSpan.FromSeconds(30));
                    Assert.True(ready == "READY", ready is null ? await errors.WaitAsync(TimeSpan.FromSeconds(5)) : ready);
                    Assert.False(writer.HasExited);
                    writer.Kill(entireProcessTree: true);
                    await writer.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(10));
                } finally { await StopAsync(writer); }
                Assert.Equal(original, File.ReadAllBytes(source));
            }
            if (scenario == "recover-missing") File.Move(source, Path.Combine(root, "source-unavailable.pdf"));
            using (var reader = Start(scenario, root)) {
                Task<string> output = reader.StandardOutput.ReadToEndAsync();
                Task<string> errors = reader.StandardError.ReadToEndAsync();
                try {
                    await reader.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(30));
                    Assert.True(reader.ExitCode == 0, await errors);
                    Assert.Contains("VERIFIED", await output);
                } catch (TimeoutException) {
                    await StopAsync(reader);
                    throw new Xunit.Sdk.XunitException($"The restarted process did not exit. Output: {await output}\nErrors: {await errors}");
                } finally { await StopAsync(reader); }
            }
            Assert.Equal(original, File.ReadAllBytes(scenario == "recover-missing" ? Path.Combine(root, "source-unavailable.pdf") : source));
            Assert.Equal(2, PdfDocument.Load(Path.Combine(root, "recovered.pdf")).Read().Pages.Count);
        } finally { Directory.Delete(root, recursive: true); }
    }

    private static Process Start(string mode, string root) {
        var info = new ProcessStartInfo(Environment.GetEnvironmentVariable("DOTNET_HOST_PATH") ?? "dotnet") {
            UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true
        };
        info.ArgumentList.Add(typeof(StudioProcessProbe).Assembly.Location);
        info.ArgumentList.Add("--studio-process-probe");
        info.ArgumentList.Add(mode);
        info.ArgumentList.Add(root);
        return Process.Start(info) ?? throw new InvalidOperationException("Could not start the Studio acceptance process.");
    }

    private static async Task StopAsync(Process process) {
        if (process.HasExited) return;
        process.Kill(entireProcessTree: true);
        await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(10));
    }
}
