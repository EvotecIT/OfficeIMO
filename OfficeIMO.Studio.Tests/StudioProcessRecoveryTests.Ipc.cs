namespace OfficeIMO.Studio.Tests;

public sealed partial class StudioProcessRecoveryTests {
    [Fact]
    public async Task OverlongUnixPipePathAllowsAnIndependentLaunch() {
        if (OperatingSystem.IsWindows()) return;
        string root = Path.Combine(Path.GetTempPath(), "studio-ipc-" + Guid.NewGuid().ToString("N"));
        string temp = Path.Combine(root, new string('t', 110));
        Directory.CreateDirectory(temp);
        try {
            using var process = Start("ipc-long-temp", root, temp);
            Task<string> output = process.StandardOutput.ReadToEndAsync();
            Task<string> errors = process.StandardError.ReadToEndAsync();
            try {
                await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(15));
                Assert.True(process.ExitCode == 0, await errors);
                Assert.Contains("VERIFIED independent launch", await output);
            } finally { await StopAsync(process); }
        } finally { Directory.Delete(root, recursive: true); }
    }
}
