using OfficeIMO.Email;
using System.Diagnostics;
using System.Threading.Tasks;

namespace OfficeIMO.Email.Store.Tests;

public sealed class LibPffPstWriterInteropTests {
    [LibPffInteropFact]
    public void GeneratedUnicodePstCanBeInspectedByLibPff() {
        string executable = Environment.GetEnvironmentVariable("OFFICEIMO_EMAIL_STORE_PFFINFO")!;
        string path = Path.Combine(Path.GetTempPath(),
            string.Concat("officeimo-libpff-interop-", Guid.NewGuid().ToString("N"), ".pst"));
        try {
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(path,
                new EmailStorePstWriterOptions("OfficeIMO libpff Interop"))) {
                string folder = writer.AddFolder("OfficeIMO Synthetic");
                writer.AddItem(folder, new EmailDocument {
                    Subject = "OfficeIMO synthetic libpff item",
                    MessageClass = "IPM.Note",
                    Date = new DateTimeOffset(2026, 7, 17, 0, 0, 0, TimeSpan.Zero)
                });
                writer.Complete();
            }

            var start = new ProcessStartInfo {
                FileName = executable,
                Arguments = Quote(path),
                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            using Process process = Process.Start(start)!;
            Task<string> outputTask = process.StandardOutput.ReadToEndAsync();
            Task<string> errorTask = process.StandardError.ReadToEndAsync();
            bool completed = process.WaitForExit(30_000);
            if (!completed) {
                try { process.Kill(); }
                catch (InvalidOperationException) { }
                process.WaitForExit();
            }
            string output = outputTask.GetAwaiter().GetResult();
            string error = errorTask.GetAwaiter().GetResult();
            Assert.True(completed, "pffinfo did not finish within 30 seconds.");
            Assert.Equal(0, process.ExitCode);
            Assert.DoesNotContain("error", string.Concat(output, Environment.NewLine, error),
                StringComparison.OrdinalIgnoreCase);
        } finally {
            try { if (File.Exists(path)) File.Delete(path); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    private static string Quote(string value) => string.Concat("\"", value.Replace("\"", "\\\""), "\"");
}
