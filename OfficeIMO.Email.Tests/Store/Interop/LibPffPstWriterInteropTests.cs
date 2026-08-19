using OfficeIMO.Email;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace OfficeIMO.Email.Store.Tests;

public sealed class LibPffPstWriterInteropTests {
    [LibPffInteropFact]
    public void GeneratedUnicodePstCanBeInspectedAndSemanticallyExportedByLibPff() {
        string directory = Path.Combine(Path.GetTempPath(),
            string.Concat("officeimo-libpff-interop-", Guid.NewGuid().ToString("N")));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "officeimo.pst");
        string exportTarget = Path.Combine(directory, "semantic-export");
        try {
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(path,
                new EmailStorePstWriterOptions("OfficeIMO libpff Interop"))) {
                string folder = writer.AddFolder("OfficeIMO Synthetic");
                var document = new EmailDocument {
                    Subject = "OfficeIMO synthetic libpff item",
                    MessageClass = "IPM.Note",
                    Date = new DateTimeOffset(2026, 7, 17, 0, 0, 0, TimeSpan.Zero),
                    From = new EmailAddress("sender@example.test", "Independent sender")
                };
                document.Body.Text = "OfficeIMO libpff semantic body";
                document.Recipients.Add(new EmailRecipient(EmailRecipientKind.To,
                    new EmailAddress("recipient@example.test", "Independent recipient")));
                document.MapiProperties.Add(new MapiProperty(0x8000, MapiPropertyType.Unicode,
                    "OfficeIMO named libpff evidence", name: new MapiNamedProperty(
                        MapiPropertySets.PublicStrings, "OfficeIMOInterop")));
                byte[] attachment = Encoding.UTF8.GetBytes("OfficeIMO libpff attachment evidence");
                document.Attachments.Add(new EmailAttachment {
                    FileName = "libpff-evidence.txt",
                    ContentType = "text/plain",
                    Content = attachment,
                    Length = attachment.LongLength
                });
                writer.AddItem(folder, document);
                writer.Complete();
            }

            LibPffResult info = RunLibPff("pffinfo", new[] { path });
            Assert.True(info.ExitCode == 0, info.AllOutput);
            Assert.Contains("Personal Folder File information", info.AllOutput, StringComparison.Ordinal);

            LibPffResult export = RunLibPff("pffexport", new[] {
                "-q", "-f", "all", "-d", "-t", exportTarget, path
            });
            Assert.True(export.ExitCode == 0, export.AllOutput);
            string exportedDirectory = string.Concat(exportTarget, ".export");
            Assert.True(Directory.Exists(exportedDirectory),
                string.Concat("pffexport did not create ", exportedDirectory, ". ", export.AllOutput));
            FileInfo[] exportedFiles = new DirectoryInfo(exportedDirectory)
                .EnumerateFiles("*", SearchOption.AllDirectories).ToArray();
            Assert.NotEmpty(exportedFiles);
            string[] textEvidence = exportedFiles.Where(file => file.Length <= 4 * 1024 * 1024)
                .Select(file => TryReadText(file.FullName)).Where(value => value != null).ToArray()!;
            string combined = string.Join("\n", textEvidence);
            Assert.Contains("OfficeIMO synthetic libpff item", combined, StringComparison.Ordinal);
            Assert.Contains("OfficeIMO libpff semantic body", combined, StringComparison.Ordinal);
            Assert.Contains("OfficeIMO libpff attachment evidence", combined, StringComparison.Ordinal);
        } finally {
            try { if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    private static LibPffResult RunLibPff(string tool, IReadOnlyList<string> arguments) {
        string? distribution = Environment.GetEnvironmentVariable("OFFICEIMO_EMAIL_STORE_LIBPFF_WSL");
        string executable;
        var effectiveArguments = new List<string>();
        string argumentText;
        if (!string.IsNullOrWhiteSpace(distribution)) {
            executable = "wsl.exe";
            foreach (string argument in arguments) {
                effectiveArguments.Add(Path.IsPathRooted(argument) ? ToWslPath(distribution!, argument) : argument);
            }
            argumentText = string.Concat("-d ", distribution, " -- ", tool, " ",
                string.Join(" ", effectiveArguments.Select(Quote)));
        } else {
            string info = Environment.GetEnvironmentVariable("OFFICEIMO_EMAIL_STORE_PFFINFO")!;
            string? configuredExport = Environment.GetEnvironmentVariable("OFFICEIMO_EMAIL_STORE_PFFEXPORT");
            executable = tool == "pffinfo" ? info : configuredExport ??
                Path.Combine(Path.GetDirectoryName(info) ?? string.Empty,
                    RuntimeInformation.IsOSPlatform(OSPlatform.Windows) ? "pffexport.exe" : "pffexport");
            effectiveArguments.AddRange(arguments);
            argumentText = string.Join(" ", effectiveArguments.Select(Quote));
        }

        var start = new ProcessStartInfo {
            FileName = executable,
            Arguments = argumentText,
            CreateNoWindow = true,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        using Process process = Process.Start(start)!;
        Task<string> outputTask = process.StandardOutput.ReadToEndAsync();
        Task<string> errorTask = process.StandardError.ReadToEndAsync();
        bool completed = process.WaitForExit(60_000);
        if (!completed) {
            try { process.Kill(); }
            catch (InvalidOperationException) { }
            process.WaitForExit();
        }
        string output = outputTask.GetAwaiter().GetResult();
        string error = errorTask.GetAwaiter().GetResult();
        Assert.True(completed, string.Concat(tool, " did not finish within 60 seconds."));
        return new LibPffResult(process.ExitCode, output, error);
    }

    private static string ToWslPath(string distribution, string path) {
        _ = distribution;
        string fullPath = Path.GetFullPath(path);
        if (fullPath.Length >= 3 && fullPath[1] == ':' &&
            (fullPath[2] == '\\' || fullPath[2] == '/')) {
            return string.Concat("/mnt/", char.ToLowerInvariant(fullPath[0]),
                fullPath.Substring(2).Replace('\\', '/'));
        }
        throw new NotSupportedException(string.Concat(
            "WSL libpff interoperability requires a local drive path: ", fullPath));
    }

    private static string? TryReadText(string path) {
        try { return File.ReadAllText(path); }
        catch (DecoderFallbackException) { return null; }
        catch (IOException) { return null; }
    }

    private static string Quote(string value) => string.Concat("\"", value.Replace("\"", "\\\""), "\"");

    private sealed class LibPffResult {
        internal LibPffResult(int exitCode, string output, string error) {
            ExitCode = exitCode;
            AllOutput = string.Concat(output, Environment.NewLine, error);
        }
        internal int ExitCode { get; }
        internal string AllOutput { get; }
    }
}
