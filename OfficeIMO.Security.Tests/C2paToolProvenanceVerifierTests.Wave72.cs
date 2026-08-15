using System.IO;
using System.Runtime.InteropServices;
using OfficeIMO.Security;
using Xunit;

namespace OfficeIMO.Security.Tests;

public sealed partial class C2paToolProvenanceVerifierTestsWave72 {
    [Fact]
    public void ProcessRunnerRejectsNonUtf8ProviderOutput() {
        string executable;
        IReadOnlyList<string> arguments;
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            executable = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.System),
                "WindowsPowerShell",
                "v1.0",
                "powershell.exe");
            arguments = new[] {
                "-NoProfile",
                "-NonInteractive",
                "-Command",
                "[Console]::OpenStandardOutput().Write([byte[]](255),0,1)"
            };
        } else {
            executable = "/bin/sh";
            arguments = new[] { "-c", "printf '\\377'" };
        }
        var request = new C2paToolProcessRequest(
            executable,
            arguments,
            Path.GetTempPath(),
            TimeSpan.FromSeconds(5),
            1024);

        Assert.Throws<InvalidDataException>(() => new C2paToolProcessRunner().Run(request));
    }
}
