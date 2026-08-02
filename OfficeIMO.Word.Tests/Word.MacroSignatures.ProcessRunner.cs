using System.Diagnostics;
using System.Runtime.InteropServices;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void MacroToolRunnerTimeoutDoesNotWaitForDescendantPipeHandles() {
            WordMacroProjectToolInvocation invocation = CreateLongRunningChildProcess();
            var runner = new WordMacroProjectProcessRunner();
            var stopwatch = Stopwatch.StartNew();

            WordMacroProjectToolResult result = runner.Run(
                invocation,
                TimeSpan.FromMilliseconds(200),
                maxOutputCharacters: 4096);

            stopwatch.Stop();
            Assert.True(result.TimedOut);
            Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(5),
                "Timed-out process cleanup took " + stopwatch.Elapsed + ".");
        }

        [Fact]
        public void MacroToolRunnerTimeoutTerminatesWindowsDescendantProcess() {
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;
            string childProcessIdPath = Path.Combine(
                _directoryWithFiles,
                "macro-tool-child-" + Guid.NewGuid().ToString("N") + ".txt");
            string command =
                "$child = Start-Process -FilePath $env:COMSPEC " +
                "-ArgumentList '/d','/s','/c','ping -n 30 127.0.0.1' -PassThru; " +
                "Set-Content -LiteralPath '" + childProcessIdPath.Replace("'", "''") + "' -Value $child.Id; " +
                "Wait-Process -Id $child.Id";
            var invocation = new WordMacroProjectToolInvocation(
                "powershell.exe",
                new[] { "-NoLogo", "-NoProfile", "-NonInteractive", "-Command", command });
            var runner = new WordMacroProjectProcessRunner();

            WordMacroProjectToolResult result = runner.Run(
                invocation,
                TimeSpan.FromSeconds(2),
                maxOutputCharacters: 4096);

            Assert.True(result.TimedOut);
            Assert.True(File.Exists(childProcessIdPath), result.Output);
            int childProcessId = int.Parse(File.ReadAllText(childProcessIdPath).Trim(),
                System.Globalization.CultureInfo.InvariantCulture);
            try {
                using Process child = Process.GetProcessById(childProcessId);
                Assert.True(child.WaitForExit(1000),
                    "The timed-out signing descendant " + childProcessId + " was still running.");
            } catch (ArgumentException) {
                // The job object already removed the descendant from the process table.
            }
        }

        [Fact]
        public void MacroToolRunnerContainsWindowsDescendantBeforeWrapperCanExit() {
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;
            string childProcessIdPath = Path.Combine(
                _directoryWithFiles,
                "macro-tool-immediate-child-" + Guid.NewGuid().ToString("N") + ".txt");
            string wrapperCommand =
                "$child = Start-Process -FilePath powershell.exe " +
                "-ArgumentList '-NoLogo','-NoProfile','-NonInteractive','-Command','Start-Sleep -Seconds 30' -PassThru; " +
                "Set-Content -LiteralPath '" + childProcessIdPath.Replace("'", "''") + "' -Value $child.Id; exit 0";
            var invocation = new WordMacroProjectToolInvocation(
                "powershell.exe",
                new[] { "-NoLogo", "-NoProfile", "-NonInteractive", "-Command", wrapperCommand });
            var runner = new WordMacroProjectProcessRunner();

            WordMacroProjectToolResult result = runner.Run(
                invocation,
                TimeSpan.FromSeconds(5),
                maxOutputCharacters: 4096);

            Assert.True(result.Succeeded, result.Output);
            Assert.True(File.Exists(childProcessIdPath), result.Output);
            int childProcessId = int.Parse(File.ReadAllText(childProcessIdPath).Trim(),
                System.Globalization.CultureInfo.InvariantCulture);
            try {
                using Process child = Process.GetProcessById(childProcessId);
                Assert.True(child.WaitForExit(1000),
                    "The immediately spawned signing descendant " + childProcessId + " was still running.");
            } catch (ArgumentException) {
                // Closing the launch-time Job Object already removed the descendant.
            }
        }

        private static WordMacroProjectToolInvocation CreateLongRunningChildProcess() {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
                return new WordMacroProjectToolInvocation(
                    Environment.GetEnvironmentVariable("COMSPEC") ?? "cmd.exe",
                    new[] { "/d", "/s", "/c", "ping -n 30 127.0.0.1" });
            }
            return new WordMacroProjectToolInvocation(
                "/bin/sh",
                new[] { "-c", "sleep 30 & wait" });
        }
    }
}
