using System.Diagnostics;
using System.Runtime.InteropServices;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public async Task MacroToolRunnerReadsNewlineFreeOutputInBoundedChunks() {
            using var reader = new ChunkOnlyTextReader(4096);
            var output = new WordMacroProjectProcessRunner.BoundedProcessOutput(64);

            await WordMacroProjectProcessRunner.ReadRedirectedOutput(reader, output);

            string value = output.ToString();
            Assert.True(reader.ReadCount > 1);
            Assert.StartsWith(new string('x', 64), value, StringComparison.Ordinal);
            Assert.EndsWith("[output truncated]", value, StringComparison.Ordinal);
        }

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
            string scriptPath = Path.Combine(
                _directoryWithFiles,
                "macro-tool-child-" + Guid.NewGuid().ToString("N") + ".ps1");
            File.WriteAllText(scriptPath,
                "$child = Start-Process -FilePath $env:COMSPEC " +
                "-ArgumentList '/d','/s','/c','ping -n 30 127.0.0.1' -PassThru\r\n" +
                "[System.IO.File]::WriteAllText('" + childProcessIdPath.Replace("'", "''") + "', [string]$child.Id)\r\n" +
                "Wait-Process -Id $child.Id\r\n");
            var invocation = new WordMacroProjectToolInvocation(
                "powershell.exe",
                new[] { "-NoLogo", "-NoProfile", "-NonInteractive", "-File", scriptPath });
            var runner = new WordMacroProjectProcessRunner();

            WordMacroProjectToolResult result = runner.Run(
                invocation,
                TimeSpan.FromSeconds(8),
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
            string scriptPath = Path.Combine(
                _directoryWithFiles,
                "macro-tool-immediate-child-" + Guid.NewGuid().ToString("N") + ".ps1");
            File.WriteAllText(scriptPath,
                "$child = Start-Process -FilePath $env:COMSPEC " +
                "-ArgumentList '/d','/s','/c','ping -n 30 127.0.0.1' -PassThru\r\n" +
                "[System.IO.File]::WriteAllText('" + childProcessIdPath.Replace("'", "''") + "', [string]$child.Id)\r\n" +
                "exit 0\r\n");
            var invocation = new WordMacroProjectToolInvocation(
                "powershell.exe",
                new[] { "-NoLogo", "-NoProfile", "-NonInteractive", "-File", scriptPath });
            var runner = new WordMacroProjectProcessRunner();

            WordMacroProjectToolResult result = runner.Run(
                invocation,
                TimeSpan.FromSeconds(15),
                maxOutputCharacters: 4096);

            string resultDetail = "ExitCode=" + result.ExitCode + ", TimedOut=" + result.TimedOut +
                                  ", Output=" + result.Output;
            Assert.False(result.TimedOut, resultDetail);
            Assert.True(File.Exists(childProcessIdPath), resultDetail);
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

        private sealed class ChunkOnlyTextReader : TextReader {
            private int _remaining;

            internal ChunkOnlyTextReader(int characters) => _remaining = characters;

            internal int ReadCount { get; private set; }

            public override string? ReadLine() =>
                throw new InvalidOperationException("The bounded reader must not materialize complete lines.");

            public override int Read(char[] buffer, int index, int count) {
                if (_remaining == 0) return 0;
                int read = Math.Min(count, _remaining);
                for (int offset = 0; offset < read; offset++) buffer[index + offset] = 'x';
                _remaining -= read;
                ReadCount++;
                return read;
            }
        }
    }
}
