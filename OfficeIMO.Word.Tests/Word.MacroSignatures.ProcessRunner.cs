using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
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
            string wrapperStartedPath = Path.Combine(
                _directoryWithFiles,
                "macro-tool-wrapper-started-" + Guid.NewGuid().ToString("N") + ".txt");
            string childStartedPath = Path.Combine(
                _directoryWithFiles,
                "macro-tool-child-started-" + Guid.NewGuid().ToString("N") + ".txt");
            string releaseChildPath = Path.Combine(
                _directoryWithFiles,
                "macro-tool-release-child-" + Guid.NewGuid().ToString("N") + ".txt");
            string childSurvivedPath = Path.Combine(
                _directoryWithFiles,
                "macro-tool-child-survived-" + Guid.NewGuid().ToString("N") + ".txt");
            string childScriptPath = Path.Combine(
                _directoryWithFiles,
                "macro-tool-child-" + Guid.NewGuid().ToString("N") + ".cmd");
            string wrapperScriptPath = Path.Combine(
                _directoryWithFiles,
                "macro-tool-wrapper-" + Guid.NewGuid().ToString("N") + ".cmd");
            string commandInterpreter = Environment.GetEnvironmentVariable("COMSPEC") ?? "cmd.exe";
            File.WriteAllText(childScriptPath,
                "@echo off\r\n" +
                "> \"" + childStartedPath + "\" echo started\r\n" +
                ":wait_for_release\r\n" +
                "if not exist \"" + releaseChildPath + "\" (\r\n" +
                "  ping -n 2 127.0.0.1 >nul\r\n" +
                "  goto wait_for_release\r\n" +
                ")\r\n" +
                "> \"" + childSurvivedPath + "\" echo survived\r\n");
            File.WriteAllText(wrapperScriptPath,
                "@echo off\r\n" +
                "start \"\" /b \"" + commandInterpreter + "\" /d /s /c call \"" + childScriptPath + "\"\r\n" +
                "> \"" + wrapperStartedPath + "\" echo started\r\n" +
                "ping -n 30 127.0.0.1 >nul\r\n");
            var invocation = new WordMacroProjectToolInvocation(
                commandInterpreter,
                new[] { "/d", "/s", "/c", "call", wrapperScriptPath });
            var runner = new WordMacroProjectProcessRunner();

            WordMacroProjectToolResult result = runner.Run(
                invocation,
                TimeSpan.FromSeconds(8),
                maxOutputCharacters: 4096);

            string resultDetail = "ExitCode=" + result.ExitCode + ", TimedOut=" + result.TimedOut +
                                  ", Output=" + result.Output;
            Assert.True(result.TimedOut, resultDetail);
            Assert.True(File.Exists(wrapperStartedPath), resultDetail);
            Assert.True(File.Exists(childStartedPath), resultDetail);

            File.WriteAllText(releaseChildPath, "release");
            Thread.Sleep(TimeSpan.FromSeconds(3));
            Assert.False(File.Exists(childSurvivedPath),
                "A signing descendant survived after its parent timed out and the Job Object closed.");
        }

        [Fact]
        public void MacroToolRunnerContainsWindowsDescendantBeforeWrapperCanExit() {
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;
            string wrapperStartedPath = Path.Combine(
                _directoryWithFiles,
                "macro-tool-wrapper-started-" + Guid.NewGuid().ToString("N") + ".txt");
            string childSurvivedPath = Path.Combine(
                _directoryWithFiles,
                "macro-tool-child-survived-" + Guid.NewGuid().ToString("N") + ".txt");
            string childScriptPath = Path.Combine(
                _directoryWithFiles,
                "macro-tool-child-" + Guid.NewGuid().ToString("N") + ".cmd");
            string wrapperScriptPath = Path.Combine(
                _directoryWithFiles,
                "macro-tool-wrapper-" + Guid.NewGuid().ToString("N") + ".cmd");
            string commandInterpreter = Environment.GetEnvironmentVariable("COMSPEC") ?? "cmd.exe";
            File.WriteAllText(childScriptPath,
                "@echo off\r\n" +
                "ping -n 3 127.0.0.1 >nul\r\n" +
                "> \"" + childSurvivedPath + "\" echo survived\r\n");
            File.WriteAllText(wrapperScriptPath,
                "@echo off\r\n" +
                "start \"\" /b \"" + commandInterpreter + "\" /d /s /c call \"" + childScriptPath + "\"\r\n" +
                "> \"" + wrapperStartedPath + "\" echo started\r\n" +
                "exit /b 0\r\n");
            var invocation = new WordMacroProjectToolInvocation(
                commandInterpreter,
                new[] { "/d", "/s", "/c", "call", wrapperScriptPath });
            var runner = new WordMacroProjectProcessRunner();

            WordMacroProjectToolResult result = runner.Run(
                invocation,
                TimeSpan.FromSeconds(8),
                maxOutputCharacters: 4096);

            string resultDetail = "ExitCode=" + result.ExitCode + ", TimedOut=" + result.TimedOut +
                                  ", Output=" + result.Output;
            Assert.False(result.TimedOut, resultDetail);
            Assert.True(File.Exists(wrapperStartedPath), resultDetail);
            Thread.Sleep(TimeSpan.FromSeconds(3));
            Assert.False(File.Exists(childSurvivedPath),
                "A signing descendant survived after its wrapper exited and the launch-time Job Object closed.");
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
