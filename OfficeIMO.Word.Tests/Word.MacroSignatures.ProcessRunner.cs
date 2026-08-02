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
