using OfficeIMO.Word;

namespace OfficeIMO.Tests {
    public partial class Word {
        private static void TraceOfficeSipInterop(string message) {
            string? tracePath = Environment.GetEnvironmentVariable("OFFICEIMO_VBA_SIP_TRACE_PATH");
            if (string.IsNullOrWhiteSpace(tracePath)) return;
            try {
                File.AppendAllText(
                    tracePath,
                    DateTimeOffset.UtcNow.ToString("O") + " " + message + Environment.NewLine);
            } catch (IOException) {
                // Diagnostic tracing must not change the interoperability contract.
            } catch (UnauthorizedAccessException) {
                // Diagnostic tracing must not change the interoperability contract.
            }
        }

        private sealed class TracingOfficeSipToolRunner : IWordMacroProjectToolRunner {
            private readonly IWordMacroProjectToolRunner _inner;

            internal TracingOfficeSipToolRunner(IWordMacroProjectToolRunner inner) => _inner = inner;

            public WordMacroProjectToolResult Run(
                WordMacroProjectToolInvocation invocation,
                TimeSpan timeout,
                int maxOutputCharacters) {
                string tool = Path.GetFileName(invocation.ExecutablePath);
                string operation = invocation.Arguments.Count > 0 &&
                    string.Equals(invocation.Arguments[0], "sign", StringComparison.OrdinalIgnoreCase)
                    ? "sign"
                    : "clear";
                TraceOfficeSipInterop("tool:" + tool + ":" + operation + ":start");
                WordMacroProjectToolResult result = _inner.Run(invocation, timeout, maxOutputCharacters);
                if (result.Succeeded && operation == "sign") {
                    CaptureSignedStage(invocation.Arguments[invocation.Arguments.Count - 1]);
                }
                TraceOfficeSipInterop(
                    "tool:" + tool + ":" + operation + ":complete:exit=" +
                    (result.ExitCode?.ToString() ?? "none") + ":timeout=" + result.TimedOut);
                return result;
            }

            private static void CaptureSignedStage(string stagingPath) {
                string? capturePath = Environment.GetEnvironmentVariable("OFFICEIMO_VBA_SIGNED_STAGE_CAPTURE_PATH");
                if (string.IsNullOrWhiteSpace(capturePath)) return;
                try {
                    File.Copy(stagingPath, capturePath!, overwrite: true);
                } catch (IOException) {
                    // Failure diagnostics must not change the signing contract.
                } catch (UnauthorizedAccessException) {
                    // Failure diagnostics must not change the signing contract.
                }
            }
        }

        private sealed class TracingOfficeSipPlatform : IWordMacroProjectPlatform {
            private readonly IWordMacroProjectPlatform _inner;

            internal TracingOfficeSipPlatform(IWordMacroProjectPlatform inner) => _inner = inner;

            public bool IsWindows => _inner.IsWindows;

            public bool TryGetSubjectInterfacePackage(string filePath, out Guid subjectGuid, out string detail) {
                TraceOfficeSipInterop("office-sip-discovery:start");
                bool supported = _inner.TryGetSubjectInterfacePackage(filePath, out subjectGuid, out detail);
                TraceOfficeSipInterop("office-sip-discovery:complete:" + supported);
                return supported;
            }

            public WordMacroProjectContentBindingResult ValidateContentBinding(
                string filePath,
                string digestAlgorithmOid,
                byte[] expectedDigest) {
                TraceOfficeSipInterop("office-sip-content-binding:start");
                WordMacroProjectContentBindingResult result = _inner.ValidateContentBinding(
                    filePath,
                    digestAlgorithmOid,
                    expectedDigest);
                TraceOfficeSipInterop(
                    "office-sip-content-binding:complete:supported=" + result.IsSupported +
                    ":valid=" + result.IsValid);
                return result;
            }
        }
    }
}
