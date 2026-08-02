using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace OfficeIMO.Word {
    internal sealed class WordMacroProjectToolInvocation {
        internal WordMacroProjectToolInvocation(string executablePath, IReadOnlyList<string> arguments) {
            ExecutablePath = executablePath;
            Arguments = arguments;
        }

        internal string ExecutablePath { get; }
        internal IReadOnlyList<string> Arguments { get; }
    }

    internal sealed class WordMacroProjectToolResult {
        internal WordMacroProjectToolResult(int? exitCode, bool timedOut, string output) {
            ExitCode = exitCode;
            TimedOut = timedOut;
            Output = output;
        }

        internal int? ExitCode { get; }
        internal bool TimedOut { get; }
        internal string Output { get; }
        internal bool Succeeded => !TimedOut && ExitCode == 0;
    }

    internal interface IWordMacroProjectToolRunner {
        WordMacroProjectToolResult Run(WordMacroProjectToolInvocation invocation, TimeSpan timeout,
            int maxOutputCharacters);
    }

    internal interface IWordMacroProjectPlatform {
        bool IsWindows { get; }
        bool TryGetSubjectInterfacePackage(string filePath, out Guid subjectGuid, out string detail);
        WordMacroProjectContentBindingResult ValidateContentBinding(
            string filePath,
            string digestAlgorithmOid,
            byte[] expectedDigest);
    }

    internal sealed class WordMacroProjectContentBindingResult {
        internal WordMacroProjectContentBindingResult(bool isSupported, bool isValid, string detail) {
            IsSupported = isSupported;
            IsValid = isValid;
            Detail = detail;
        }

        internal bool IsSupported { get; }
        internal bool IsValid { get; }
        internal string Detail { get; }
    }

    internal sealed class WordMacroProjectSigningDependencies {
        internal WordMacroProjectSigningDependencies(
            IWordMacroProjectToolRunner runner,
            IWordMacroProjectPlatform platform) {
            Runner = runner;
            Platform = platform;
        }

        internal IWordMacroProjectToolRunner Runner { get; }
        internal IWordMacroProjectPlatform Platform { get; }

        internal static WordMacroProjectSigningDependencies Default { get; } =
            new WordMacroProjectSigningDependencies(
                new WordMacroProjectProcessRunner(),
                new WordMacroProjectWindowsPlatform());
    }

    internal sealed class WordMacroProjectProcessRunner : IWordMacroProjectToolRunner {
        private const int ExitGraceMilliseconds = 1000;

        public WordMacroProjectToolResult Run(WordMacroProjectToolInvocation invocation, TimeSpan timeout,
            int maxOutputCharacters) {
            if (invocation == null) throw new ArgumentNullException(nameof(invocation));
            if (timeout <= TimeSpan.Zero || timeout.TotalMilliseconds > int.MaxValue) {
                throw new ArgumentOutOfRangeException(nameof(timeout));
            }
            if (maxOutputCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(maxOutputCharacters));

            var startInfo = new ProcessStartInfo {
                FileName = invocation.ExecutablePath,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };
#if NET6_0_OR_GREATER
            foreach (string argument in invocation.Arguments) startInfo.ArgumentList.Add(argument);
#else
            startInfo.Arguments = string.Join(" ", invocation.Arguments.Select(QuoteWindowsArgument));
#endif

            using var process = new Process { StartInfo = startInfo };
            var output = new BoundedProcessOutput(maxOutputCharacters);
            var standardOutputCompleted = new TaskCompletionSource<bool>();
            var standardErrorCompleted = new TaskCompletionSource<bool>();
            process.OutputDataReceived += (_, args) => {
                if (args.Data == null) standardOutputCompleted.TrySetResult(true);
                else output.Append(args.Data);
            };
            process.ErrorDataReceived += (_, args) => {
                if (args.Data == null) standardErrorCompleted.TrySetResult(true);
                else output.Append(args.Data);
            };
            WordMacroProjectProcessTree? processTree = null;
            try {
                if (!process.Start()) {
                    return new WordMacroProjectToolResult(null, false,
                        "The external signing tool process did not start.");
                }
                if (!WordMacroProjectProcessTree.TryAttach(process, out processTree, out string containmentDetail)) {
                    TryTerminateProcess(process, processTree);
                    return new WordMacroProjectToolResult(null, false,
                        "The external signing tool process tree could not be contained. " + containmentDetail);
                }
                process.BeginOutputReadLine();
                process.BeginErrorReadLine();
                bool exited = process.WaitForExit(checked((int)timeout.TotalMilliseconds));
                if (!exited) {
                    bool treeTerminated = processTree?.TerminateAndWait(
                        TimeSpan.FromMilliseconds(ExitGraceMilliseconds)) ?? false;
                    TryTerminateProcess(process, processTree, treeTerminated);
                    process.WaitForExit(ExitGraceMilliseconds);
                    if (processTree != null && !treeTerminated) {
                        output.Append("The contained signing process tree did not confirm termination within the bounded grace period.");
                    }
                    DrainRedirectedOutput(process, standardOutputCompleted.Task, standardErrorCompleted.Task, output);
                    return new WordMacroProjectToolResult(null, true, output.ToString());
                }
                DrainRedirectedOutput(process, standardOutputCompleted.Task, standardErrorCompleted.Task, output);
                return new WordMacroProjectToolResult(process.ExitCode, false, output.ToString());
            } catch (Exception exception) when (exception is Win32Exception || exception is IOException ||
                exception is InvalidOperationException || exception is UnauthorizedAccessException ||
                exception is PlatformNotSupportedException || exception is NotSupportedException) {
                return new WordMacroProjectToolResult(null, false, exception.Message);
            } finally {
                processTree?.Dispose();
            }
        }

        private static void TryTerminateProcess(
            Process process,
            WordMacroProjectProcessTree? processTree,
            bool processTreeAlreadyTerminated = false) {
            if (processTree != null && !processTreeAlreadyTerminated) {
                processTree.TerminateAndWait(TimeSpan.FromMilliseconds(ExitGraceMilliseconds));
            }
            try {
#if NET6_0_OR_GREATER
                if (!process.HasExited) process.Kill(entireProcessTree: true);
#else
                if (!process.HasExited) process.Kill();
#endif
            } catch (Exception exception) when (exception is InvalidOperationException ||
                exception is Win32Exception || exception is NotSupportedException) {
            }
        }

        private static void DrainRedirectedOutput(
            Process process,
            Task standardOutputCompleted,
            Task standardErrorCompleted,
            BoundedProcessOutput output) {
            if (Task.WaitAll(new[] { standardOutputCompleted, standardErrorCompleted }, ExitGraceMilliseconds)) {
                return;
            }
            try { process.CancelOutputRead(); } catch (InvalidOperationException) { }
            try { process.CancelErrorRead(); } catch (InvalidOperationException) { }
            output.Append("Redirected process output did not close within the bounded drain period.");
        }

#if !NET6_0_OR_GREATER
        private static string QuoteWindowsArgument(string argument) {
            if (argument.Length > 0 && argument.All(character =>
                    !char.IsWhiteSpace(character) && character != '"')) {
                return argument;
            }

            var result = new StringBuilder(argument.Length + 2);
            result.Append('"');
            int backslashes = 0;
            foreach (char character in argument) {
                if (character == '\\') {
                    backslashes++;
                    continue;
                }
                if (character == '"') {
                    result.Append('\\', backslashes * 2 + 1);
                    result.Append('"');
                    backslashes = 0;
                    continue;
                }
                result.Append('\\', backslashes);
                backslashes = 0;
                result.Append(character);
            }
            result.Append('\\', backslashes * 2);
            result.Append('"');
            return result.ToString();
        }
#endif

        private sealed class BoundedProcessOutput {
            private readonly int _maximum;
            private readonly StringBuilder _builder = new StringBuilder();
            private readonly object _sync = new object();
            private bool _truncated;

            internal BoundedProcessOutput(int maximum) => _maximum = maximum;

            internal void Append(string? line) {
                if (line == null) return;
                lock (_sync) {
                    if (_builder.Length >= _maximum) {
                        _truncated = true;
                        return;
                    }
                    int remaining = _maximum - _builder.Length;
                    string value = line.Length > remaining ? line.Substring(0, remaining) : line;
                    _builder.Append(value);
                    if (_builder.Length < _maximum) _builder.AppendLine();
                    if (value.Length != line.Length) _truncated = true;
                }
            }

            public override string ToString() {
                lock (_sync) {
                    string value = _builder.ToString().Trim();
                    return _truncated ? value + Environment.NewLine + "[output truncated]" : value;
                }
            }
        }
    }

    internal sealed class WordMacroProjectWindowsPlatform : IWordMacroProjectPlatform {
        private static readonly Guid MicrosoftOfficeOpenXmlSip =
            new Guid("6E64D5BD-CEB0-4B66-B4A0-15AC71775C48");

        public bool IsWindows => RuntimeInformation.IsOSPlatform(OSPlatform.Windows);

        public bool TryGetSubjectInterfacePackage(string filePath, out Guid subjectGuid, out string detail) {
            subjectGuid = Guid.Empty;
            if (!IsWindows) {
                detail = "Microsoft Office SIP discovery is available on Windows only.";
                return false;
            }
            try {
                if (!CryptSIPRetrieveSubjectGuid(filePath, IntPtr.Zero, out subjectGuid)) {
                    detail = "Windows could not resolve a registered Subject Interface Package for the Word file.";
                    return false;
                }
                if (subjectGuid != MicrosoftOfficeOpenXmlSip) {
                    detail = "The registered SIP subject GUID " + subjectGuid +
                        " is not Microsoft's OOXML Office SIP " + MicrosoftOfficeOpenXmlSip + ".";
                    return false;
                }
                detail = "Microsoft's OOXML Office SIP is registered for the Word file.";
                return true;
            } catch (Exception exception) when (exception is DllNotFoundException ||
                exception is EntryPointNotFoundException || exception is BadImageFormatException) {
                detail = "Microsoft Office SIP discovery failed. " + exception.Message;
                return false;
            }
        }

        public WordMacroProjectContentBindingResult ValidateContentBinding(
            string filePath,
            string digestAlgorithmOid,
            byte[] expectedDigest) {
            if (!IsWindows) {
                return new WordMacroProjectContentBindingResult(false, false,
                    "Microsoft Office SIP content validation is available on Windows only.");
            }
            if (string.IsNullOrWhiteSpace(digestAlgorithmOid) || expectedDigest == null ||
                expectedDigest.Length == 0 || expectedDigest.Length > 1024) {
                return new WordMacroProjectContentBindingResult(true, false,
                    "The selected VBA signature does not contain a bounded Authenticode subject digest.");
            }
            if (!TryGetSubjectInterfacePackage(filePath, out Guid subjectGuid, out string sipDetail)) {
                return new WordMacroProjectContentBindingResult(false, false, sipDetail);
            }

            IntPtr subjectGuidPointer = IntPtr.Zero;
            IntPtr fileNamePointer = IntPtr.Zero;
            IntPtr algorithmPointer = IntPtr.Zero;
            IntPtr indirectDataPointer = IntPtr.Zero;
            try {
                subjectGuidPointer = Marshal.AllocHGlobal(Marshal.SizeOf<Guid>());
                Marshal.StructureToPtr(subjectGuid, subjectGuidPointer, false);
                fileNamePointer = Marshal.StringToHGlobalUni(Path.GetFullPath(filePath));
                algorithmPointer = Marshal.StringToHGlobalAnsi(digestAlgorithmOid);
                var subject = new SipSubjectInfo {
                    cbSize = checked((uint)Marshal.SizeOf<SipSubjectInfo>()),
                    pgSubjectType = subjectGuidPointer,
                    hFile = new IntPtr(-1),
                    pwsFileName = fileNamePointer,
                    pwsDisplayName = fileNamePointer,
                    DigestAlgorithm = new CryptAlgorithmIdentifier { pszObjId = algorithmPointer },
                    dwEncodingType = 0x00010001,
                    dwUnionChoice = 0
                };

                uint signedDataBytes = 0;
                bool signatureSelected = CryptSIPGetSignedDataMsg(
                    ref subject, out _, 0, ref signedDataBytes, IntPtr.Zero);
                int selectionError = Marshal.GetLastWin32Error();
                const int ErrorInsufficientBuffer = 122;
                if ((!signatureSelected && selectionError != ErrorInsufficientBuffer) || signedDataBytes == 0) {
                    return new WordMacroProjectContentBindingResult(true, false,
                        "The Office SIP could not select the highest-precedence VBA signature. Windows error " +
                        selectionError + ".");
                }
                uint indirectDataBytes = 0;
                if (!CryptSIPCreateIndirectData(ref subject, ref indirectDataBytes, IntPtr.Zero) ||
                    indirectDataBytes == 0 || indirectDataBytes > 1024 * 1024) {
                    return NativeFailure("The Office SIP could not size the current VBA subject digest.");
                }
                indirectDataPointer = Marshal.AllocHGlobal(checked((int)indirectDataBytes));
                if (!CryptSIPCreateIndirectData(ref subject, ref indirectDataBytes, indirectDataPointer)) {
                    return NativeFailure("The Office SIP could not calculate the current VBA subject digest.");
                }
                SipIndirectData indirect = Marshal.PtrToStructure<SipIndirectData>(indirectDataPointer);
                if (indirect.Digest.cbData == 0 || indirect.Digest.cbData > 1024 ||
                    indirect.Digest.pbData == IntPtr.Zero) {
                    return new WordMacroProjectContentBindingResult(true, false,
                        "The Office SIP returned an invalid subject digest.");
                }
                string? actualAlgorithm = Marshal.PtrToStringAnsi(indirect.DigestAlgorithm.pszObjId);
                byte[] actualDigest = new byte[indirect.Digest.cbData];
                Marshal.Copy(indirect.Digest.pbData, actualDigest, 0, actualDigest.Length);
                bool valid = string.Equals(actualAlgorithm, digestAlgorithmOid, StringComparison.Ordinal) &&
                    FixedTimeEquals(actualDigest, expectedDigest);
                return new WordMacroProjectContentBindingResult(true, valid,
                    valid
                        ? "Microsoft's Office SIP reproduced the signed VBA subject digest."
                        : "Microsoft's Office SIP subject digest does not match the selected VBA signature.");
            } catch (Exception exception) when (exception is DllNotFoundException ||
                exception is EntryPointNotFoundException || exception is BadImageFormatException ||
                exception is ArgumentException || exception is OverflowException || exception is IOException) {
                return new WordMacroProjectContentBindingResult(false, false,
                    "Microsoft Office SIP content validation failed. " + exception.Message);
            } finally {
                if (indirectDataPointer != IntPtr.Zero) Marshal.FreeHGlobal(indirectDataPointer);
                if (algorithmPointer != IntPtr.Zero) Marshal.FreeHGlobal(algorithmPointer);
                if (fileNamePointer != IntPtr.Zero) Marshal.FreeHGlobal(fileNamePointer);
                if (subjectGuidPointer != IntPtr.Zero) Marshal.FreeHGlobal(subjectGuidPointer);
            }
        }

        private static WordMacroProjectContentBindingResult NativeFailure(string message) =>
            new WordMacroProjectContentBindingResult(true, false,
                message + " Windows error " + Marshal.GetLastWin32Error() + ".");

        private static bool FixedTimeEquals(byte[] left, byte[] right) {
            int difference = left.Length ^ right.Length;
            int maximum = Math.Max(left.Length, right.Length);
            for (int index = 0; index < maximum; index++) {
                byte leftValue = index < left.Length ? left[index] : (byte)0;
                byte rightValue = index < right.Length ? right[index] : (byte)0;
                difference |= leftValue ^ rightValue;
            }
            return difference == 0;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct CryptDataBlob {
            internal uint cbData;
            internal IntPtr pbData;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct CryptAlgorithmIdentifier {
            internal IntPtr pszObjId;
            internal CryptDataBlob Parameters;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct CryptAttributeTypeValue {
            internal IntPtr pszObjId;
            internal CryptDataBlob Value;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct SipIndirectData {
            internal CryptAttributeTypeValue Data;
            internal CryptAlgorithmIdentifier DigestAlgorithm;
            internal CryptDataBlob Digest;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct SipSubjectInfo {
            internal uint cbSize;
            internal IntPtr pgSubjectType;
            internal IntPtr hFile;
            internal IntPtr pwsFileName;
            internal IntPtr pwsDisplayName;
            internal uint dwReserved1;
            internal uint dwIntVersion;
            internal IntPtr hProv;
            internal CryptAlgorithmIdentifier DigestAlgorithm;
            internal uint dwFlags;
            internal uint dwEncodingType;
            internal uint dwReserved2;
            internal uint fdwCAPISettings;
            internal uint fdwSecuritySettings;
            internal uint dwIndex;
            internal uint dwUnionChoice;
            internal IntPtr psFlat;
            internal IntPtr pClientData;
        }

        [DllImport("crypt32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool CryptSIPRetrieveSubjectGuid(
            string fileName,
            IntPtr fileHandle,
            out Guid subjectGuid);

        [DllImport("crypt32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool CryptSIPGetSignedDataMsg(
            ref SipSubjectInfo subjectInfo,
            out uint encodingType,
            uint index,
            ref uint signedDataBytes,
            IntPtr signedData);

        [DllImport("crypt32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool CryptSIPCreateIndirectData(
            ref SipSubjectInfo subjectInfo,
            ref uint indirectDataBytes,
            IntPtr indirectData);
    }
}
