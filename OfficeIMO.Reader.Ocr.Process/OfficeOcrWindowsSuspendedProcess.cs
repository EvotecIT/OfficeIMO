using Microsoft.Win32.SafeHandles;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace OfficeIMO.Reader.Ocr.Process;

/// <summary>Creates a Windows OCR process suspended so it cannot escape its Job Object before assignment.</summary>
internal static class OfficeOcrWindowsSuspendedProcess {
    private const uint CreateNoWindow = 0x08000000;
    private const uint CreateSuspended = 0x00000004;
    private const uint CreateUnicodeEnvironment = 0x00000400;
    private const uint ExtendedStartupInfoPresent = 0x00080000;
    private const uint HandleFlagInherit = 0x00000001;
    private const int ProcThreadAttributeHandleList = 0x00020002;
    private const uint StartfUseStdHandles = 0x00000100;

    internal static OfficeOcrStartedProcess Start(System.Diagnostics.ProcessStartInfo startInfo, OfficeOcrProcessLifetime lifetime) {
        lifetime.PrepareWindowsJob();
        CreatePipePair(out SafeFileHandle outputRead, out SafeFileHandle outputWrite);
        using (outputWrite)
        using (outputRead) {
            CreatePipePair(out SafeFileHandle errorRead, out SafeFileHandle errorWrite);
            using (errorWrite)
            using (errorRead) {
                CreateInputPipePair(out SafeFileHandle inputRead, out SafeFileHandle inputWrite);
                using (inputRead)
                using (inputWrite) {
                    using ProcessAttributeList attributes = ProcessAttributeList.Create(new[] {
                        inputRead.DangerousGetHandle(),
                        outputWrite.DangerousGetHandle(),
                        errorWrite.DangerousGetHandle()
                    });
                    var startupInfo = new StartupInfoEx {
                        StartupInfo = new StartupInfo {
                            Size = Marshal.SizeOf<StartupInfoEx>(),
                            Flags = StartfUseStdHandles,
                            StandardInput = inputRead.DangerousGetHandle(),
                            StandardOutput = outputWrite.DangerousGetHandle(),
                            StandardError = errorWrite.DangerousGetHandle()
                        },
                        AttributeList = attributes.Pointer
                    };
                    var commandLine = new StringBuilder(OfficeOcrProcessRunner.QuoteArgument(startInfo.FileName)
                        + (string.IsNullOrWhiteSpace(startInfo.Arguments) ? string.Empty : " " + startInfo.Arguments));
                    IntPtr environment = BuildEnvironmentBlock(startInfo);
                    ProcessInformation processInformation = default;
                    try {
                        if (!CreateProcess(
                            null,
                            commandLine,
                            IntPtr.Zero,
                            IntPtr.Zero,
                            inheritHandles: true,
                            CreateSuspended | CreateNoWindow | CreateUnicodeEnvironment | ExtendedStartupInfoPresent,
                            environment,
                            string.IsNullOrWhiteSpace(startInfo.WorkingDirectory) ? null : startInfo.WorkingDirectory,
                            ref startupInfo,
                            out processInformation)) {
                            throw new Win32Exception(Marshal.GetLastWin32Error(), "Unable to create the suspended OCR process.");
                        }

                        using var processHandle = new SafeFileHandle(processInformation.Process, ownsHandle: true);
                        using var threadHandle = new SafeFileHandle(processInformation.Thread, ownsHandle: true);
                        if (!lifetime.AssignSuspendedWindowsProcess(processHandle.DangerousGetHandle())) {
                            int error = Marshal.GetLastWin32Error();
                            _ = TerminateProcess(processHandle.DangerousGetHandle(), 1);
                            throw new Win32Exception(error, "Unable to assign the suspended OCR process to its Windows Job Object.");
                        }

                        System.Diagnostics.Process process;
                        try {
                            process = System.Diagnostics.Process.GetProcessById(unchecked((int)processInformation.ProcessId));
                            process.EnableRaisingEvents = true;
                        } catch {
                            _ = TerminateProcess(processHandle.DangerousGetHandle(), 1);
                            throw;
                        }

                        if (ResumeThread(threadHandle.DangerousGetHandle()) == uint.MaxValue) {
                            int error = Marshal.GetLastWin32Error();
                            _ = TerminateProcess(processHandle.DangerousGetHandle(), 1);
                            process.Dispose();
                            throw new Win32Exception(error, "Unable to resume the OCR process after Job Object assignment.");
                        }

                        var output = new StreamReader(
                            new FileStream(TransferOwnership(outputRead), FileAccess.Read, 4096, isAsync: false),
                            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: false),
                            detectEncodingFromByteOrderMarks: true);
                        var errorReader = new StreamReader(
                            new FileStream(TransferOwnership(errorRead), FileAccess.Read, 4096, isAsync: false),
                            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: false),
                            detectEncodingFromByteOrderMarks: true);
                        return new OfficeOcrStartedProcess(process, output, errorReader);
                    } finally {
                        if (environment != IntPtr.Zero) Marshal.FreeHGlobal(environment);
                    }
                }
            }
        }
    }

    private static SafeFileHandle TransferOwnership(SafeFileHandle source) {
        IntPtr handle = source.DangerousGetHandle();
        source.SetHandleAsInvalid();
        return new SafeFileHandle(handle, ownsHandle: true);
    }

    private static void CreatePipePair(out SafeFileHandle read, out SafeFileHandle write) {
        var securityAttributes = new SecurityAttributes {
            Length = Marshal.SizeOf<SecurityAttributes>(),
            InheritHandle = true
        };
        if (!CreatePipe(out read, out write, ref securityAttributes, 0)) {
            throw new Win32Exception(Marshal.GetLastWin32Error(), "Unable to create an OCR process pipe.");
        }
        if (!SetHandleInformation(read.DangerousGetHandle(), HandleFlagInherit, 0)) {
            int error = Marshal.GetLastWin32Error();
            read.Dispose();
            write.Dispose();
            throw new Win32Exception(error, "Unable to restrict OCR process pipe inheritance.");
        }
    }

    private static void CreateInputPipePair(out SafeFileHandle read, out SafeFileHandle write) {
        var securityAttributes = new SecurityAttributes {
            Length = Marshal.SizeOf<SecurityAttributes>(),
            InheritHandle = true
        };
        if (!CreatePipe(out read, out write, ref securityAttributes, 0)) {
            throw new Win32Exception(Marshal.GetLastWin32Error(), "Unable to create an OCR process input pipe.");
        }
        if (!SetHandleInformation(write.DangerousGetHandle(), HandleFlagInherit, 0)) {
            int error = Marshal.GetLastWin32Error();
            read.Dispose();
            write.Dispose();
            throw new Win32Exception(error, "Unable to restrict OCR process input-pipe inheritance.");
        }
    }

    private static IntPtr BuildEnvironmentBlock(System.Diagnostics.ProcessStartInfo startInfo) {
        string[] keys = startInfo.EnvironmentVariables.Keys.Cast<string>()
            .OrderBy(static key => key, StringComparer.OrdinalIgnoreCase)
            .ToArray();
        var block = new StringBuilder();
        foreach (string key in keys) {
            block.Append(key).Append('=').Append(startInfo.EnvironmentVariables[key]).Append('\0');
        }
        if (keys.Length == 0) block.Append('\0');
        return Marshal.StringToHGlobalUni(block.ToString());
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CreatePipe(out SafeFileHandle readPipe, out SafeFileHandle writePipe, ref SecurityAttributes pipeAttributes, uint size);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetHandleInformation(IntPtr handle, uint mask, uint flags);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CreateProcess(
        string? applicationName,
        [In, Out] StringBuilder commandLine,
        IntPtr processAttributes,
        IntPtr threadAttributes,
        [MarshalAs(UnmanagedType.Bool)] bool inheritHandles,
        uint creationFlags,
        IntPtr environment,
        string? currentDirectory,
        ref StartupInfoEx startupInfo,
        out ProcessInformation processInformation);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool InitializeProcThreadAttributeList(
        IntPtr attributeList,
        int attributeCount,
        int flags,
        ref IntPtr size);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UpdateProcThreadAttribute(
        IntPtr attributeList,
        uint flags,
        IntPtr attribute,
        IntPtr value,
        IntPtr size,
        IntPtr previousValue,
        IntPtr returnSize);

    [DllImport("kernel32.dll")]
    private static extern void DeleteProcThreadAttributeList(IntPtr attributeList);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern uint ResumeThread(IntPtr thread);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool TerminateProcess(IntPtr process, uint exitCode);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct StartupInfo {
        internal int Size;
        internal string? Reserved;
        internal string? Desktop;
        internal string? Title;
        internal uint X;
        internal uint Y;
        internal uint XSize;
        internal uint YSize;
        internal uint XCountChars;
        internal uint YCountChars;
        internal uint FillAttribute;
        internal uint Flags;
        internal ushort ShowWindow;
        internal ushort Reserved2;
        internal IntPtr Reserved2Pointer;
        internal IntPtr StandardInput;
        internal IntPtr StandardOutput;
        internal IntPtr StandardError;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct StartupInfoEx {
        internal StartupInfo StartupInfo;
        internal IntPtr AttributeList;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct ProcessInformation {
        internal IntPtr Process;
        internal IntPtr Thread;
        internal uint ProcessId;
        internal uint ThreadId;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct SecurityAttributes {
        internal int Length;
        internal IntPtr SecurityDescriptor;
        [MarshalAs(UnmanagedType.Bool)] internal bool InheritHandle;
    }

    private sealed class ProcessAttributeList : IDisposable {
        private IntPtr _handleList;

        private ProcessAttributeList(IntPtr pointer, IntPtr handleList) {
            Pointer = pointer;
            _handleList = handleList;
        }

        internal IntPtr Pointer { get; private set; }

        internal static ProcessAttributeList Create(IReadOnlyList<IntPtr> handles) {
            IntPtr size = IntPtr.Zero;
            _ = InitializeProcThreadAttributeList(IntPtr.Zero, 1, 0, ref size);
            if (size == IntPtr.Zero) {
                throw new Win32Exception(Marshal.GetLastWin32Error(), "Unable to size the OCR process attribute list.");
            }

            IntPtr attributes = Marshal.AllocHGlobal(size);
            IntPtr handleList = Marshal.AllocHGlobal(IntPtr.Size * handles.Count);
            bool initialized = false;
            try {
                if (!InitializeProcThreadAttributeList(attributes, 1, 0, ref size)) {
                    throw new Win32Exception(Marshal.GetLastWin32Error(), "Unable to initialize the OCR process attribute list.");
                }
                initialized = true;
                for (int index = 0; index < handles.Count; index++) {
                    Marshal.WriteIntPtr(handleList, index * IntPtr.Size, handles[index]);
                }
                if (!UpdateProcThreadAttribute(
                    attributes,
                    0,
                    new IntPtr(ProcThreadAttributeHandleList),
                    handleList,
                    new IntPtr(IntPtr.Size * handles.Count),
                    IntPtr.Zero,
                    IntPtr.Zero)) {
                    throw new Win32Exception(Marshal.GetLastWin32Error(), "Unable to restrict inherited OCR process handles.");
                }
                return new ProcessAttributeList(attributes, handleList);
            } catch {
                if (initialized) DeleteProcThreadAttributeList(attributes);
                Marshal.FreeHGlobal(attributes);
                Marshal.FreeHGlobal(handleList);
                throw;
            }
        }

        public void Dispose() {
            if (Pointer != IntPtr.Zero) {
                DeleteProcThreadAttributeList(Pointer);
                Marshal.FreeHGlobal(Pointer);
                Pointer = IntPtr.Zero;
            }
            if (_handleList != IntPtr.Zero) {
                Marshal.FreeHGlobal(_handleList);
                _handleList = IntPtr.Zero;
            }
        }
    }
}
