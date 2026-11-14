using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace OfficeIMO.Core.Internal {
    internal static partial class OfficeTemporaryFile {
        // OWNER RIGHTS is resolved against the new file's owner. The protected DACL excludes
        // broad inheritable directory permissions before any output bytes are written.
        private const string OwnerOnlyDacl = "D:P(A;;GA;;;OW)(A;;GA;;;SY)(A;;GA;;;BA)";

        internal static FileStream CreateWindowsOwnerOnly(string path, int bufferSize, FileOptions options) {
            if (!ConvertStringSecurityDescriptorToSecurityDescriptor(OwnerOnlyDacl, 1, out IntPtr descriptor, out _))
                throw new IOException("Unable to prepare owner-only Windows staging permissions (OS error " + Marshal.GetLastWin32Error() + ").");
            try {
                var attributes = new SecurityAttributes {
                    Length = Marshal.SizeOf<SecurityAttributes>(),
                    SecurityDescriptor = descriptor,
                    InheritHandle = false
                };
                const uint genericReadWrite = 0xC0000000U;
                const uint deleteAccess = 0x00010000U;
                const uint normalAttribute = 0x00000080U;
                uint access = genericReadWrite | ((options & FileOptions.DeleteOnClose) != 0 ? deleteAccess : 0);
                uint flags = normalAttribute | (unchecked((uint)options) & ~unchecked((uint)FileOptions.Asynchronous));
                SafeFileHandle handle = CreateFile(ToExtendedPath(path), access, 0, ref attributes, 1, flags, IntPtr.Zero);
                if (handle.IsInvalid) {
                    int error = Marshal.GetLastWin32Error();
                    handle.Dispose();
                    throw new IOException("Unable to create owner-only Windows staging file (OS error " + error + ").");
                }
                try {
                    // A synchronous handle supports FileStream's asynchronous fallback without exposing the file.
                    return new FileStream(handle, FileAccess.ReadWrite, bufferSize, isAsync: false);
                } catch {
                    handle.Dispose();
                    throw;
                }
            } finally {
                LocalFree(descriptor);
            }
        }

        private const int FileObject = 1;
        private const uint DaclInformation = 0x00000004U;
        private const uint ProtectedDaclInformation = 0x80000000U;
        private const uint UnprotectedDaclInformation = 0x20000000U;
        private const ushort DaclProtectedControl = 0x1000;
        private const int InvalidFunctionError = 1;
        private const int FileNotFoundError = 2;
        private const int NotSupportedError = 50;

        // File.Move keeps the staging file's protected DACL. Before a complete staging file is
        // published as a new file, replace it with an empty unprotected DACL so the file takes the
        // destination folder's inheritable permissions, as a directly created file would. A staging
        // file that already inherits them is left alone: changing a DACL needs WRITE_DAC, which
        // shares such as Azure Files can withhold from users who may still create files.
        internal static void ApplyInheritedWindowsPermissions(string path) {
            int readError = ReadWindowsDacl(path, out IntPtr current, out _, out bool isProtected);
            LocalFree(current);
            if (IsWithoutAclSupport(readError) || (readError == 0 && !isProtected)) return;

            if (!ConvertStringSecurityDescriptorToSecurityDescriptor("D:", 1, out IntPtr descriptor, out _))
                throw new IOException("Unable to prepare inherited Windows permissions (OS error " + Marshal.GetLastWin32Error() + ").");
            try {
                if (!GetSecurityDescriptorDacl(descriptor, out bool present, out IntPtr dacl, out _) || !present || dacl == IntPtr.Zero)
                    throw new IOException("Unable to prepare inherited Windows permissions (OS error " + Marshal.GetLastWin32Error() + ").");
                int error = SetNamedSecurityInfo(ToExtendedPath(path), FileObject, DaclInformation | UnprotectedDaclInformation,
                    IntPtr.Zero, IntPtr.Zero, dacl, IntPtr.Zero);
                if (error != 0 && !IsWithoutAclSupport(error))
                    throw new IOException("Unable to apply the destination folder's permissions to '" + path + "' (OS error " + error + ").");
            } finally {
                LocalFree(descriptor);
            }
        }

        // A move that replaces an existing file must keep that file's permissions, as File.Replace
        // does, so a restricted document is never widened to the folder's permissions. When they
        // cannot be read or applied the staging file stays owner-only and the commit fails.
        internal static void CopyWindowsPermissions(string sourcePath, string destinationPath) {
            int error = ReadWindowsDacl(sourcePath, out IntPtr descriptor, out IntPtr dacl, out bool isProtected);
            try {
                if (error == FileNotFoundError) {
                    ApplyInheritedWindowsPermissions(destinationPath);
                    return;
                }
                if (IsWithoutAclSupport(error)) return;
                if (error != 0)
                    throw new IOException("Unable to read the permissions of '" + sourcePath + "' (OS error " + error + ").");
                error = SetNamedSecurityInfo(ToExtendedPath(destinationPath), FileObject,
                    DaclInformation | (isProtected ? ProtectedDaclInformation : UnprotectedDaclInformation),
                    IntPtr.Zero, IntPtr.Zero, dacl, IntPtr.Zero);
                if (error != 0 && !IsWithoutAclSupport(error))
                    throw new IOException("Unable to keep the permissions of '" + sourcePath + "' (OS error " + error + ").");
            } finally {
                LocalFree(descriptor);
            }
        }

        // Returns the Win32 error. The caller frees the descriptor, which owns the DACL.
        private static int ReadWindowsDacl(string path, out IntPtr descriptor, out IntPtr dacl, out bool isProtected) {
            isProtected = false;
            int error = GetNamedSecurityInfo(ToExtendedPath(path), FileObject, DaclInformation,
                IntPtr.Zero, IntPtr.Zero, out dacl, IntPtr.Zero, out descriptor);
            if (error != 0) return error;
            if (!GetSecurityDescriptorControl(descriptor, out ushort control, out _)) return Marshal.GetLastWin32Error();
            isProtected = (control & DaclProtectedControl) != 0;
            return 0;
        }

        // These calls go straight to Win32, which limits ordinary paths to MAX_PATH even where
        // .NET's own file APIs accept longer ones. The extended-length form lifts that limit, so a
        // staging file next to a deeply nested destination can still be created and secured.
        internal static string ToExtendedPath(string path) {
            string fullPath = Path.GetFullPath(path);
            if (fullPath.StartsWith(@"\\?\", StringComparison.Ordinal) || fullPath.StartsWith(@"\\.\", StringComparison.Ordinal))
                return fullPath;
            if (fullPath.StartsWith(@"\\", StringComparison.Ordinal))
                return @"\\?\UNC\" + fullPath.Substring(2);
            return @"\\?\" + fullPath;
        }

        // A file system without ACL support never applied the staging DACL either.
        private static bool IsWithoutAclSupport(int error) => error == InvalidFunctionError || error == NotSupportedError;

        [StructLayout(LayoutKind.Sequential)]
        private struct SecurityAttributes {
            public int Length;
            public IntPtr SecurityDescriptor;
            [MarshalAs(UnmanagedType.Bool)] public bool InheritHandle;
        }

        [DllImport("advapi32.dll", EntryPoint = "ConvertStringSecurityDescriptorToSecurityDescriptorW", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool ConvertStringSecurityDescriptorToSecurityDescriptor(
            string securityDescriptor, uint revision, out IntPtr descriptor, out uint descriptorSize);

        [DllImport("kernel32.dll", EntryPoint = "CreateFileW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern SafeFileHandle CreateFile(string path, uint access, uint share, ref SecurityAttributes attributes,
            uint creationDisposition, uint flags, IntPtr templateFile);

        [DllImport("advapi32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetSecurityDescriptorDacl(IntPtr descriptor,
            [MarshalAs(UnmanagedType.Bool)] out bool daclPresent, out IntPtr dacl,
            [MarshalAs(UnmanagedType.Bool)] out bool daclDefaulted);

        [DllImport("advapi32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetSecurityDescriptorControl(IntPtr descriptor, out ushort control, out uint revision);

        [DllImport("advapi32.dll", EntryPoint = "GetNamedSecurityInfoW", CharSet = CharSet.Unicode)]
        private static extern int GetNamedSecurityInfo(string objectName, int objectType, uint securityInfo,
            IntPtr owner, IntPtr group, out IntPtr dacl, IntPtr sacl, out IntPtr descriptor);

        [DllImport("advapi32.dll", EntryPoint = "SetNamedSecurityInfoW", CharSet = CharSet.Unicode)]
        private static extern int SetNamedSecurityInfo(string objectName, int objectType, uint securityInfo,
            IntPtr owner, IntPtr group, IntPtr dacl, IntPtr sacl);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr LocalFree(IntPtr memory);
    }
}
