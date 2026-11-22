using System.IO;
using System.Runtime.InteropServices;

#if OFFICEIMO_READER_CORE
namespace OfficeIMO.Reader.Internal.IO {
#else
namespace OfficeIMO.Core.Internal.IO {
#endif
    internal static class OfficeStableFileRead {
        /// <summary>Opens a source with enforced write exclusion where supported. Otherwise the caller can validate an owned snapshot.</summary>
        internal static FileStream? TryOpen(string path, int bufferSize = 4096, bool useAsync = false) {
            // FileShare on Unix is advisory, so it cannot establish the stability needed for package admission.
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return null;
            const int sharingViolation = unchecked((int)0x80070020);
            try {
                return new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read | FileShare.Delete, bufferSize, useAsync);
            } catch (IOException exception) when (exception.HResult == sharingViolation) {
                return null;
            }
        }
    }
}
