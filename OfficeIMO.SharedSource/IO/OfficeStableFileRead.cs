using System.IO;

#if OFFICEIMO_READER_CORE
namespace OfficeIMO.Reader.Internal.IO {
#else
namespace OfficeIMO.Core.Internal.IO {
#endif
    internal static class OfficeStableFileRead {
        /// <summary>Opens a source without sharing writes, or lets the caller validate an owned snapshot if the open conflicts.</summary>
        internal static FileStream? TryOpen(string path, int bufferSize = 4096, bool useAsync = false) {
            const int sharingViolation = unchecked((int)0x80070020);
            try {
                return new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read | FileShare.Delete, bufferSize, useAsync);
            } catch (IOException exception) when (exception.HResult == sharingViolation) {
                return null;
            }
        }
    }
}
