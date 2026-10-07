using System.IO;

namespace OfficeIMO.Internal {
    internal static partial class OfficePathIdentity {
        // Published workflow assemblies call the three-argument signature.
        internal static FileStream OpenRegularFileForRead(string path, string physicalRoot, int bufferSize) =>
            OpenRegularFileForRead(path, physicalRoot, bufferSize, FileShare.Read);
    }
}
