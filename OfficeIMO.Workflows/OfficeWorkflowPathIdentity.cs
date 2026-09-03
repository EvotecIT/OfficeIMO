using OfficeIMO.Internal;

namespace OfficeIMO.Workflows;

/// <summary>Routes workflow path identity through the filesystem-aware OfficeIMO owner.</summary>
internal static class OfficeWorkflowPathIdentity {
    internal static bool SupportsPhysicalIdentity => OfficePathIdentity.SupportsPhysicalIdentity;

    internal static string Normalize(string path) => OfficePathIdentity.Normalize(path);

    internal static string Normalize(string path, bool caseInsensitive) =>
        OfficePathIdentity.Normalize(path, caseInsensitive);

    internal static bool AreEquivalent(string left, string right) => OfficePathIdentity.AreEquivalent(left, right);

    internal static bool IsSameOrDescendant(string candidatePath, string rootPath) =>
        OfficePathIdentity.IsSameOrDescendant(candidatePath, rootPath);

    internal static string ResolvePhysicalPath(string path) => OfficePathIdentity.ResolvePhysicalPath(path);

    internal static FileStream OpenRegularFileForRead(string path, string physicalRoot, int bufferSize) =>
        OfficePathIdentity.OpenRegularFileForRead(path, physicalRoot, bufferSize);

    internal static string GetPhysicalIdentityKey(string path, FileStream stream) =>
        OfficePathIdentity.GetPhysicalIdentityKey(path, stream.SafeFileHandle);

    internal static StringComparer GetComparer(string path) => OfficePathIdentity.GetComparer(path);

    internal static StringComparison GetComparison(string path) => OfficePathIdentity.GetComparison(path);

    internal static bool IsCaseInsensitiveFileSystem(string path) =>
        OfficePathIdentity.IsCaseInsensitiveFileSystem(path);
}
