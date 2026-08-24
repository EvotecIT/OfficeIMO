using OfficeIMO.Internal;

namespace OfficeIMO.Tool;

/// <summary>Provides link-aware path comparisons shared by command areas that write files.</summary>
internal static class OfficeImoToolPathSafety {
    /// <summary>Determines whether two paths resolve to the same file-system location.</summary>
    internal static bool PathsEqual(string firstPath, string secondPath) =>
        OfficePathIdentity.AreEquivalent(firstPath, secondPath);

    /// <summary>Resolves every existing symbolic-link segment while preserving a non-existing tail.</summary>
    internal static string ResolveExistingLinks(string path) =>
        OfficePathIdentity.ResolvePhysicalPath(path);

    /// <summary>Determines whether a resolved candidate is the parent path itself or one of its descendants.</summary>
    internal static bool IsSameOrChildPath(string parentPath, string candidatePath) =>
        OfficePathIdentity.IsSameOrDescendant(candidatePath, parentPath);
}
