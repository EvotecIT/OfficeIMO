using OfficeIMO.Internal;

namespace OfficeIMO.Workflows;

/// <summary>Routes workflow path identity through the filesystem-aware OfficeIMO owner.</summary>
internal static class OfficeWorkflowPathIdentity {
    internal static string Normalize(string path) => OfficePathIdentity.Normalize(path);

    internal static bool AreEquivalent(string left, string right) => OfficePathIdentity.AreEquivalent(left, right);

    internal static bool IsSameOrDescendant(string candidatePath, string rootPath) =>
        OfficePathIdentity.IsSameOrDescendant(candidatePath, rootPath);

    internal static StringComparer GetComparer(string path) => OfficePathIdentity.GetComparer(path);

    internal static StringComparison GetComparison(string path) => OfficePathIdentity.GetComparison(path);
}
