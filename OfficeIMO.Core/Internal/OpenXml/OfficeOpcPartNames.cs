namespace OfficeIMO.Core.Internal.OpenXml;

/// <summary>
/// Part-name rules from ECMA-376 Part 2 shared by the package reader and the package security inspector: canonical
/// part paths, the once-decoded identity two spellings of a part name share, and relationship target resolution.
/// </summary>
internal static class OfficeOpcPartNames {
    private const string RelationshipsExtension = ".rels";

    // These checks run for every package part and relationship target, so they walk the path instead of
    // splitting it into segment strings.
    internal static bool IsDirectoryOrEmptySegmentTarget(string path) {
        if (path.Length == 0) return true;
        int lastSeparator = -1;
        for (int index = 0; index < path.Length; index++) {
            if (path[index] != '/') continue;
            // Every segment after a separator must be non-empty.
            if (index == path.Length - 1 || path[index + 1] == '/') return true;
            lastSeparator = index;
        }
        return IsDotSegment(path, lastSeparator + 1, path.Length - lastSeparator - 1);
    }

    internal static bool IsCanonicalPartPath(string path) {
        if (path.Length < 2 || path[0] != '/' || IsDirectoryOrEmptySegmentTarget(path)) return false;
        int segmentStart = 1;
        for (int index = 1; index <= path.Length; index++) {
            if (index < path.Length && path[index] != '/') {
                char value = path[index];
                if (value == '\\' || char.IsControl(value)) return false;
                continue;
            }
            if (IsDotSegment(path, segmentStart, index - segmentStart)) return false;
            segmentStart = index + 1;
        }
        return true;
    }

    internal static bool IsDotSegment(string path, int start, int length) =>
        length == 1 && path[start] == '.'
        || length == 2 && path[start] == '.' && path[start + 1] == '.';

    /// <summary>Decodes each segment once; encoded separators, dot segments and controls cannot identify another part.</summary>
    internal static string? DecodePartIdentity(string path) {
        if (!IsCanonicalPartPath(path)) return null;
        // A canonical path without escapes is its own identity.
        if (path.IndexOf('%') < 0) return path;
        string[] segments = path.Split('/');
        for (int index = 1; index < segments.Length; index++) {
            string? decoded = DecodePartIdentitySegment(segments[index]);
            if (decoded == null) return null;
            segments[index] = decoded;
        }
        return string.Join("/", segments);
    }

    internal static string? DecodePartIdentitySegment(string segment) {
        for (int offset = 0; offset < segment.Length; offset++) {
            if (segment[offset] != '%') continue;
            if (offset + 2 >= segment.Length || !IsPartIdentityHexDigit(segment[offset + 1])
                || !IsPartIdentityHexDigit(segment[offset + 2])) return null;
            offset += 2;
        }
        string decoded;
        try { decoded = Uri.UnescapeDataString(segment); }
        catch (UriFormatException) { return null; }
        if (decoded.Length == 0 || decoded == "." || decoded == ".."
            || decoded.IndexOf('/') >= 0 || decoded.IndexOf('\\') >= 0) return null;
        foreach (char value in decoded) if (char.IsControl(value)) return null;
        return decoded;
    }

    internal static bool IsPartIdentityHexDigit(char value) => value >= '0' && value <= '9'
        || value >= 'a' && value <= 'f' || value >= 'A' && value <= 'F';

    /// <summary>The relationships part of a source part, or of the package when the source is null or <c>/</c>.</summary>
    /// <summary>
    /// The relationship target that leads from <paramref name="sourcePartName"/> to <paramref name="targetPartName"/>:
    /// a path relative to the source part's folder, climbing with <c>..</c> as far as the folders differ.
    /// </summary>
    internal static string GetRelativeTarget(string sourcePartName, string targetPartName) {
        string[] source = sourcePartName.Split('/');
        string[] target = targetPartName.Split('/');
        // Folders only: the last segment of the source is the part itself.
        int common = 0;
        while (common < source.Length - 1 && common < target.Length - 1 && string.Equals(source[common], target[common], StringComparison.Ordinal)) common++;
        var builder = new System.Text.StringBuilder();
        for (int index = common; index < source.Length - 1; index++) builder.Append("../");
        for (int index = common; index < target.Length; index++) {
            if (index > common) builder.Append('/');
            builder.Append(target[index]);
        }
        return builder.ToString();
    }

    internal static string GetRelationshipsPartName(string? sourcePartName) {
        if (sourcePartName == null || sourcePartName == "/") return "/_rels/.rels";
        int slash = sourcePartName.LastIndexOf('/');
        return sourcePartName.Substring(0, slash + 1) + "_rels/" + sourcePartName.Substring(slash + 1) + ".rels";
    }

    private static readonly char[] QueryOrFragment = { '?', '#' };

    /// <summary>Resolves a relationship target against the source of its relationships part, returning the target
    /// part name or null for external, absolute, malformed or escaping targets. The target's spelling is kept.</summary>
    internal static string? ResolveRelationshipTarget(string relationshipPartName, string? target) =>
        string.IsNullOrWhiteSpace(target) ? null : ResolveRelationshipTargetFrom(GetRelationshipSourceDirectory(relationshipPartName), target);

    /// <summary>The directory relationship targets of a relationships part resolve against, or null when the part
    /// name is not a relationships part. Callers resolving many targets of one part compute it once.</summary>
    internal static string? GetRelationshipSourceDirectory(string relationshipPartName) {
        string normalizedRelationshipPart = relationshipPartName.Replace('\\', '/');
        int relationshipsSegment = normalizedRelationshipPart.LastIndexOf("/_rels/", StringComparison.OrdinalIgnoreCase);
        if (relationshipsSegment < 0) return null;

        string relationshipFile = normalizedRelationshipPart.Substring(relationshipsSegment + 7);
        if (string.Equals(relationshipFile, ".rels", StringComparison.OrdinalIgnoreCase)) return "/";
        if (!relationshipFile.EndsWith(RelationshipsExtension, StringComparison.OrdinalIgnoreCase)) return null;
        string sourceFile = relationshipFile.Substring(0, relationshipFile.Length - RelationshipsExtension.Length);
        string sourcePart = normalizedRelationshipPart.Substring(0, relationshipsSegment) + "/" + sourceFile;
        int separator = sourcePart.LastIndexOf('/');
        return separator <= 0 ? "/" : sourcePart.Substring(0, separator + 1);
    }

    /// <summary>Resolves a target against a directory from <see cref="GetRelationshipSourceDirectory"/>.</summary>
    internal static string? ResolveRelationshipTargetFrom(string? sourceDirectory, string? target) {
        if (sourceDirectory == null || string.IsNullOrWhiteSpace(target)) return null;
        string normalizedTarget = target!.Replace('\\', '/');
        // OPC uses one leading slash for a package-root part reference.
        // Authority/UNC targets and absolute URI schemes remain unresolved.
        if (normalizedTarget.StartsWith("//", StringComparison.Ordinal)) return null;
        // Only a target with a scheme separator, or a backslash path, can parse as an absolute URI.
        if (!target.StartsWith("/", StringComparison.Ordinal) && (target.IndexOf(':') >= 0 || target[0] == '\\')
            && Uri.TryCreate(target, UriKind.Absolute, out Uri? absolute) && absolute.IsAbsoluteUri) return null;

        int suffix = normalizedTarget.IndexOfAny(QueryOrFragment);
        if (suffix >= 0) normalizedTarget = normalizedTarget.Substring(0, suffix);
        if (IsDirectoryOrEmptySegmentTarget(normalizedTarget)) return null;
        string combined = normalizedTarget.StartsWith("/", StringComparison.Ordinal)
            ? normalizedTarget
            : sourceDirectory + normalizedTarget;
        if (IsDirectoryOrEmptySegmentTarget(combined)) return null;
        // A path without dot segments, escapes or controls is already its own resolved spelling.
        if (IsPlainPartPath(combined)) return combined;
        var segments = new List<string>();
        foreach (string segment in combined.Split('/')) {
            if (segment.Length == 0 || segment == ".") continue;
            if (segment == "..") {
                if (segments.Count == 0) return null;
                segments.RemoveAt(segments.Count - 1);
            } else {
                // Validate the original spelling before a later parent segment
                // can remove it; retain the spelling for once-decoded lookup.
                if (DecodePartIdentitySegment(segment) == null) return null;
                segments.Add(segment);
            }
        }
        return "/" + string.Join("/", segments);
    }

    // True for "/a/b" paths whose segments are neither dot segments nor carry escapes or control characters. Callers
    // have already rejected empty segments and a trailing separator.
    private static bool IsPlainPartPath(string path) {
        if (path.Length == 0 || path[0] != '/') return false;
        for (int index = 0; index < path.Length; index++) {
            char value = path[index];
            if (value == '%' || char.IsControl(value)) return false;
            if (value == '.' && path[index - 1] == '/') {
                int end = index + 1;
                if (end < path.Length && path[end] == '.') end++;
                if (end == path.Length || path[end] == '/') return false;
            }
        }
        return true;
    }
}
