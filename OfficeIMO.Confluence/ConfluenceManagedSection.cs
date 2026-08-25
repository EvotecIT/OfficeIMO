using System.Buffers;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Confluence;

/// <summary>Behavior when a managed section marker pair does not exist.</summary>
public enum ConfluenceMissingSectionBehavior {
    Fail,
    Append,
}

/// <summary>Pure result of replacing one OfficeIMO-managed page-body section.</summary>
public sealed class ConfluenceManagedSectionResult {
    internal ConfluenceManagedSectionResult(string sectionId, string originalBody, string updatedBody, bool created) {
        SectionId = sectionId;
        OriginalBody = originalBody;
        UpdatedBody = updatedBody;
        WasCreated = created;
        OriginalSha256 = ComputeHash(originalBody);
        UpdatedSha256 = ComputeHash(updatedBody);
    }

    public string SectionId { get; }
    public string OriginalBody { get; }
    public string UpdatedBody { get; }
    public bool WasCreated { get; }
    public bool Changed => !string.Equals(OriginalBody, UpdatedBody, StringComparison.Ordinal);
    public string OriginalSha256 { get; }
    public string UpdatedSha256 { get; }

    private static string ComputeHash(string value) {
        using SHA256 sha = SHA256.Create();
        value ??= string.Empty;
        const int charactersPerChunk = 4096;
        byte[] buffer = ArrayPool<byte>.Shared.Rent(Encoding.UTF8.GetMaxByteCount(charactersPerChunk));
        byte[] hash;
        try {
            int offset = 0;
            while (offset < value.Length) {
                int characterCount = Math.Min(charactersPerChunk, value.Length - offset);
                if (offset + characterCount < value.Length
                    && char.IsHighSurrogate(value[offset + characterCount - 1])
                    && char.IsLowSurrogate(value[offset + characterCount])) {
                    characterCount--;
                }
                int byteCount = Encoding.UTF8.GetBytes(value, offset, characterCount, buffer, 0);
                sha.TransformBlock(buffer, 0, byteCount, buffer, 0);
                offset += characterCount;
            }
            sha.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
            hash = sha.Hash ?? throw new CryptographicException("SHA-256 did not produce a hash.");
        } finally {
            ArrayPool<byte>.Shared.Return(buffer, clearArray: true);
        }
        var builder = new StringBuilder(hash.Length * 2);
        foreach (byte item in hash) builder.Append(item.ToString("x2", System.Globalization.CultureInfo.InvariantCulture));
        return builder.ToString();
    }
}

/// <summary>Safely replaces content between stable OfficeIMO markers while preserving the rest of the page.</summary>
public static class ConfluenceManagedSection {
    private static readonly Regex SectionIdPattern = new Regex("^[A-Za-z0-9._-]{1,100}$", RegexOptions.CultureInvariant);
    private const string MarkerPrefix = "<!-- officeimo:section:";

    public static ConfluenceManagedSectionResult Apply(
        string existingBody,
        string sectionId,
        string replacement,
        ConfluenceMissingSectionBehavior missingBehavior = ConfluenceMissingSectionBehavior.Fail) {
        existingBody ??= string.Empty;
        replacement ??= string.Empty;
        ValidateSectionId(sectionId);
        if (replacement.IndexOf(MarkerPrefix, StringComparison.OrdinalIgnoreCase) >= 0) throw new ArgumentException("Replacement content cannot contain OfficeIMO section markers.", nameof(replacement));

        string start = StartMarker(sectionId);
        string end = EndMarker(sectionId);
        int startIndex = existingBody.IndexOf(start, StringComparison.Ordinal);
        int endIndex = existingBody.IndexOf(end, StringComparison.Ordinal);

        if (startIndex < 0 && endIndex < 0) {
            if (missingBehavior == ConfluenceMissingSectionBehavior.Fail) throw new InvalidOperationException("Managed section '" + sectionId + "' does not exist.");
            string separator = existingBody.Length == 0 ? string.Empty : existingBody.EndsWith("\n", StringComparison.Ordinal) ? "\n" : "\n\n";
            string appended = existingBody + separator + start + "\n" + replacement + "\n" + end;
            return new ConfluenceManagedSectionResult(sectionId, existingBody, appended, created: true);
        }

        if (startIndex < 0 || endIndex < 0 || endIndex < startIndex) throw new InvalidOperationException("Managed section '" + sectionId + "' has unmatched or reversed markers.");
        if (existingBody.IndexOf(start, startIndex + start.Length, StringComparison.Ordinal) >= 0 || existingBody.IndexOf(end, endIndex + end.Length, StringComparison.Ordinal) >= 0) {
            throw new InvalidOperationException("Managed section '" + sectionId + "' appears more than once.");
        }

        int contentStart = startIndex + start.Length;
        string updated = CreateUpdatedBody(existingBody, replacement, contentStart, endIndex);
        return new ConfluenceManagedSectionResult(sectionId, existingBody, updated, created: false);
    }

    public static string StartMarker(string sectionId) { ValidateSectionId(sectionId); return MarkerPrefix + sectionId + ":start -->"; }
    public static string EndMarker(string sectionId) { ValidateSectionId(sectionId); return MarkerPrefix + sectionId + ":end -->"; }

    private static void ValidateSectionId(string sectionId) {
        if (string.IsNullOrWhiteSpace(sectionId) || !SectionIdPattern.IsMatch(sectionId)) throw new ArgumentException("Section id must contain 1-100 letters, digits, dots, underscores, or hyphens.", nameof(sectionId));
    }

    private static string CreateUpdatedBody(string existingBody, string replacement, int contentStart, int endIndex) {
#if NETSTANDARD2_0 || NETFRAMEWORK
        return existingBody.Substring(0, contentStart) + "\n" + replacement + "\n" + existingBody.Substring(endIndex);
#else
        int suffixLength = existingBody.Length - endIndex;
        int length = checked(contentStart + 1 + replacement.Length + 1 + suffixLength);
        var state = new SectionReplacementState(existingBody, replacement, contentStart, endIndex);
        return string.Create(length, state, static (destination, replacementState) => {
            int offset = 0;
            replacementState.ExistingBody.AsSpan(0, replacementState.ContentStart).CopyTo(destination);
            offset += replacementState.ContentStart;
            destination[offset++] = '\n';
            replacementState.Replacement.AsSpan().CopyTo(destination.Slice(offset));
            offset += replacementState.Replacement.Length;
            destination[offset++] = '\n';
            replacementState.ExistingBody.AsSpan(replacementState.EndIndex).CopyTo(destination.Slice(offset));
        });
#endif
    }

#if !NETSTANDARD2_0 && !NETFRAMEWORK
    private readonly struct SectionReplacementState {
        internal SectionReplacementState(string existingBody, string replacement, int contentStart, int endIndex) {
            ExistingBody = existingBody;
            Replacement = replacement;
            ContentStart = contentStart;
            EndIndex = endIndex;
        }

        internal string ExistingBody { get; }
        internal string Replacement { get; }
        internal int ContentStart { get; }
        internal int EndIndex { get; }
    }
#endif
}
