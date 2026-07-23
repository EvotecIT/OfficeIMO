using OfficeIMO.Epub;
using OfficeIMO.Reader;
using System.Linq;

namespace OfficeIMO.Reader.Epub;

internal static partial class EpubReaderAdapter {
    private static string? ResolveEpubMarkdownReference(
        string sourcePath,
        EpubChapter chapter,
        string value) {
        EpubReference reference = EpubReference.Resolve(chapter.Path, chapter.BaseHref, value);
        if (!reference.IsValid) return null;
        if (reference.Kind == EpubReferenceKind.Data) return reference.ResolvedValue;
        string? location = BuildEpubResolvedLocation(sourcePath, reference, includeFragment: true);
        return location?
            .Replace(" ", "%20")
            .Replace("(", "%28")
            .Replace(")", "%29");
    }

    private static IReadOnlyList<OfficeDocumentLink> ResolveEpubChapterLinks(
        string sourcePath,
        EpubChapter chapter,
        IReadOnlyList<OfficeDocumentLink> links,
        List<OfficeDocumentDiagnostic> diagnostics) {
        var resolvedLinks = new List<OfficeDocumentLink>(links.Count);
        foreach (OfficeDocumentLink link in links) {
            string? rawReference = !string.IsNullOrWhiteSpace(link.Uri)
                ? link.Uri
                : !string.IsNullOrWhiteSpace(link.DestinationName)
                    ? "#" + link.DestinationName
                    : null;
            if (rawReference == null) {
                resolvedLinks.Add(link);
                continue;
            }

            EpubReference reference = EpubReference.Resolve(chapter.Path, chapter.BaseHref, rawReference);
            AddEpubReferenceDiagnostic(sourcePath, chapter, reference, diagnostics);
            if (!reference.IsValid) continue;
            if (reference.Kind == EpubReferenceKind.Data) {
                resolvedLinks.Add(link);
                continue;
            }

            string? location = BuildEpubResolvedLocation(sourcePath, reference, includeFragment: true);
            if (!string.IsNullOrWhiteSpace(location)) link.Uri = location;
            resolvedLinks.Add(link);
        }
        return resolvedLinks;
    }

    private static IReadOnlyList<ReaderVisual> ResolveEpubChapterVisuals(
        string sourcePath,
        EpubChapter chapter,
        IReadOnlyList<ReaderVisual> visuals,
        IReadOnlyDictionary<string, OfficeDocumentAsset> assetsByLocation,
        List<OfficeDocumentAsset> chapterAssets,
        List<OfficeDocumentDiagnostic> diagnostics) {
        var resolvedVisuals = new List<ReaderVisual>(visuals.Count);
        foreach (ReaderVisual visual in visuals) {
            if (string.IsNullOrWhiteSpace(visual.SourceName)) {
                resolvedVisuals.Add(visual);
                continue;
            }
            if (string.Equals(visual.SourceName, "data-uri", StringComparison.Ordinal)) {
                resolvedVisuals.Add(visual);
                continue;
            }
            EpubReference reference = EpubReference.Resolve(chapter.Path, chapter.BaseHref, visual.SourceName!);
            AddEpubReferenceDiagnostic(sourcePath, chapter, reference, diagnostics);
            if (!reference.IsValid) continue;
            if (reference.Kind == EpubReferenceKind.Data) {
                resolvedVisuals.Add(visual);
                continue;
            }

            string? location = BuildEpubResolvedLocation(sourcePath, reference, includeFragment: true);
            if (!string.IsNullOrWhiteSpace(location)) visual.SourceName = location;
            OfficeDocumentAsset? asset = FindEpubAsset(sourcePath, reference, assetsByLocation);
            if (asset != null && !chapterAssets.Contains(asset)) chapterAssets.Add(asset);
            resolvedVisuals.Add(visual);
        }
        return resolvedVisuals;
    }

    private static OfficeDocumentAsset? FindEpubAsset(
        string sourcePath,
        EpubReference reference,
        IReadOnlyDictionary<string, OfficeDocumentAsset> assetsByLocation) {
        string? targetLocation = BuildEpubResolvedLocation(sourcePath, reference, includeFragment: false);
        if (string.IsNullOrWhiteSpace(targetLocation)) return null;

        assetsByLocation.TryGetValue(targetLocation!, out OfficeDocumentAsset? exact);
        if (exact != null || reference.Kind != EpubReferenceKind.Container || string.IsNullOrWhiteSpace(reference.ContainerPath)) {
            return exact;
        }

        string packageLocation = BuildVirtualPath(sourcePath, reference.ContainerPath!);
        assetsByLocation.TryGetValue(packageLocation, out OfficeDocumentAsset? packageAsset);
        return packageAsset;
    }

    private static string? BuildEpubResolvedLocation(
        string sourcePath,
        EpubReference reference,
        bool includeFragment) {
        if (!reference.IsValid || reference.Kind == EpubReferenceKind.Data) return null;
        if (reference.Kind == EpubReferenceKind.External) {
            return includeFragment ? reference.ResolvedValue : reference.Target;
        }
        if (string.IsNullOrWhiteSpace(reference.ContainerPath)
            || string.IsNullOrWhiteSpace(reference.ContainerUrlPath)) return null;

        string location = BuildVirtualPath(sourcePath, reference.ContainerUrlPath!);
        string? resolved = includeFragment ? reference.ResolvedValue : reference.Target;
        if (!string.IsNullOrWhiteSpace(resolved) && resolved!.Length > reference.ContainerUrlPath!.Length) {
            location += resolved.Substring(reference.ContainerUrlPath.Length);
        }
        return location;
    }

    private static void AddEpubReferenceDiagnostic(
        string sourcePath,
        EpubChapter chapter,
        EpubReference reference,
        List<OfficeDocumentDiagnostic> diagnostics) {
        string? code = null;
        string? message = null;
        OfficeDocumentDiagnosticCategory category = OfficeDocumentDiagnosticCategory.Parsing;
        var attributes = new Dictionary<string, string>(StringComparer.Ordinal) {
            ["reference"] = reference.Original,
            ["referenceKind"] = reference.Kind.ToString(),
            ["referenceError"] = reference.Error.ToString()
        };

        if (!reference.IsValid && reference.Error != EpubReferenceError.Empty) {
            bool unsafeReference = reference.Error == EpubReferenceError.EscapesContainer
                || reference.Error == EpubReferenceError.FileUrl
                || reference.Error == EpubReferenceError.ControlCharacter
                || reference.Error == EpubReferenceError.InvalidPath;
            code = unsafeReference ? "epub.reference.unsafe" : "epub.reference.invalid";
            category = unsafeReference ? OfficeDocumentDiagnosticCategory.Security : OfficeDocumentDiagnosticCategory.Parsing;
            message = $"Ignored EPUB reference '{reference.Original}' because it could not be resolved safely ({reference.Error}).";
        } else if (reference.IsValid && !reference.IsConforming) {
            code = "epub.reference.non-conforming";
            message = $"Resolved non-conforming EPUB reference '{reference.Original}' within the container.";
            if (!string.IsNullOrWhiteSpace(reference.ResolvedValue)) attributes["resolvedValue"] = reference.ResolvedValue!;
        }
        if (code == null || message == null) return;

        string chapterLocation = BuildVirtualPath(sourcePath, chapter.Path);
        bool duplicate = diagnostics.Any(diagnostic =>
            string.Equals(diagnostic.Code, code, StringComparison.Ordinal)
            && string.Equals(diagnostic.Location?.Path, chapterLocation, StringComparison.Ordinal)
            && diagnostic.Attributes.TryGetValue("reference", out string? existing)
            && string.Equals(existing, reference.Original, StringComparison.Ordinal));
        if (duplicate) return;

        diagnostics.Add(new OfficeDocumentDiagnostic {
            Severity = OfficeDocumentDiagnosticSeverity.Warning,
            Category = category,
            Code = code,
            Message = message,
            Source = "OfficeIMO.Reader.Epub",
            IsRecoverable = true,
            Location = new ReaderLocation { Path = chapterLocation, SourceBlockKind = "chapter-reference" },
            Attributes = attributes
        });
    }
}
